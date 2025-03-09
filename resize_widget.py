import win32gui
import win32con
from tkinter import Toplevel, Label
from pycaw.pycaw import AudioUtilities
import os
import win32process
import psutil

class MuteWidget:
    def __init__(self, hwnd, position, size, process_name, debug_mode=False, app_state=None):
        self.hwnd = hwnd
        self.size = size
        self.process_name = process_name
        self.debug_mode = debug_mode
        self.last_reason = "Unknown"
        self.app_state = app_state
        
        self.window = Toplevel()
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)
        self.window.geometry(f"{size}x{size}+{position[0]}+{position[1]}")
        
        # Create tooltip window
        self.tooltip = Toplevel()
        self.tooltip.withdraw()  # Hide initially
        self.tooltip.overrideredirect(True)
        self.tooltip.attributes('-topmost', True)
        self.tooltip.configure(bg='#2b2b2b')
        self.tooltip_label = Label(self.tooltip, bg='#2b2b2b', fg='white', 
                                 font=('Arial', 8), padx=5, pady=3)
        self.tooltip_label.pack()
        
        # Bind mouse events for tooltip
        self.window.bind('<Enter>', self.show_tooltip)
        self.window.bind('<Leave>', self.hide_tooltip)
        self.window.bind('<Motion>', self.update_tooltip_position)
        
        # Set initial mute state and color
        self.update_mute_state("Initial state")
        
        # Bind click event
        self.window.bind('<Button-1>', self.toggle_mute)

    def show_tooltip(self, event):
        """Show the tooltip"""
        if self.tooltip.state() == 'withdrawn':
            self.update_tooltip_position(event)
            self.tooltip.deiconify()

    def hide_tooltip(self, event):
        """Hide the tooltip"""
        self.tooltip.withdraw()

    def update_tooltip_position(self, event):
        """Update tooltip position near the widget"""
        x = self.window.winfo_rootx() + self.size + 5
        y = self.window.winfo_rooty()
        self.tooltip.geometry(f"+{x}+{y}")

    def exists(self):
        """Check if widget window still exists"""
        return self.window.winfo_exists()

    def update_position(self, x, y):
        """Update widget position"""
        self.window.geometry(f"{self.size}x{self.size}+{x}+{y}")

    def destroy(self):
        """Destroy the widget window"""
        if self.exists():
            self.tooltip.destroy()
            self.window.destroy()

    def update_mute_state(self, reason="Unknown"):
        """Update widget color based on mute state"""
        try:
            # Get exe name from window handle
            _, pid = win32process.GetWindowThreadProcessId(self.hwnd)
            process = psutil.Process(pid)
            exe_name = os.path.basename(process.exe())

            is_muted = False
            is_force_muted = self.app_state and self.app_state.is_force_muted(exe_name)
            
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and os.path.basename(session.Process.exe()) == exe_name:
                    volume = session.SimpleAudioVolume
                    if volume:
                        is_muted = volume.GetMute()
                        if is_force_muted:
                            self.last_reason = "Force Muted"
                        else:
                            self.last_reason = reason
                        break
            
            self.window.configure(bg='#ff6b6b' if is_muted else '#69db7c')
            status = f"{'Muted' if is_muted else 'Unmuted'}\nReason: {self.last_reason}"
            if is_force_muted:
                status += "\n(Force Mute enabled)"
            self.tooltip_label.config(text=status)
            
            if self.debug_mode:
                print(f"Updated mute widget state for {exe_name}: {status}")
        except Exception as e:
            if self.debug_mode:
                print(f"Error updating mute state: {e}")
            self.window.configure(bg='gray')
            self.tooltip_label.config(text="Error getting mute state")

    def toggle_mute(self, event=None):
        """Toggle mute state of the application"""
        print(f"chaning app state")
        try:
            # Get exe name from window handle
            _, pid = win32process.GetWindowThreadProcessId(self.hwnd)
            process = psutil.Process(pid)
            exe_name = os.path.basename(process.exe())

            print(f"for exe_name{exe_name} {self.app_state}")

            # Toggle force mute state
            if self.app_state:
                current_force_mute = self.app_state.is_force_muted(exe_name)
                self.app_state.save_force_mute_app(exe_name, not current_force_mute)
                
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if session.Process and os.path.basename(session.Process.exe()) == exe_name:
                        volume = session.SimpleAudioVolume
                        if volume:
                            volume.SetMute(not current_force_mute, None)
                            self.update_mute_state("Manual toggle via widget")
                            if self.debug_mode:
                                print(f"Toggled force mute for {exe_name}: {'Muted' if not current_force_mute else 'Unmuted'} (Manual toggle)")
                            break
        except Exception as e:
            import traceback
            traceback.print_exc()
            if self.debug_mode:
                print(f"Error toggling mute: {e}")

class ResizeWidgetManager:
    def __init__(self, debug_mode=False, widget_size=10, app_state=None):
        self.widgets = {}
        self.mute_widgets = {}
        self.move_widgets = {}  # Add storage for move widgets
        self.debug_mode = debug_mode
        self.widget_size = widget_size
        self.app_state = app_state
        self.active_window = None
        self.is_resizing = False

    def create_or_update_widgets(self, window_name, hwnd):
        """Create or update resize widgets for a window"""
        try:
            # Check if window still exists and is visible
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                self.remove_widgets_for_hwnd(window_name, hwnd)
                return

            # Only show widgets for active window
            active_hwnd = win32gui.GetForegroundWindow()
            if active_hwnd != hwnd:
                self.remove_widgets_for_hwnd(window_name, hwnd)
                return

            if self.debug_mode:
                print(f"\nUpdating resize widgets for {window_name}")
                print(f"Window handle: {hwnd}")
            
            # Get window rect
            rect = win32gui.GetWindowRect(hwnd)
            x, y, right, bottom = rect
            width = right - x
            height = bottom - y
            
            if self.debug_mode:
                print(f"Window dimensions: {width}x{height} at ({x},{y})")

            widget_key = f"{window_name}_{hwnd}"
            
            # Create new widgets if they don't exist for this window
            if widget_key not in self.widgets:
                if self.debug_mode:
                    print(f"Creating new resize widgets for {widget_key}")
                
                self.widgets[widget_key] = []
                
                # Create widgets for each corner
                corners = [
                    ('nw', x, y),
                    ('ne', right - self.widget_size, y),
                    ('sw', x, bottom - self.widget_size),
                    ('se', right - self.widget_size, bottom - self.widget_size)
                ]
                
                for corner, wx, wy in corners:
                    try:
                        widget = ResizeWidget(
                            hwnd=hwnd,
                            corner=corner,
                            position=(wx, wy),
                            size=self.widget_size,
                            debug_mode=self.debug_mode,
                            manager=self
                        )
                        self.widgets[widget_key].append(widget)
                        
                        # Create mute widget next to NE corner
                        if corner == 'ne':
                            # Create move widget at top center
                            move_x = x + (width - self.widget_size) // 2
                            move_y = y
                            move_widget = MoveWidget(
                                hwnd=hwnd,
                                position=(move_x, move_y),
                                size=self.widget_size,
                                debug_mode=self.debug_mode,
                                manager=self
                            )
                            self.move_widgets[widget_key] = move_widget
                            
                            # Create mute widget next to move widget
                            mute_widget = MuteWidget(
                                hwnd=hwnd,
                                position=(move_x - self.widget_size * 2, move_y),
                                size=self.widget_size,
                                process_name=window_name,
                                debug_mode=self.debug_mode,
                                app_state=self.app_state
                            )
                            self.mute_widgets[widget_key] = mute_widget
                            mute_widget.update_mute_state("Widget created")
                        
                        if self.debug_mode:
                            print(f"Created {corner} widget at ({wx},{wy})")
                    except Exception as e:
                        print(f"Error creating {corner} widget: {e}")
            else:
                # Update existing widgets positions
                if self.debug_mode:
                    print(f"Updating existing widgets for {widget_key}")
                
                corners = [
                    ('nw', x, y),
                    ('ne', right - self.widget_size, y),
                    ('sw', x, bottom - self.widget_size),
                    ('se', right - self.widget_size, bottom - self.widget_size)
                ]
                
                for widget, (corner, wx, wy) in zip(self.widgets[widget_key], corners):
                    try:
                        if widget.exists():
                            widget.update_position(wx, wy)
                            
                            # Update move and mute widget positions if this is the NE corner
                            if corner == 'ne':
                                move_x = x + (width - self.widget_size) // 2
                                move_y = y
                                
                                # Update move widget
                                if widget_key in self.move_widgets:
                                    move_widget = self.move_widgets[widget_key]
                                    if move_widget.exists():
                                        move_widget.update_position(move_x, move_y)
                                    else:
                                        # Recreate move widget if it was destroyed
                                        self.move_widgets[widget_key] = MoveWidget(
                                            hwnd=hwnd,
                                            position=(move_x, move_y),
                                            size=self.widget_size,
                                            debug_mode=self.debug_mode,
                                            manager=self
                                        )
                                
                                # Update mute widget
                                if widget_key in self.mute_widgets:
                                    mute_widget = self.mute_widgets[widget_key]
                                    if mute_widget.exists():
                                        mute_widget.update_position(move_x - self.widget_size * 2, move_y)
                                        mute_widget.update_mute_state("Window position updated")
                                    else:
                                        # Recreate mute widget if it was destroyed
                                        self.mute_widgets[widget_key] = MuteWidget(
                                            hwnd=hwnd,
                                            position=(move_x - self.widget_size * 2, move_y),
                                            size=self.widget_size,
                                            process_name=window_name,
                                            debug_mode=self.debug_mode,
                                            app_state=self.app_state
                                        )
                                        self.mute_widgets[widget_key].update_mute_state("Widget recreated")
                            
                            if self.debug_mode:
                                print(f"Updated {corner} widget to ({wx},{wy})")
                        else:
                            if self.debug_mode:
                                print(f"Widget for {corner} no longer exists, recreating")
                            new_widget = ResizeWidget(
                                hwnd=hwnd,
                                corner=corner,
                                position=(wx, wy),
                                size=self.widget_size,
                                debug_mode=self.debug_mode,
                                manager=self
                            )
                            self.widgets[widget_key][self.widgets[widget_key].index(widget)] = new_widget
                    except Exception as e:
                        print(f"Error updating {corner} widget: {e}")
                    
        except Exception as e:
            print(f"Error in create_or_update_widgets: {e}")
            import traceback
            traceback.print_exc()

    def remove_widgets_for_hwnd(self, window_name, hwnd):
        """Remove widgets for a specific window handle"""
        if self.is_resizing:
            return

        widget_key = f"{window_name}_{hwnd}"
        if widget_key in self.widgets:
            if self.debug_mode:
                print(f"Removing resize widgets for {widget_key}")
            for widget in self.widgets[widget_key]:
                widget.destroy()
            del self.widgets[widget_key]
            
            # Remove move widget
            if widget_key in self.move_widgets:
                self.move_widgets[widget_key].destroy()
                del self.move_widgets[widget_key]
            
            # Remove mute widget
            if widget_key in self.mute_widgets:
                self.mute_widgets[widget_key].destroy()
                del self.mute_widgets[widget_key]

    def remove_widgets(self, window_name):
        """Remove all resize widgets for a window name"""
        for key in list(self.widgets.keys()):
            if key.startswith(f"{window_name}_"):
                if self.debug_mode:
                    print(f"Removing resize widgets for {key}")
                for widget in self.widgets[key]:
                    widget.destroy()
                del self.widgets[key]
                
                # Remove move widget
                if key in self.move_widgets:
                    self.move_widgets[key].destroy()
                    del self.move_widgets[key]
                
                # Remove mute widget
                if key in self.mute_widgets:
                    self.mute_widgets[key].destroy()
                    del self.mute_widgets[key]

    def update_all_widgets(self):
        """Update all existing widgets with new size"""
        for window_name in list(self.widgets.keys()):
            try:
                # Extract hwnd from the key
                hwnd = int(window_name.split('_')[-1])
                window_base_name = window_name.rsplit('_', 1)[0]
                self.create_or_update_widgets(window_base_name, hwnd)
            except Exception as e:
                if self.debug_mode:
                    print(f"Error updating widgets for {window_name}: {e}")

    def cleanup_closed_windows(self):
        """Remove widgets for windows that no longer exist"""
        for window_name in list(self.widgets.keys()):
            try:
                hwnd = int(window_name.split('_')[-1])
                if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                    self.remove_widgets_for_hwnd(window_name.rsplit('_', 1)[0], hwnd)
            except Exception as e:
                if self.debug_mode:
                    print(f"Error cleaning up widgets for {window_name}: {e}")

class ResizeWidget:
    def __init__(self, hwnd, corner, position, size, debug_mode, manager:ResizeWidgetManager):
        self.hwnd = hwnd
        self.corner = corner
        self.size = size
        self.debug_mode = debug_mode
        self.manager = manager
        
        self.window = Toplevel()
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)
        self.window.geometry(f"{size}x{size}+{position[0]}+{position[1]}")
        self.window.configure(bg='gray')
        
        # Store initial positions for drag calculation
        self.window.bind('<Button-1>', self.start_resize)
        self.window.bind('<B1-Motion>', self.do_resize)
        self.window.bind('<ButtonRelease-1>', self.end_resize)

        self.start_x = None
        self.start_y = None

    def exists(self):
        """Check if widget window still exists"""
        return self.window.winfo_exists()

    def update_position(self, x, y):
        """Update widget position"""
        self.window.geometry(f"{self.size}x{self.size}+{x}+{y}")

    def destroy(self):
        """Destroy the widget window"""
        print("destroy widget")
        if self.exists():
            self.window.destroy()

    def start_resize(self, event):
        self.manager.is_resizing = True
        """Start window resize operation"""
        self.start_x = event.x_root
        self.start_y = event.y_root
        self.start_rect = win32gui.GetWindowRect(self.hwnd)
        print(f"starting resize {self.corner}")

    def end_resize(self, event):
        self.manager.is_resizing = False
        """Handle end of resize operation"""
        print("ending resize")
        try:
            if self.manager:
                # Get the window name from the manager's widgets
                window_name = None
                for key, widgets in self.manager.widgets.items():
                    if any(widget is self for widget in widgets):
                        window_name = key.rsplit('_', 1)[0]
                        break
                
                if window_name:
                    # Update all widgets for this window
                    self.manager.create_or_update_widgets(window_name, self.hwnd)
        except Exception as e:
            print(f"Error in end_resize: {e}")

    def do_resize(self, event):
        """Handle window resize operation"""
        try:
            print("do resize")
            if self.start_x is None:
                print("start_x is not set")
                return

            dx = event.x_root - self.start_x
            dy = event.y_root - self.start_y
            x, y, right, bottom = self.start_rect
            
            if self.corner == 'nw':
                new_rect = (x + dx, y + dy, right, bottom)
            elif self.corner == 'ne':
                new_rect = (x, y + dy, right + dx, bottom)
            elif self.corner == 'sw':
                new_rect = (x + dx, y, right, bottom + dy)
            elif self.corner == 'se':
                new_rect = (x, y, right + dx, bottom + dy)
            
            # Apply new window position and size
            win32gui.SetWindowPos(self.hwnd, 0, 
                                new_rect[0], new_rect[1],
                                new_rect[2] - new_rect[0],
                                new_rect[3] - new_rect[1],
                                win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
            
            # Update widget positions during drag
            if self.manager:
                rect = win32gui.GetWindowRect(self.hwnd)
                x, y, right, bottom = rect
                corners = [
                    ('nw', x, y),
                    ('ne', right - self.size, y),
                    ('sw', x, bottom - self.size),
                    ('se', right - self.size, bottom - self.size)
                ]
                
                # Find our window in the manager's widgets
                for key, widgets in self.manager.widgets.items():
                    if any(widget is self for widget in widgets):
                        for widget, (corner, wx, wy) in zip(widgets, corners):
                            if widget is not self:  # Don't update the widget being dragged
                                widget.update_position(wx, wy)
                        break
                                
        except Exception as e:
            print(f"Error resizing window: {e}") 


class MoveWidget:
    def __init__(self, hwnd, position, size, manager: ResizeWidgetManager, debug_mode=False):
        self.hwnd = hwnd
        self.size = size
        self.debug_mode = debug_mode
        self.manager = manager
        
        self.window = Toplevel()
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)
        self.window.geometry(f"{size}x{size}+{position[0]}+{position[1]}")
        self.window.configure(bg='#4b6eaf')  # Use a distinct color
        
        # Create tooltip window
        self.tooltip = Toplevel()
        self.tooltip.withdraw()  # Hide initially
        self.tooltip.overrideredirect(True)
        self.tooltip.attributes('-topmost', True)
        self.tooltip.configure(bg='#2b2b2b')
        self.tooltip_label = Label(self.tooltip, bg='#2b2b2b', fg='white',
                                 text="Click and drag to move window",
                                 font=('Arial', 8), padx=5, pady=3)
        self.tooltip_label.pack()
        
        # Bind mouse events
        self.window.bind('<Button-1>', self.start_move)
        self.window.bind('<B1-Motion>', self.do_move)
        self.window.bind('<Enter>', self.show_tooltip)
        self.window.bind('<Leave>', self.hide_tooltip)
        self.window.bind('<Motion>', self.update_tooltip_position)
        
        self.start_x = None
        self.start_y = None
        self.start_rect = None

    def show_tooltip(self, event):
        """Show the tooltip"""
        if self.tooltip.state() == 'withdrawn':
            self.update_tooltip_position(event)
            self.tooltip.deiconify()

    def hide_tooltip(self, event):
        """Hide the tooltip"""
        self.tooltip.withdraw()

    def update_tooltip_position(self, event):
        """Update tooltip position near the widget"""
        x = self.window.winfo_rootx() + self.size + 5
        y = self.window.winfo_rooty()
        self.tooltip.geometry(f"+{x}+{y}")

    def exists(self):
        """Check if widget window still exists"""
        return self.window.winfo_exists()

    def update_position(self, x, y):
        """Update widget position"""
        self.window.geometry(f"{self.size}x{self.size}+{x}+{y}")

    def destroy(self):
        """Destroy the widget window"""
        if self.exists():
            self.tooltip.destroy()
            self.window.destroy()

    def start_move(self, event):
        """Start window move operation"""
        self.start_x = event.x_root
        self.start_y = event.y_root
        self.start_rect = win32gui.GetWindowRect(self.hwnd)
        self.manager.is_resizing = True
        if self.debug_mode:
            print(f"Starting move from ({self.start_x}, {self.start_y})")

    def do_move(self, event):
        """Handle window move operation"""
        try:
            if self.start_x is None:
                return
            self.manager.is_resizing = False

            dx = event.x_root - self.start_x
            dy = event.y_root - self.start_y
            x, y, right, bottom = self.start_rect
            width = right - x
            height = bottom - y
            
            # Move window to new position
            win32gui.SetWindowPos(self.hwnd, 0,
                                x + dx, y + dy, width, height,
                                win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
            
            if self.debug_mode:
                print(f"Moving window by ({dx}, {dy})")
                
        except Exception as e:
            if self.debug_mode:
                print(f"Error moving window: {e}")
