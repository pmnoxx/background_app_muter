import win32gui
import win32con
from tkinter import Toplevel

class ResizeWidgetManager:
    def __init__(self, debug_mode=False, widget_size=10):
        self.widgets = {}
        self.debug_mode = debug_mode
        self.widget_size = widget_size
        self.active_window = None

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
                            debug_mode=self.debug_mode
                        )
                        self.widgets[widget_key].append(widget)
                        
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
                                debug_mode=self.debug_mode
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
        widget_key = f"{window_name}_{hwnd}"
        if widget_key in self.widgets:
            if self.debug_mode:
                print(f"Removing resize widgets for {widget_key}")
            for widget in self.widgets[widget_key]:
                widget.destroy()
            del self.widgets[widget_key]

    def remove_widgets(self, window_name):
        """Remove all resize widgets for a window name"""
        for key in list(self.widgets.keys()):
            if key.startswith(f"{window_name}_"):
                if self.debug_mode:
                    print(f"Removing resize widgets for {key}")
                for widget in self.widgets[key]:
                    widget.destroy()
                del self.widgets[key]

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
    def __init__(self, hwnd, corner, position, size, debug_mode=False):
        self.hwnd = hwnd
        self.corner = corner
        self.size = size
        self.debug_mode = debug_mode
        
        self.window = Toplevel()
        self.window.overrideredirect(True)
        self.window.attributes('-topmost', True)
        self.window.geometry(f"{size}x{size}+{position[0]}+{position[1]}")
        self.window.configure(bg='gray')
        
        # Store initial positions for drag calculation
        self.window.bind('<Button-1>', self.start_resize)
        self.window.bind('<B1-Motion>', self.do_resize)

    def exists(self):
        """Check if widget window still exists"""
        return self.window.winfo_exists()

    def update_position(self, x, y):
        """Update widget position"""
        self.window.geometry(f"{self.size}x{self.size}+{x}+{y}")

    def destroy(self):
        """Destroy the widget window"""
        if self.exists():
            self.window.destroy()

    def start_resize(self, event):
        """Start window resize operation"""
        self.start_x = event.x_root
        self.start_y = event.y_root
        self.start_rect = win32gui.GetWindowRect(self.hwnd)

    def do_resize(self, event):
        """Handle window resize operation"""
        try:
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
                                
        except Exception as e:
            print(f"Error resizing window: {e}") 