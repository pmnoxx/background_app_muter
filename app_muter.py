import os
import sys
import psutil
import toml
import win32gui
import win32process
import win32con
import win32ui
import ctypes
import win32api
import pyuac
import time
from tkinter import Tk, Listbox, Button, Label, END, Checkbutton, IntVar, Scale, Toplevel, Frame, Entry, StringVar, OptionMenu, LabelFrame, messagebox

from pycaw.pycaw import AudioUtilities, IAudioMeterInformation

def read_config(filename):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_path = os.path.join(script_dir, filename)

        with open(full_path, "r") as toml_file:
            data = toml.load(toml_file)
            return data
    except FileNotFoundError:
        print(f"File '{filename}' not found.")
        return {}

class AppState:
    def __init__(self):
        # Add version constant
        self.VERSION = "1.0.0"
        
        # Create main window first
        self.root = Tk()
        
        # Load configuration
        self.config = read_config("config.toml")
        self.runtime = read_config("runtime.toml")
        
        self.DEFAULT_EXCEPTION_LIST = self.config.get("DEFAULT_EXCEPTIONS", ["chrome.exe", "firefox.exe", "msedge.exe"])
        self.MUTE_GROUPS = self.config.get("MUTE_GROUPS", [])
        
        # Get settings
        runtime_settings = self.runtime.get("SETTINGS", {})
        window = self.config.get("WINDOW", {})
        
        # Store theme colors
        self.theme = {
            'bg': window.get('background_color', '#2b2b2b'),
            'fg': window.get('foreground_color', 'white'),
            'button': window.get('button_color', '#3c3f41'),
            'active': window.get('active_color', '#4b6eaf')
        }
        
        # Store window settings
        self.window = {
            'min_width': window.get('min_width', 300),
            'min_height': window.get('min_height', 500),
            'width': window.get('width', 400),
            'height': window.get('height', 600)
        }

        # Initialize state variables
        self.exceptions_list = self.runtime.get("CURRENT_EXCEPTIONS", [])
        self.to_unmute = []
        self.last_foreground_app_pid = None
        self.zero_cnt = 0

        # Create Tkinter variables with defaults from runtime config
        self.mute_last_app = IntVar(value=runtime_settings.get("mute_last_app", 0))
        self.force_mute_fg_var = IntVar(value=runtime_settings.get("force_mute_fg", 0))
        self.force_mute_bg_var = IntVar(value=runtime_settings.get("force_mute_bg", 0))
        self.lock_var = IntVar(value=runtime_settings.get("lock", 0))
        self.mute_foreground_when_background = IntVar(value=runtime_settings.get("mute_foreground_when_background", 0))

        # Get window geometry from runtime settings
        window_settings = self.runtime.get("WINDOW_STATE", {})
        self.window_state = {
            'geometry': window_settings.get('geometry', None),
            'maximized': window_settings.get('maximized', False)
        }

        # Load exceptions if none exist
        if not self.exceptions_list:
            self.exceptions_list = self.DEFAULT_EXCEPTION_LIST.copy()
            self.save_exceptions()

        # Load app-specific volumes
        self.app_volumes = self.config.get("APP_VOLUMES", {})

        # Add pid_match_apps setting
        self.pid_match_apps = self.config.get("PID_MATCH_APPS", [])

        # Add hide_titlebar_apps setting
        self.hide_titlebar_apps = self.config.get("HIDE_TITLEBAR_APPS", [])

        # Add maximize_apps setting
        self.maximize_apps = self.config.get("MAXIMIZE_APPS", [])

        # Define resolution presets by aspect ratio
        self.RESOLUTION_PRESETS = {
            "8k": {"width": 7680, "height": 4320},
            # 16:9 options
            "16:9 720p": {"width": 1280, "height": 720},
            "16:9 1080p": {"width": 1920, "height": 1080},
            "16:9 1440p": {"width": 2560, "height": 1440},
            "16:9 4K": {"width": 3840, "height": 2160},
            "16:9 Fit Screen": {"width": "fit_16_9", "height": "fit_16_9"},
            
            # 19.5:9 options
            "19.5:9 720p": {"width": 1560, "height": 720},
            "19.5:9 1080p": {"width": 2340, "height": 1080},
            "19.5:9 1440p": {"width": 3120, "height": 1440},
            "19.5:9 Fit Screen": {"width": "fit_19_5_9", "height": "fit_19_5_9"},
            
            # 21:9 options
            "21:9 720p": {"width": 1720, "height": 720},
            "21:9 1080p": {"width": 2560, "height": 1080},
            "21:9 1440p": {"width": 3440, "height": 1440},
            "21:9 Fit Screen": {"width": "fit_21_9", "height": "fit_21_9"},

            # 24:9 options
            "24:9 720p": {"width": 1920, "height": 720},
            "24:9 1080p": {"width": 2880, "height": 1080},
            "24:9 1440p": {"width": 3840, "height": 1440},
            "24:9 Fit Screen": {"width": "fit_24_9", "height": "fit_24_9"},
            
            # 32:9 options
            "32:9 720p": {"width": 2560, "height": 720},
            "32:9 1080p": {"width": 3840, "height": 1080},
            "32:9 Fit Screen": {"width": "fit_32_9", "height": "fit_32_9"},
        }
        
        # Add custom resolution settings
        self.custom_resolution_apps = self.config.get("CUSTOM_RESOLUTION_APPS", {})

        # Define window placement options
        self.WINDOW_PLACEMENTS = {
            "No Change": "no_change",
            "Centered": "center",
            "Top": "top",
            "Bottom": "bottom",
            "Left": "left",
            "Right": "right",
            "Top Left": "top_left",
            "Top Right": "top_right",
            "Bottom Left": "bottom_left",
            "Bottom Right": "bottom_right"
        }
        
        # Add window placement settings
        self.window_placements = self.config.get("WINDOW_PLACEMENTS", {})

        # Define border style options
        self.BORDER_STYLES = {
            "No Change": "no_change",
            "Normal": "normal",
            "Thin": "thin",
            "None": "none",
            "Dialog": "dialog",
            "Tool": "tool"
        }
        
        # Add border style settings
        self.border_styles = self.config.get("BORDER_STYLES", {})

        # Add options settings
        self.options = self.config.get("OPTIONS", {
            "window_check_interval": 1000,  # milliseconds
            "volume_check_interval": 100,   # milliseconds
            "list_update_interval": 100,    # milliseconds
            "debug_mode": False,
        })

        # Add startup delay settings
        self.startup_delays = self.config.get("STARTUP_DELAYS", {})
        self.app_start_times = {}  # Track when apps were first seen

        # Start combined window state checks
        self.root.after(self.options["window_check_interval"], self.check_all_window_states)

        self.volume_control = None  # Add reference to volume control window

    def setup_main_window(self):
        """Initialize main window settings"""
        self.root.title(f"App Muter v{self.VERSION}")
        self.root.configure(bg=self.theme['bg'])
        
        # Set unique app ID and icon for Windows taskbar
        try:
            import ctypes
            myappid = 'mycompany.appmuter.subversion.1'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            
            # Load icon for both window and taskbar
            icon_path = self.ensure_app_icon()
            if icon_path:
                from PIL import Image, ImageTk
                icon = Image.open(icon_path)
                # Convert to PhotoImage for window icon
                photo = ImageTk.PhotoImage(icon)
                self.root.iconphoto(True, photo)
                # Convert to ICO for taskbar
                icon_ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.ico")
                if not os.path.exists(icon_ico):
                    icon.save(icon_ico, format='ICO', sizes=[(256, 256)])
                self.root.iconbitmap(icon_ico)
        except Exception as e:
            print(f"Error setting app ID/icon: {e}")
            
        # Center window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - self.window['width']/2)
        center_y = int(screen_height/2 - self.window['height']/2)
        self.root.geometry(f'{self.window["width"]}x{self.window["height"]}+{center_x}+{center_y}')
        
        # Bind window events
        self.root.bind("<Configure>", lambda e: self.save_window_state())

    def ensure_app_icon(self):
        """Generate and save app icon if it doesn't exist"""
        try:
            from PIL import Image, ImageDraw
            
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app_icon.png")
            
            # Only generate if icon doesn't exist
            if not os.path.exists(icon_path):
                # Create base image with transparency
                size = 256
                image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
                draw = ImageDraw.Draw(image)
                
                # Define colors
                bg_color = '#2b2b2b'
                fg_color = '#4b6eaf'
                
                # Draw main circle
                padding = size * 0.1
                circle_bbox = [padding, padding, size-padding, size-padding]
                draw.ellipse(circle_bbox, fill=bg_color)
                
                # Draw speaker symbol
                speaker_size = size * 0.4
                speaker_x = size * 0.3
                speaker_y = size * 0.3
                
                # Speaker box
                box_points = [
                    (speaker_x, speaker_y),
                    (speaker_x + speaker_size*0.4, speaker_y),
                    (speaker_x + speaker_size*0.8, speaker_y - speaker_size*0.2),
                    (speaker_x + speaker_size*0.8, speaker_y + speaker_size*1.2),
                    (speaker_x + speaker_size*0.4, speaker_y + speaker_size),
                    (speaker_x, speaker_y + speaker_size),
                ]
                draw.polygon(box_points, fill=fg_color)
                
                # Sound waves
                wave_x = speaker_x + speaker_size*0.9
                wave_y = speaker_y + speaker_size*0.5
                wave_radius = speaker_size * 0.2
                
                for i in range(3):
                    draw.arc([wave_x + i*wave_radius, wave_y - wave_radius,
                             wave_x + wave_radius + i*wave_radius, wave_y + wave_radius],
                            -60, 60, fill=fg_color, width=int(size*0.02))
                
                # Draw mute line
                line_width = int(size*0.04)
                draw.line([(size*0.2, size*0.8), (size*0.8, size*0.2)], 
                         fill='#ff6b6b', width=line_width)
                
                # Save as PNG
                image.save(icon_path, format='PNG')
            
            return icon_path
            
        except Exception as e:
            print(f"Error generating icon: {e}")
            return None

    def save_exceptions(self):
        """Save exceptions to runtime file"""
        self.runtime["CURRENT_EXCEPTIONS"] = self.exceptions_list
        self.save_runtime()

    def save_runtime(self):
        """Save runtime settings to file"""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            runtime_path = os.path.join(script_dir, "runtime.toml")
            
            # Update runtime settings
            self.runtime["SETTINGS"] = {
                "mute_last_app": self.mute_last_app.get(),
                "force_mute_fg": self.force_mute_fg_var.get(),
                "force_mute_bg": self.force_mute_bg_var.get(),
                "mute_foreground_when_background": self.mute_foreground_when_background.get(),
                "lock": self.lock_var.get(),
            }
            
            with open(runtime_path, "w") as f:
                toml.dump(self.runtime, f)
        except Exception as e:
            print(f"Error saving runtime config: {e}")

    def add_exception(self, app_name):
        if app_name and app_name not in self.exceptions_list:
            self.exceptions_list.append(app_name)
            self.to_unmute.append(app_name)
            self.save_exceptions()

    def remove_exception(self, app_name):
        if app_name and app_name in self.exceptions_list:
            self.exceptions_list.remove(app_name)
            self.save_exceptions()

    def update_params(self):
        """Update runtime parameters"""
        self.save_runtime()

    def save_window_state(self):
        """Save current window position and size"""
        try:
            # Get current window state
            if self.root.state() == 'zoomed':  # Window is maximized
                self.window_state['maximized'] = True
                # Store the last known normal geometry
                if hasattr(self.root, 'last_normal_geometry'):
                    self.window_state['geometry'] = self.root.last_normal_geometry
            else:
                self.window_state['maximized'] = False
                self.window_state['geometry'] = self.root.geometry()
                # Store current geometry for when window is unmaximized
                self.root.last_normal_geometry = self.root.geometry()

            # Update runtime settings
            self.runtime["WINDOW_STATE"] = self.window_state
            self.save_runtime()
        except Exception as e:
            print(f"Error saving window state: {e}")

    def restore_window_state(self):
        """Restore saved window position and size"""
        if self.window_state['geometry']:
            self.root.geometry(self.window_state['geometry'])
        if self.window_state['maximized']:
            self.root.state('zoomed')

    def save_config(self):
        """Save configuration settings to file"""
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(script_dir, "config.toml")
            
            with open(config_path, "w") as f:
                toml.dump(self.config, f)
        except Exception as e:
            print(f"Error saving config: {e}")

    def save_app_volume(self, app_name, volume):
        """Save volume setting for specific app"""
        self.app_volumes[app_name] = volume
        self.config["APP_VOLUMES"] = self.app_volumes
        self.save_config()

    def get_app_volume(self, app_name):
        """Get volume setting for specific app"""
        return self.app_volumes.get(app_name, 100)

    def save_pid_match_app(self, app_name, should_match_pid):
        """Save PID matching setting for specific app"""
        if should_match_pid and app_name not in self.pid_match_apps:
            self.pid_match_apps.append(app_name)
        elif not should_match_pid and app_name in self.pid_match_apps:
            self.pid_match_apps.remove(app_name)
        
        self.config["PID_MATCH_APPS"] = self.pid_match_apps
        self.save_config()

    def save_hide_titlebar_app(self, app_name, should_hide):
        """Save hide titlebar setting for specific app"""
        if should_hide and app_name not in self.hide_titlebar_apps:
            self.hide_titlebar_apps.append(app_name)
        elif not should_hide and app_name in self.hide_titlebar_apps:
            self.hide_titlebar_apps.remove(app_name)
            # Restore title bars when disabling the option
            self.restore_title_bars(app_name)
        
        self.config["HIDE_TITLEBAR_APPS"] = self.hide_titlebar_apps
        self.save_config()

    def restore_title_bars(self, app_name):
        """Restore title bars for all windows of given app"""
        try:
            def enum_windows_callback(hwnd, _):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(pid)
                    process_name = os.path.basename(process.exe())
                    
                    if process_name == app_name:
                        # Get current window style
                        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                        
                        # Add title bar back
                        new_style = style | win32con.WS_CAPTION
                        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, new_style)
                        # Force window to redraw
                        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                            win32con.SWP_NOMOVE | 
                                            win32con.SWP_NOSIZE | 
                                            win32con.SWP_NOZORDER |
                                            win32con.SWP_NOACTIVATE |
                                            win32con.SWP_FRAMECHANGED)
                except:
                    pass  # Ignore errors for inaccessible windows
                return True

            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            print(f"Error restoring title bars: {e}")

    def check_all_window_states(self):
        """Check and manage all window states (title bars, borders, sizes)"""
        try:
            def enum_windows_callback(hwnd, _):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    process = psutil.Process(pid)
                    process_name = os.path.basename(process.exe())
                    
                    # Track first time we see this app
                    current_time = time.time()
                    if process_name not in self.app_start_times:
                        self.app_start_times[process_name] = current_time
                    
                    # Check if we need to wait before resizing
                    startup_delay = self.startup_delays.get(process_name, 0)
                    if startup_delay > 0:
                        time_since_start = current_time - self.app_start_times[process_name]
                        if time_since_start < startup_delay:
                            if self.options["debug_mode"]:
                                print(f"Waiting {startup_delay - time_since_start:.1f}s before resizing {process_name}")
                            return True
                    
                    # Track if we need to update window
                    needs_update = False
                    needs_style_update = False
                    update_flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE
                    x = y = width = height = 0
                    
                    # Get current style once
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    original_style = style
                    
                    # Handle title bars and borders
                    if process_name in self.hide_titlebar_apps:
                        style &= ~win32con.WS_CAPTION
                        needs_style_update = True
                        
                        # Apply border style if set
                        border_style = self.border_styles.get(process_name, "no_change")
                        if border_style != "no_change":
                            style &= ~(win32con.WS_BORDER | win32con.WS_THICKFRAME | win32con.WS_DLGFRAME)
                            
                            if border_style == "normal":
                                style |= (win32con.WS_THICKFRAME | win32con.WS_BORDER)
                            elif border_style == "thin":
                                style |= win32con.WS_BORDER
                            elif border_style == "dialog":
                                style |= win32con.WS_DLGFRAME
                            elif border_style == "tool":
                                style |= win32con.WS_BORDER
                                style &= ~(win32con.WS_MAXIMIZEBOX | win32con.WS_MINIMIZEBOX)
                    
                    # Apply style changes if needed
                    if needs_style_update and style != original_style:
                        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
                        needs_update = True
                        update_flags |= win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_FRAMECHANGED
                    
                    # Handle custom resolutions
                    if process_name in self.custom_resolution_apps:
                        if win32gui.IsWindowVisible(hwnd) and not win32gui.IsIconic(hwnd):
                            placement = win32gui.GetWindowPlacement(hwnd)
                            if placement[1] != win32con.SW_SHOWMAXIMIZED:
                                settings = self.custom_resolution_apps[process_name]
                                
                                # Debug window info
                                print(f"\nWindow debug for {process_name}:")
                                print(f"  Window handle: {hwnd}")
                                
                                # Ensure process and window are DPI aware
                                try:
                                    user32 = ctypes.windll.user32
                                    process_handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, pid)
                                    user32.SetProcessDpiAwarenessContext(process_handle, -4)
                                    win32gui.SetWindowDisplayAffinity(hwnd, 0)
                                except:
                                    user32.SetProcessDPIAware()
                                
                                # Get screen dimensions and calculate size/position
                                monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
                                monitor_info = win32api.GetMonitorInfo(monitor)
                                work_area = monitor_info['Work']
                                screen_width = work_area[2] - work_area[0]
                                screen_height = work_area[3] - work_area[1]
                                
                                # Calculate target dimensions
                                target_width = settings["width"]
                                target_height = settings["height"]
                                
                                if isinstance(target_width, str) and target_width.startswith("fit_"):
                                    aspect_ratio = target_width.split("_", 1)[1]
                                    if aspect_ratio == "16_9":
                                        ratio = 16/9
                                    elif aspect_ratio == "19_5_9":
                                        ratio = 19.5/9
                                    elif aspect_ratio == "21_9":
                                        ratio = 21/9
                                    elif aspect_ratio == "24_9":
                                        ratio = 24/9
                                    elif aspect_ratio == "32_9":
                                        ratio = 32/9
                                    
                                    # Calculate dimensions that fit the screen while maintaining aspect ratio
                                    # Ensure we're using DPI-aware dimensions
                                    dpi_scale = user32.GetDpiForWindow(hwnd) / 96.0
                                    scaled_width = int(screen_width / dpi_scale)
                                    scaled_height = int(screen_height / dpi_scale)
                                    
                                    if (scaled_width/scaled_height) > ratio:
                                        # Screen is wider than target ratio, fit to height
                                        target_height = scaled_height
                                        target_width = int(scaled_height * ratio)
                                    else:
                                        # Screen is taller than target ratio, fit to width
                                        target_width = scaled_width
                                        target_height = int(scaled_width / ratio)
                                    
                                    # Scale back to actual pixels
                                    target_width = int(target_width * dpi_scale)
                                    target_height = int(target_height * dpi_scale)
                                
                                # Calculate position
                                placement = self.window_placements.get(process_name, "center")
                                new_x, new_y = self.get_window_position(placement, screen_width, screen_height, 
                                                                      target_width, target_height)
                                
                                if new_x is None or new_y is None:
                                    rect = win32gui.GetWindowRect(hwnd)
                                    new_x, new_y = rect[0], rect[1]
                                
                                # Update position and size
                                x, y = new_x, new_y
                                width, height = target_width, target_height
                                needs_update = True
                                update_flags &= ~(win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                                
                                print(f"  Target size: {width}x{height}")
                                print(f"  Position: {x},{y}")
                    
                    # Apply all window updates at once
                    if needs_update:
                        try:
                            user32.SetWindowPos(hwnd, 0, x, y, width, height, update_flags)
                            
                            # Verify size if we changed it
                            if width > 0 and height > 0:
                                time.sleep(0.1)
                                new_rect = win32gui.GetWindowRect(hwnd)
                                new_width = new_rect[2] - new_rect[0]
                                new_height = new_rect[3] - new_rect[1]
                                
                                if new_width != width or new_height != height:
                                    print(f"  Size mismatch, retrying... ({width}x{height} != {new_width}x{new_height})")
                                    user32.SetWindowPos(hwnd, 0, x, y, width, height, update_flags)
                        except Exception as e:
                            print(f"  Error updating window: {e}")
                    
                except Exception as e:
                    print(f"Window callback error: {e}")
                return True

            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            print(f"Error checking window states: {e}")
        
        # Schedule next check
        self.root.after(self.options["window_check_interval"], self.check_all_window_states)

    def save_custom_resolution(self, app_name, enabled, preset=None):
        """Save custom resolution setting for specific app"""
        if enabled and preset in self.RESOLUTION_PRESETS:
            self.custom_resolution_apps[app_name] = self.RESOLUTION_PRESETS[preset]
        elif app_name in self.custom_resolution_apps:
            del self.custom_resolution_apps[app_name]
        
        self.config["CUSTOM_RESOLUTION_APPS"] = self.custom_resolution_apps
        self.save_config()

    def save_window_placement(self, app_name, placement):
        """Save window placement setting for specific app"""
        if placement in self.WINDOW_PLACEMENTS.values():
            self.window_placements[app_name] = placement
        elif app_name in self.window_placements:
            del self.window_placements[app_name]
        
        self.config["WINDOW_PLACEMENTS"] = self.window_placements
        self.save_config()

    def save_border_style(self, app_name, style):
        """Save border style setting for specific app"""
        if style in self.BORDER_STYLES.values():
            self.border_styles[app_name] = style
        elif app_name in self.border_styles:
            del self.border_styles[app_name]
        
        self.config["BORDER_STYLES"] = self.border_styles
        self.save_config()

    def apply_window_style(self, hwnd, style_name):
        """Apply window border style"""
        if style_name == "no_change":
            return
            
        # Get current style
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        
        # Store caption state
        has_caption = style & win32con.WS_CAPTION
        
        # Clear existing border styles but preserve other styles
        style &= ~(win32con.WS_BORDER | win32con.WS_THICKFRAME | win32con.WS_DLGFRAME)
        
        # Restore caption if it was present
        if has_caption:
            style |= win32con.WS_CAPTION
        
        # Apply new style
        if style_name == "normal":
            style |= win32con.WS_OVERLAPPEDWINDOW
        elif style_name == "thin":
            style |= win32con.WS_BORDER
        elif style_name == "none":
            pass  # No border
        elif style_name == "dialog":
            style |= win32con.WS_DLGFRAME
        elif style_name == "tool":
            style |= win32con.WS_BORDER
            style &= ~(win32con.WS_MAXIMIZEBOX | win32con.WS_MINIMIZEBOX)
        
        # Apply style
        win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
        win32gui.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                            win32con.SWP_NOMOVE | 
                            win32con.SWP_NOSIZE | 
                            win32con.SWP_NOZORDER |
                            win32con.SWP_NOACTIVATE |
                            win32con.SWP_FRAMECHANGED)

    def get_window_position(self, placement, screen_width, screen_height, window_width, window_height):
        """Calculate window position based on placement setting"""
        if placement == "no_change":
            return None, None
        elif placement == "center":
            return (screen_width - window_width) // 2, (screen_height - window_height) // 2
        elif placement == "top":
            return (screen_width - window_width) // 2, 0
        elif placement == "bottom":
            return (screen_width - window_width) // 2, screen_height - window_height
        elif placement == "left":
            return 0, (screen_height - window_height) // 2
        elif placement == "right":
            return screen_width - window_width, (screen_height - window_height) // 2
        elif placement == "top_left":
            return 0, 0
        elif placement == "top_right":
            return screen_width - window_width, 0
        elif placement == "bottom_left":
            return 0, screen_height - window_height
        elif placement == "bottom_right":
            return screen_width - window_width, screen_height - window_height
        return None, None

    def save_options(self):
        """Save options to config file"""
        self.config["OPTIONS"] = self.options
        self.save_config()

class VolumeControlWindow:
    def __init__(self, parent, app_state):
        self.window = Toplevel(parent)
        self.window.title("App Volume Control")
        self.window.configure(bg=app_state.theme['bg'])
        
        # Search frame
        search_frame = Frame(self.window, bg=app_state.theme['bg'])
        search_frame.pack(fill='x', padx=5, pady=5)
        
        Label(search_frame, text="Search:", 
              bg=app_state.theme['bg'], 
              fg=app_state.theme['fg']).pack(side='left')
        
        self.search_var = StringVar()
        self.search_var.trace('w', self.filter_apps)
        Entry(search_frame, textvariable=self.search_var,
              bg=app_state.theme['button'],
              fg=app_state.theme['fg']).pack(side='left', fill='x', expand=True)
        
        # Apps frame
        self.apps_frame = Frame(self.window, bg=app_state.theme['bg'])
        self.apps_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.volume_vars = {}
        self.pid_match_vars = {}
        self.mute_labels = {}
        self.volume_labels = {}
        self.hide_titlebar_vars = {}  # Add hide titlebar vars
        self.maximize_vars = {}  # Add maximize vars
        self.resolution_vars = {}
        self.preset_vars = {}  # Store resolution preset StringVars
        self.placement_vars = {}  # Store placement StringVars
        self.border_vars = {}  # Add border style vars
        self.mute_vars = {}  # Add dictionary to store mute variables
        self.update_app_list()
        
        # Start periodic status updates
        self.update_mute_status()
        
    def filter_apps(self, *args):
        search_text = self.search_var.get().lower()
        for widget in self.apps_frame.winfo_children():
            widget.destroy()
        
        self.update_app_list(search_text)
        
    def update_app_list(self, filter_text=""):
        sessions = AudioUtilities.GetAllSessions()
        
        for session in sessions:
            if not session.Process:
                continue
                
            try:
                process = psutil.Process(session.ProcessId)
                app_name = os.path.basename(process.exe())
                
                if filter_text and filter_text not in app_name.lower():
                    continue
                
                frame = Frame(self.apps_frame, bg=app_state.theme['bg'])
                frame.pack(fill='x', padx=5, pady=2)
                
                # App name label
                Label(frame, text=app_name, 
                      bg=app_state.theme['bg'],
                      fg=app_state.theme['fg'],
                      width=20, anchor='w').pack(side='left')
                
                # Status frame to hold mute and volume labels
                status_frame = Frame(frame, bg=app_state.theme['bg'])
                status_frame.pack(side='left', padx=5)
                
                # Add mute checkbox with stored variable
                mute_var = IntVar(value=0)
                self.mute_vars[app_name] = mute_var  # Store the mute variable
                Checkbutton(status_frame, text="Mute",
                           variable=mute_var,
                           bg=app_state.theme['bg'],
                           fg=app_state.theme['fg'],
                           selectcolor=app_state.theme['button'],
                           activebackground=app_state.theme['bg'],
                           command=lambda a=app_name: self.on_mute_change(a)).pack(side='left', padx=5)
                
                # Mute status label
                mute_label = Label(status_frame, text="", width=8,
                                 bg=app_state.theme['bg'],
                                 fg=app_state.theme['fg'])
                mute_label.pack(side='left')
                self.mute_labels[app_name] = mute_label
                
                # Volume level label
                volume_label = Label(status_frame, text="", width=6,
                                   bg=app_state.theme['bg'],
                                   fg=app_state.theme['fg'])
                volume_label.pack(side='left')
                self.volume_labels[app_name] = volume_label
                
                # Volume slider
                volume_var = IntVar(value=app_state.get_app_volume(app_name))
                self.volume_vars[app_name] = volume_var
                
                scale = Scale(frame, from_=0, to=100,
                            orient='horizontal',
                            variable=volume_var,
                            bg=app_state.theme['bg'],
                            fg=app_state.theme['fg'],
                            troughcolor=app_state.theme['button'],
                            activebackground=app_state.theme['active'],
                            command=lambda v, a=app_name: self.on_volume_change(a, v))
                scale.pack(side='left', fill='x', expand=True)
                
                # PID match checkbox
                pid_match_var = IntVar(value=1 if app_name in app_state.pid_match_apps else 0)
                self.pid_match_vars[app_name] = pid_match_var
                
                Checkbutton(frame, text="Match PID",
                           variable=pid_match_var,
                           bg=app_state.theme['bg'],
                           fg=app_state.theme['fg'],
                           selectcolor=app_state.theme['button'],
                           activebackground=app_state.theme['bg'],
                           command=lambda a=app_name: self.on_pid_match_change(a)).pack(side='left', padx=5)

                # Hide titlebar checkbox
                hide_titlebar_var = IntVar(value=1 if app_name in app_state.hide_titlebar_apps else 0)
                self.hide_titlebar_vars[app_name] = hide_titlebar_var
                
                Checkbutton(frame, text="Hide Title",
                           variable=hide_titlebar_var,
                           bg=app_state.theme['bg'],
                           fg=app_state.theme['fg'],
                           selectcolor=app_state.theme['button'],
                           activebackground=app_state.theme['bg'],
                           command=lambda a=app_name: self.on_hide_titlebar_change(a)).pack(side='left', padx=5)

                # Maximize checkbox
                maximize_var = IntVar(value=1 if app_name in app_state.maximize_apps else 0)
                self.maximize_vars[app_name] = maximize_var
                
                Checkbutton(frame, text="Maximize",
                           variable=maximize_var,
                           bg=app_state.theme['bg'],
                           fg=app_state.theme['fg'],
                           selectcolor=app_state.theme['button'],
                           activebackground=app_state.theme['bg'],
                           command=lambda a=app_name: self.on_maximize_change(a)).pack(side='left', padx=5)
                
                # Custom resolution frame
                resolution_frame = Frame(frame, bg=app_state.theme['bg'])
                resolution_frame.pack(side='left', padx=5)
                
                # Resolution checkbox and dropdown
                has_custom = app_name in app_state.custom_resolution_apps
                resolution_var = IntVar(value=1 if has_custom else 0)
                self.resolution_vars[app_name] = resolution_var
                
                # Get current preset based on saved dimensions
                current_dims = app_state.custom_resolution_apps.get(app_name, {})
                current_preset = "1080p"  # default
                for preset, dims in app_state.RESOLUTION_PRESETS.items():
                    if dims == current_dims:
                        current_preset = preset
                
                preset_var = StringVar(value=current_preset)
                self.preset_vars[app_name] = preset_var
                
                Checkbutton(resolution_frame, text="Resolution:",
                           variable=resolution_var,
                           bg=app_state.theme['bg'],
                           fg=app_state.theme['fg'],
                           selectcolor=app_state.theme['button'],
                           activebackground=app_state.theme['bg'],
                           command=lambda a=app_name: self.on_resolution_change(a)).pack(side='left')
                
                # Resolution preset dropdown
                resolution_menu = OptionMenu(resolution_frame, preset_var, 
                                          "8k", 
                                          "16:9 720p", "16:9 1080p", "16:9 1440p", "16:9 4K", "16:9 Fit Screen",
                                          "19.5:9 720p", "19.5:9 1080p", "19.5:9 1440p", "19.5:9 Fit Screen",
                                          "21:9 720p", "21:9 1080p", "21:9 1440p", "21:9 Fit Screen",
                                          "24:9 720p", "24:9 1080p", "24:9 1440p", "24:9 Fit Screen",
                                          "32:9 720p", "32:9 1080p", "32:9 Fit Screen",
                                          command=lambda *args, a=app_name: self.on_resolution_change(a))
                resolution_menu.config(bg=app_state.theme['button'],
                                    fg=app_state.theme['fg'],
                                    activebackground=app_state.theme['active'])
                resolution_menu["menu"].config(bg=app_state.theme['button'],
                                            fg=app_state.theme['fg'])
                resolution_menu.pack(side='left', padx=2)
                
                # Window placement dropdown
                placement_frame = Frame(frame, bg=app_state.theme['bg'])
                placement_frame.pack(side='left', padx=5)
                
                Label(placement_frame, text="Position:",
                      bg=app_state.theme['bg'],
                      fg=app_state.theme['fg']).pack(side='left')
                
                current_placement = app_state.window_placements.get(app_name, "center")
                placement_var = StringVar(value=[k for k, v in app_state.WINDOW_PLACEMENTS.items() 
                                               if v == current_placement][0])
                self.placement_vars[app_name] = placement_var
                
                placement_menu = OptionMenu(placement_frame, placement_var, 
                                          *app_state.WINDOW_PLACEMENTS.keys(),
                                          command=lambda *args, a=app_name: self.on_placement_change(a))
                placement_menu.config(bg=app_state.theme['button'],
                                   fg=app_state.theme['fg'],
                                   activebackground=app_state.theme['active'])
                placement_menu["menu"].config(bg=app_state.theme['button'],
                                           fg=app_state.theme['fg'])
                placement_menu.pack(side='left', padx=2)
                
                # Border style dropdown
                border_frame = Frame(frame, bg=app_state.theme['bg'])
                border_frame.pack(side='left', padx=5)
                
                Label(border_frame, text="Border:",
                      bg=app_state.theme['bg'],
                      fg=app_state.theme['fg']).pack(side='left')
                
                current_style = app_state.border_styles.get(app_name, "no_change")
                border_var = StringVar(value=[k for k, v in app_state.BORDER_STYLES.items() 
                                            if v == current_style][0])
                self.border_vars[app_name] = border_var
                
                border_menu = OptionMenu(border_frame, border_var, 
                                       *app_state.BORDER_STYLES.keys(),
                                       command=lambda *args, a=app_name: self.on_border_change(a))
                border_menu.config(bg=app_state.theme['button'],
                                 fg=app_state.theme['fg'],
                                 activebackground=app_state.theme['active'])
                border_menu["menu"].config(bg=app_state.theme['button'],
                                         fg=app_state.theme['fg'])
                border_menu.pack(side='left', padx=2)
                
                # Add startup delay setting
                delay_frame = Frame(frame, bg=app_state.theme['bg'])
                delay_frame.pack(side='left', padx=5)
                
                Label(delay_frame, text="Startup Delay (s):",
                      bg=app_state.theme['bg'],
                      fg=app_state.theme['fg']).pack(side='left')
                
                delay_var = StringVar(value=str(app_state.startup_delays.get(app_name, 0)))
                delay_entry = Entry(delay_frame, textvariable=delay_var,
                                  bg=app_state.theme['button'],
                                  fg=app_state.theme['fg'],
                                  width=5)
                delay_entry.pack(side='left', padx=2)
                
                def save_delay(event, app=app_name, var=delay_var):
                    try:
                        delay = int(var.get())
                        if delay >= 0:
                            app_state.startup_delays[app] = delay
                            app_state.config["STARTUP_DELAYS"] = app_state.startup_delays
                            app_state.save_config()
                        else:
                            var.set("0")
                    except ValueError:
                        var.set(str(app_state.startup_delays.get(app, 0)))
                
                delay_entry.bind('<Return>', save_delay)
                delay_entry.bind('<FocusOut>', save_delay)
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue

    def on_volume_change(self, app_name, value):
        app_state.save_app_volume(app_name, int(float(value)))

    def on_pid_match_change(self, app_name):
        should_match_pid = bool(self.pid_match_vars[app_name].get())
        app_state.save_pid_match_app(app_name, should_match_pid)

    def on_hide_titlebar_change(self, app_name):
        should_hide = bool(self.hide_titlebar_vars[app_name].get())
        app_state.save_hide_titlebar_app(app_name, should_hide)

    def on_maximize_change(self, app_name):
        should_maximize = bool(self.maximize_vars[app_name].get())
        app_state.save_maximize_app(app_name, should_maximize)

    def on_resolution_change(self, app_name):
        enabled = bool(self.resolution_vars[app_name].get())
        preset = self.preset_vars[app_name].get() if enabled else None
        app_state.save_custom_resolution(app_name, enabled, preset)

    def on_placement_change(self, app_name):
        placement_name = self.placement_vars[app_name].get()
        placement = app_state.WINDOW_PLACEMENTS[placement_name]
        app_state.save_window_placement(app_name, placement)

    def on_border_change(self, app_name):
        style_name = self.border_vars[app_name].get()
        style = app_state.BORDER_STYLES[style_name]
        app_state.save_border_style(app_name, style)

    def on_mute_change(self, app_name):
        """Handle mute checkbox changes"""
        should_mute = self.mute_vars[app_name].get()
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                try:
                    process = psutil.Process(session.ProcessId)
                    if os.path.basename(process.exe()) == app_name:
                        volume = session.SimpleAudioVolume
                        if volume:
                            volume.SetMute(should_mute, None)
                            # Add app to exceptions if unmuting
                            if not should_mute and app_name not in app_state.exceptions_list:
                                app_state.add_exception(app_name)
                            # Remove from exceptions if muting
                            elif should_mute and app_name in app_state.exceptions_list:
                                app_state.remove_exception(app_name)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

    def update_mute_status(self):
        """Update mute status and volume for all apps"""
        sessions = AudioUtilities.GetAllSessions()
        
        for session in sessions:
            if not session.Process:
                continue
                
            try:
                process = psutil.Process(session.ProcessId)
                app_name = os.path.basename(process.exe())
                
                if app_name in self.mute_labels:
                    volume = session.SimpleAudioVolume
                    if volume:
                        is_muted = volume.GetMute()
                        current_volume = int(volume.GetMasterVolume() * 100)
                        
                        # Update mute checkbox state
                        if app_name in self.mute_vars:
                            self.mute_vars[app_name].set(1 if is_muted else 0)
                        
                        # Update mute status
                        self.mute_labels[app_name].config(
                            text="Muted" if is_muted else "Unmuted",
                            fg="#ff6b6b" if is_muted else "#69db7c"
                        )
                        
                        # Update volume label
                        self.volume_labels[app_name].config(
                            text=f"{current_volume}%",
                            fg=app_state.theme['fg']
                        )
                    
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        # Schedule next update
        if not self.window.winfo_exists():
            return
        self.window.after(100, self.update_mute_status)

# Function to check if a specific process ID is the foreground window
def is_foreground_process(pid):
    if pid <= 0:
        return False
    try:
        # Get the handle to the foreground window
        foreground_window = win32gui.GetForegroundWindow()
        # Get the process id of the foreground window
        _, foreground_pid = win32process.GetWindowThreadProcessId(foreground_window)

        if foreground_pid <= 0:
            return False

        fg_process = psutil.Process(foreground_pid)
        fg_process_exe_name = os.path.basename(fg_process.exe())

        bg_process = psutil.Process(pid)
        bg_process_exe_name = os.path.basename(bg_process.exe())

        if fg_process_exe_name != bg_process_exe_name:
            # Check if both processes are in the mute group
            if fg_process_exe_name in app_state.MUTE_GROUPS and bg_process_exe_name in app_state.MUTE_GROUPS:
                    return True
            return False
            
        # Check if app should match PID
        if bg_process_exe_name in app_state.pid_match_apps:
            return pid == foreground_pid
        return True  # Match by exe name only

    except psutil.NoSuchProcess:
        return False

# Function to update the lists in the GUI
def update_lists():
    # Remember the current selections
    selected_exception_index = lb_exceptions.curselection()
    selected_non_exception_index = lb_non_exceptions.curselection()

    # Clear the current listbox entries
    lb_exceptions.delete(0, END)
    lb_non_exceptions.delete(0, END)

    # Get the list of all the current sessions
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process:
            try:
                process = psutil.Process(session.ProcessId)
                process_exe_name = os.path.basename(process.exe())
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                process_exe_name = "N/A"

            # Populate the listboxes
            if process_exe_name in app_state.exceptions_list:
                lb_exceptions.insert(END, process_exe_name)
            else:
                lb_non_exceptions.insert(END, process_exe_name)

    # Restore the previous selections if possible
    if selected_exception_index:
        lb_exceptions.selection_set(selected_exception_index)
    if selected_non_exception_index:
        lb_non_exceptions.selection_set(selected_non_exception_index)

    # Schedule the next update
    app_state.root.after(100, update_lists)

# Function to mute/unmute applications
def mute_unmute_apps():
    global app_state

    if app_state.lock_var.get():
        app_state.root.after(100, mute_unmute_apps)
        return

    # Get the list of all the current sessions
    sessions = AudioUtilities.GetAllSessions()

    # First pass - check for audio activity
    non_zero_other = False
    peak_value = 0
    for session in sessions:
        if session.Process:
            try:
                process = psutil.Process(session.ProcessId)
                process_name = os.path.basename(process.exe())
                if process_name in app_state.exceptions_list:
                    volume = session.SimpleAudioVolume
                    if volume is not None:
                        audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                        peak_value = audio_meter.GetPeakValue()
                        if peak_value > 0:
                            non_zero_other = True
            except:
                continue

    if non_zero_other:
        app_state.zero_cnt = 0
    else:
        app_state.zero_cnt = app_state.zero_cnt + 1

    # Second pass - handle muting
    for session in sessions:
        if not session.Process:
            continue
            
        try:
            process = psutil.Process(session.ProcessId)
            process_name = os.path.basename(process.exe())
        except:
            continue

        volume = session.SimpleAudioVolume
        if volume is None:
            continue

        # Check if app has manual mute override
        manual_mute = None
        if hasattr(app_state, 'volume_control') and app_state.volume_control is not None:
            if process_name in app_state.volume_control.mute_vars:
                manual_mute = bool(app_state.volume_control.mute_vars[process_name].get())

        # If manual mute is set, respect it
        if manual_mute is not None:
            if volume.GetMute() != manual_mute:
                volume.SetMute(manual_mute, None)
                reason = "Manual Mute Override"
                print(f"{'Muted' if manual_mute else 'Unmuted'}({process.pid}): {process_name} - Reason: {reason}")
            continue

        # Rest of existing muting logic
        should_be_muted = False
        mute_reason = "Unknown"

        if process_name in app_state.exceptions_list:
            volume_value = float(app_state.get_app_volume(process_name)) / 100
            if volume.GetMute() == 0 and abs(volume.GetMasterVolume() - volume_value) > 0.001:
                volume.SetMasterVolume(volume_value, None)

            if app_state.force_mute_bg_var.get() == 1:
                should_be_muted = True
                mute_reason = "Force Mute Background"
        else:
            volume_value = float(app_state.get_app_volume(process_name)) / 100

            if volume.GetMute() == 0 and abs(volume.GetMasterVolume() - volume_value) > 0.001:
                volume.SetMasterVolume(volume_value, None)

            # Check if the process ID is the foreground process
            if app_state.force_mute_fg_var.get() == 1:
                should_be_muted = True
                mute_reason = "Force Mute Foreground"
            elif app_state.mute_foreground_when_background.get() == 1 and app_state.zero_cnt <= 30:
                should_be_muted = True
                mute_reason = "Background Audio Playing"
            elif is_foreground_process(process.pid):
                app_state.last_foreground_app_pid = process.pid
                # Unmute the audio if it's in the foreground
                should_be_muted = False
                mute_reason = "Foreground App"
            else:
                should_be_muted = True
                mute_reason = f"Not Foreground App {process.pid} {is_foreground_process(process.pid)}"
                if  app_state.mute_last_app.get() and process.pid == app_state.last_foreground_app_pid:
                    if not non_zero_other:
                        should_be_muted = False
                        mute_reason = "Last Active App"

        if volume.GetMute() == 0 and should_be_muted:
            volume.SetMute(1, None)
            print(f"Muted({process.pid}): {process_name} - Reason: {mute_reason}")
        elif volume.GetMute() == 1 and not should_be_muted:
            volume.SetMute(0, None)
            print(f"Unmuted({process.pid}): {process_name} - Reason: {mute_reason}")

    app_state.to_unmute.clear()
    app_state.root.after(100, mute_unmute_apps)

# Add this debug function at the top level
def debug_mute_decision(process_name, process_id, should_be_muted, reason):
    if process_name == "chrome.exe":
        print(f"\nDebug chrome.exe mute decision:")
        print(f"  PID: {process_id}")
        
        # Get foreground process info
        foreground_window = win32gui.GetForegroundWindow()
        _, foreground_pid = win32process.GetWindowThreadProcessId(foreground_window)
        try:
            fg_process = psutil.Process(foreground_pid)
            fg_process_name = os.path.basename(fg_process.exe())
        except:
            fg_process_name = "unknown"
            
        # Get background process info
        try:
            bg_process = psutil.Process(process_id)
            bg_process_name = os.path.basename(bg_process.exe())
        except:
            bg_process_name = "unknown"
            
        print(f"  Foreground process: {fg_process_name} (PID: {foreground_pid})")
        print(f"  Background process: {bg_process_name} (PID: {process_id})")
        print(f"  In exceptions list: {process_name in app_state.exceptions_list}")
        print(f"  Force mute background: {app_state.force_mute_bg_var.get()}")
        print(f"  Force mute foreground: {app_state.force_mute_fg_var.get()}")
        print(f"  Is foreground: {is_foreground_process(process_id)}")
        print(f"  Background audio playing: {app_state.zero_cnt <= 30}")
        print(f"  Is last active: {process_id == app_state.last_foreground_app_pid}")
        print(f"  Keep last active unmuted: {not app_state.mute_last_app.get()}")
        print(f"  Should be muted: {should_be_muted}")
        print(f"  Reason: {reason}")

class OptionsWindow:
    def __init__(self, parent, app_state):
        self.window = Toplevel(parent)
        self.window.title("Options")
        self.window.configure(bg=app_state.theme['bg'])
        
        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create main frame
        main_frame = Frame(self.window, bg=app_state.theme['bg'])
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Intervals section
        intervals_frame = LabelFrame(main_frame, text="Update Intervals", 
                                   bg=app_state.theme['bg'],
                                   fg=app_state.theme['fg'])
        intervals_frame.pack(fill='x', padx=5, pady=5)
        
        # Window check interval
        window_frame = Frame(intervals_frame, bg=app_state.theme['bg'])
        window_frame.pack(fill='x', padx=5, pady=2)
        
        Label(window_frame, text="Window Check Interval (ms):",
              bg=app_state.theme['bg'],
              fg=app_state.theme['fg']).pack(side='left')
        
        window_var = StringVar(value=str(app_state.options["window_check_interval"]))
        window_entry = Entry(window_frame, textvariable=window_var,
                           bg=app_state.theme['button'],
                           fg=app_state.theme['fg'],
                           width=10)
        window_entry.pack(side='right')
        
        # Volume check interval
        volume_frame = Frame(intervals_frame, bg=app_state.theme['bg'])
        volume_frame.pack(fill='x', padx=5, pady=2)
        
        Label(volume_frame, text="Volume Check Interval (ms):",
              bg=app_state.theme['bg'],
              fg=app_state.theme['fg']).pack(side='left')
        
        volume_var = StringVar(value=str(app_state.options["volume_check_interval"]))
        volume_entry = Entry(volume_frame, textvariable=volume_var,
                           bg=app_state.theme['button'],
                           fg=app_state.theme['fg'],
                           width=10)
        volume_entry.pack(side='right')
        
        # List update interval
        list_frame = Frame(intervals_frame, bg=app_state.theme['bg'])
        list_frame.pack(fill='x', padx=5, pady=2)
        
        Label(list_frame, text="List Update Interval (ms):",
              bg=app_state.theme['bg'],
              fg=app_state.theme['fg']).pack(side='left')
        
        list_var = StringVar(value=str(app_state.options["list_update_interval"]))
        list_entry = Entry(list_frame, textvariable=list_var,
                          bg=app_state.theme['button'],
                          fg=app_state.theme['fg'],
                          width=10)
        list_entry.pack(side='right')
        
        # Debug mode
        debug_frame = Frame(main_frame, bg=app_state.theme['bg'])
        debug_frame.pack(fill='x', padx=5, pady=5)
        
        debug_var = IntVar(value=int(app_state.options["debug_mode"]))
        Checkbutton(debug_frame, text="Debug Mode",
                   variable=debug_var,
                   bg=app_state.theme['bg'],
                   fg=app_state.theme['fg'],
                   selectcolor=app_state.theme['button'],
                   activebackground=app_state.theme['bg']).pack(side='left')
        
        # Add OCR section
        ocr_frame = LabelFrame(main_frame, text="OCR Settings", 
                             bg=app_state.theme['bg'],
                             fg=app_state.theme['fg'])
        ocr_frame.pack(fill='x', padx=5, pady=5)
        
        # Selection buttons
        selection_frame = Frame(ocr_frame, bg=app_state.theme['bg'])
        selection_frame.pack(fill='x', padx=5, pady=2)
        
        Button(selection_frame, text="Select Name Area",
               command=app_state.select_name_area,
               bg=app_state.theme['button'],
               fg=app_state.theme['fg'],
               activebackground=app_state.theme['active']).pack(side='left', padx=5)
        
        Button(selection_frame, text="Select Message Area",
               command=app_state.select_message_area,
               bg=app_state.theme['button'],
               fg=app_state.theme['fg'],
               activebackground=app_state.theme['active']).pack(side='left', padx=5)
        
        # Buttons
        button_frame = Frame(main_frame, bg=app_state.theme['bg'])
        button_frame.pack(fill='x', pady=10)
        
        def save_options():
            try:
                app_state.options["window_check_interval"] = int(window_var.get())
                app_state.options["volume_check_interval"] = int(volume_var.get())
                app_state.options["list_update_interval"] = int(list_var.get())
                app_state.options["debug_mode"] = bool(debug_var.get())
                app_state.save_options()
                self.window.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers for intervals")
        
        Button(button_frame, text="Save",
               command=save_options,
               bg=app_state.theme['button'],
               fg=app_state.theme['fg'],
               activebackground=app_state.theme['active']).pack(side='right', padx=5)
        
        Button(button_frame, text="Cancel",
               command=self.window.destroy,
               bg=app_state.theme['button'],
               fg=app_state.theme['fg'],
               activebackground=app_state.theme['active']).pack(side='right', padx=5)

if __name__ == "__main__":

    if not pyuac.isUserAdmin():
        pyuac.runAsAdmin(wait=False)
        sys.exit(0)
    
    # Create global state instance
    app_state = AppState()
    app_state.setup_main_window()

    # Create main frame for lists
    lists_frame = Frame(app_state.root, bg=app_state.theme['bg'])
    lists_frame.pack(fill='both', expand=True, padx=10, pady=10)

    # Left list (Exceptions)
    left_frame = Frame(lists_frame, bg=app_state.theme['bg'])
    left_frame.pack(side='left', fill='both', expand=True)
    
    label_exceptions = Label(left_frame, text="Exceptions (Not Muted)", 
                           bg=app_state.theme['bg'], 
                           fg=app_state.theme['fg'])
    label_exceptions.pack()
    lb_exceptions = Listbox(left_frame, bg='#3c3f41', fg='white', 
                           selectbackground='#4b6eaf',
                           height=10)
    lb_exceptions.pack(fill='both', expand=True)

    # Center buttons
    center_frame = Frame(lists_frame, bg=app_state.theme['bg'])
    center_frame.pack(side='left', padx=10)
    
    def safe_add_exception():
        selection = lb_non_exceptions.curselection()
        if not selection:  # If nothing is selected
            return
        app_name = lb_non_exceptions.get(selection)
        app_state.add_exception(app_name)

    def safe_remove_exception():
        selection = lb_exceptions.curselection()
        if not selection:  # If nothing is selected
            return
        app_name = lb_exceptions.get(selection)
        app_state.remove_exception(app_name)
    
    btn_add = Button(center_frame, text="◄ Add to Exceptions", 
                    command=safe_add_exception,
                    bg=app_state.theme['button'], 
                    fg=app_state.theme['fg'],
                    activebackground=app_state.theme['active'])
    btn_add.pack(pady=5)
    
    btn_remove = Button(center_frame, text="Remove from Exceptions ►", 
                       command=safe_remove_exception,
                       bg=app_state.theme['button'],
                       fg=app_state.theme['fg'],
                       activebackground=app_state.theme['active'])
    btn_remove.pack(pady=5)

    # Right list (Non-exceptions)
    right_frame = Frame(lists_frame, bg=app_state.theme['bg'])
    right_frame.pack(side='left', fill='both', expand=True)
    
    label_non_exceptions = Label(right_frame, text="Non-Exceptions (Muted)",
                               bg=app_state.theme['bg'],
                               fg=app_state.theme['fg'])
    label_non_exceptions.pack()
    lb_non_exceptions = Listbox(right_frame, bg='#3c3f41', fg='white',
                               selectbackground='#4b6eaf',
                               height=10)
    lb_non_exceptions.pack(fill='both', expand=True)

    # Middle Section (Checkboxes)
    checkbox_frame = Frame(app_state.root, bg=app_state.theme['bg'])
    checkbox_frame.pack(fill='x', padx=10, pady=5)

    # Left column of checkboxes
    left_checks = Frame(checkbox_frame, bg=app_state.theme['bg'])
    left_checks.pack(side='left', expand=True)
    
    cb_mute_last_app = Checkbutton(left_checks, text="Keep Last Active App Unmuted", 
                                  variable=app_state.mute_last_app,
                                  bg='#2b2b2b', fg='white', 
                                  selectcolor='#3c3f41', 
                                  activebackground='#2b2b2b')
    cb_mute_last_app.pack(in_=left_checks, anchor='w', pady=2)

    cb_force_mute_fg = Checkbutton(left_checks, text="Always Mute Foreground Apps", 
                                  variable=app_state.force_mute_fg_var,
                                  bg='#2b2b2b', fg='white', 
                                  selectcolor='#3c3f41', 
                                  activebackground='#2b2b2b')
    cb_force_mute_fg.pack(in_=left_checks, anchor='w', pady=2)

    cb_force_mute_bg = Checkbutton(left_checks, text="Always Mute Background Apps", 
                                  variable=app_state.force_mute_bg_var,
                                  bg='#2b2b2b', fg='white', 
                                  selectcolor='#3c3f41', 
                                  activebackground='#2b2b2b')
    cb_force_mute_bg.pack(in_=left_checks, anchor='w', pady=2)

    # Right column of checkboxes
    right_checks = Frame(checkbox_frame, bg=app_state.theme['bg'])
    right_checks.pack(side='left', expand=True)
    
    cb_lock = Checkbutton(right_checks, text="Pause Auto-Muting", 
                         variable=app_state.lock_var,
                         bg='#2b2b2b', fg='white', 
                         selectcolor='#3c3f41', 
                         activebackground='#2b2b2b')
    cb_lock.pack(in_=right_checks, anchor='w', pady=2)

    cb_mute_forground_when_background = Checkbutton(right_checks, 
                                                   text="Auto-Mute Active App When Others Play Sound",
                                                   variable=app_state.mute_foreground_when_background,
                                                   bg='#2b2b2b', fg='white', 
                                                   selectcolor='#3c3f41', 
                                                   activebackground='#2b2b2b')
    cb_mute_forground_when_background.pack(in_=right_checks, anchor='w', pady=2)

    # Bottom Section (Volume Control and Options Buttons)
    bottom_frame = Frame(app_state.root, bg=app_state.theme['bg'])
    bottom_frame.pack(fill='x', padx=10, pady=10)
    
    def show_volume_control():
        app_state.volume_control = VolumeControlWindow(app_state.root, app_state)

    def show_options():
        OptionsWindow(app_state.root, app_state)

    btn_volume_control = Button(bottom_frame, text="App Volume Settings",
                              command=show_volume_control,
                              bg=app_state.theme['button'],
                              fg=app_state.theme['fg'],
                              activebackground=app_state.theme['active'])
    btn_volume_control.pack(side='left', pady=5)

    btn_options = Button(bottom_frame, text="Options",
                        command=show_options,
                        bg=app_state.theme['button'],
                        fg=app_state.theme['fg'],
                        activebackground=app_state.theme['active'])
    btn_options.pack(side='right', pady=5)

    # Schedule the first update of the lists
    app_state.root.after(100, update_lists)
    app_state.root.after(100, mute_unmute_apps)

    # Update variable traces
    app_state.mute_last_app.trace_add("write", lambda *args: app_state.update_params())
    app_state.force_mute_fg_var.trace_add("write", lambda *args: app_state.update_params())
    app_state.force_mute_bg_var.trace_add("write", lambda *args: app_state.update_params())
    app_state.mute_foreground_when_background.trace_add("write", lambda *args: app_state.update_params())

    # Restore window position and size
    app_state.restore_window_state()

    # Start the GUI loop
    app_state.root.mainloop()