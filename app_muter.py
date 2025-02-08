import os
import sys
import psutil
import toml
import win32gui
import win32process

from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from tkinter import Tk, Listbox, Button, Label, END, Checkbutton, IntVar, Scale, Toplevel, Frame, Entry, StringVar

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

    def setup_main_window(self):
        """Initialize main window settings"""
        self.root.title("App Muter")
        self.root.configure(bg=self.theme['bg'])
        self.root.minsize(self.window['min_width'], self.window['min_height'])
        
        # Center window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - self.window['width']/2)
        center_y = int(screen_height/2 - self.window['height']/2)
        self.root.geometry(f'{self.window["width"]}x{self.window["height"]}+{center_x}+{center_y}')
        
        # Bind window events
        self.root.bind("<Configure>", lambda e: self.save_window_state())

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
        self.update_app_list()
        
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
                
                Label(frame, text=app_name, 
                      bg=app_state.theme['bg'],
                      fg=app_state.theme['fg'],
                      width=20, anchor='w').pack(side='left')
                
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
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    
    def on_volume_change(self, app_name, value):
        app_state.save_app_volume(app_name, int(float(value)))

# Function to check if a specific process ID is the foreground window
def is_foreground_process(pid):
    if pid <= 0:
        return []
    try:
        # Get the handle to the foreground window
        foreground_window = win32gui.GetForegroundWindow()
        # Get the process id of the foreground window
        _, foreground_pid = win32process.GetWindowThreadProcessId(foreground_window)

        if foreground_pid <= 0:
            return []

        fg_process = psutil.Process(foreground_pid)
        fg_process_exe_name = os.path.basename(fg_process.exe())

        bg_process = psutil.Process(pid)
        bg_process_exe_name = os.path.basename(bg_process.exe())

        if fg_process_exe_name != bg_process_exe_name:
            for proc_group in app_state.MUTE_GROUPS:
                if fg_process_exe_name in proc_group and bg_process_exe_name in proc_group:
                    return True
        return pid == foreground_pid
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

    non_zero_other = False
    for session in sessions:
        if session.Process:
            process_id = session.ProcessId
            try:
                process = psutil.Process(session.ProcessId)
                process_name = os.path.basename(process.exe())
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

            # Get the simple audio volume of the session
            volume = session.SimpleAudioVolume

            # Skip muting Chrome windows
            if process_name in app_state.exceptions_list:
                if volume is not None:
                    audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                    peak_value = audio_meter.GetPeakValue()
                    if peak_value > 0:
                        non_zero_other = True

    if non_zero_other:
        app_state.zero_cnt = 0
    else:
        app_state.zero_cnt = app_state.zero_cnt + 1

    for session in sessions:
        if not session.Process:
            continue
        process_id = session.ProcessId
        try:
            process = psutil.Process(session.ProcessId)
            process_name = os.path.basename(process.exe())
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

        # Get the simple audio volume of the session
        volume = session.SimpleAudioVolume
        if volume is None:
            continue

        # Unmute items removed from exceptions
        if process_name in app_state.to_unmute and volume.GetMute() == 1:
            print(f"Unmuted({process_id}): {process_name} [{process_name}]")
            volume.SetMute(0, None)

        # Query the IAudioMeterInformation interface
        audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
        peak_value = audio_meter.GetPeakValue()
        current_volume = volume.GetMasterVolume()
        should_be_muted = False

        # Skip muting Chrome windows
        if process_name in app_state.exceptions_list:
            volume_value = float(app_state.get_app_volume(process_name)) / 100

            if current_volume is None or abs(current_volume - volume_value) > 0.001:
                volume.SetMasterVolume(volume_value, None)

            if app_state.force_mute_bg_var.get() == 1:
                should_be_muted = True
        else:
            volume_value = float(app_state.get_app_volume(process_name)) / 100

            if current_volume is None or abs(current_volume - volume_value) > 0.001:
                volume.SetMasterVolume(volume_value, None)

            # Check if the process ID is the foreground process
            if app_state.force_mute_fg_var.get() == 1 or (app_state.mute_foreground_when_background.get() == 1 and app_state.zero_cnt <= 30):
                should_be_muted = True
            elif is_foreground_process(process_id):
                app_state.last_foreground_app_pid = process_id
                # Unmute the audio if it's in the foreground

                should_be_muted = False
            else:
                should_be_muted = True
                if not app_state.mute_last_app.get() and process_id == app_state.last_foreground_app_pid:
                    if not non_zero_other:
                        should_be_muted = False

        if volume.GetMute() == 0 and should_be_muted:
            volume.SetMute(1, None)
            print(f"Muted({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")
        elif volume.GetMute() == 1 and not should_be_muted:
            volume.SetMute(0, None)
            print(f"Unmuted({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")

    app_state.to_unmute.clear()
    app_state.root.after(100, mute_unmute_apps)

if __name__ == "__main__":
    # Create global state instance
    app_state = AppState()
    app_state.setup_main_window()

    APP_GROUPS = [
        ["steam.exe", "steamwebhelper.exe"]
    ]

    # Create the listboxes and labels
    label_exceptions = Label(app_state.root, text="Exceptions (Not Muted)", 
                           bg=app_state.theme['bg'], 
                           fg=app_state.theme['fg'])
    label_exceptions.pack()
    lb_exceptions = Listbox(app_state.root, bg='#3c3f41', fg='white', selectbackground='#4b6eaf')
    lb_exceptions.pack()

    label_non_exceptions = Label(app_state.root, text="Non-Exceptions (Muted)", bg='#2b2b2b', fg='white')
    label_non_exceptions.pack()
    lb_non_exceptions = Listbox(app_state.root, bg='#3c3f41', fg='white', selectbackground='#4b6eaf')
    lb_non_exceptions.pack()

    # Create the add and remove buttons
    btn_add = Button(app_state.root, text="Add to Exceptions", 
                    command=lambda: app_state.add_exception(lb_non_exceptions.get(lb_non_exceptions.curselection())),
                    bg='#3c3f41', fg='white', activebackground='#4b6eaf')
    btn_add.pack()

    btn_remove = Button(app_state.root, text="Remove from Exceptions", 
                       command=lambda: app_state.remove_exception(lb_exceptions.get(lb_exceptions.curselection())),
                       bg='#3c3f41', fg='white', activebackground='#4b6eaf')
    btn_remove.pack()

    # Checkbox for keeping last app unmuted
    cb_mute_last_app = Checkbutton(app_state.root, text="Keep Last Active App Unmuted", 
                                  variable=app_state.mute_last_app,
                                  bg='#2b2b2b', fg='white', 
                                  selectcolor='#3c3f41', 
                                  activebackground='#2b2b2b')
    cb_mute_last_app.pack()

    cb_force_mute_fg = Checkbutton(app_state.root, text="Always Mute Foreground Apps", 
                                  variable=app_state.force_mute_fg_var,
                                  bg='#2b2b2b', fg='white', 
                                  selectcolor='#3c3f41', 
                                  activebackground='#2b2b2b')
    cb_force_mute_fg.pack()

    cb_force_mute_bg = Checkbutton(app_state.root, text="Always Mute Background Apps", 
                                  variable=app_state.force_mute_bg_var,
                                  bg='#2b2b2b', fg='white', 
                                  selectcolor='#3c3f41', 
                                  activebackground='#2b2b2b')
    cb_force_mute_bg.pack()

    cb_lock = Checkbutton(app_state.root, text="Pause Auto-Muting", 
                         variable=app_state.lock_var,
                         bg='#2b2b2b', fg='white', 
                         selectcolor='#3c3f41', 
                         activebackground='#2b2b2b')
    cb_lock.pack()

    cb_mute_forground_when_background = Checkbutton(app_state.root, 
                                                   text="Auto-Mute Active App When Others Play Sound",
                                                   variable=app_state.mute_foreground_when_background,
                                                   bg='#2b2b2b', fg='white', 
                                                   selectcolor='#3c3f41', 
                                                   activebackground='#2b2b2b')
    cb_mute_forground_when_background.pack()

    # Schedule the first update of the lists
    app_state.root.after(100, update_lists)
    app_state.root.after(100, mute_unmute_apps)

    # Update variable traces
    app_state.mute_last_app.trace("w", lambda *args: app_state.update_params())
    app_state.force_mute_fg_var.trace("w", lambda *args: app_state.update_params())
    app_state.force_mute_bg_var.trace("w", lambda *args: app_state.update_params())
    app_state.mute_foreground_when_background.trace("w", lambda *args: app_state.update_params())

    # Restore window position and size
    app_state.restore_window_state()
    
    # Add to main window creation:
    def show_volume_control():
        VolumeControlWindow(app_state.root, app_state)

    # Add button to main window
    btn_volume_control = Button(app_state.root, text="App Volume Settings",
                              command=show_volume_control,
                              bg=app_state.theme['button'],
                              fg=app_state.theme['fg'],
                              activebackground=app_state.theme['active'])
    btn_volume_control.pack(pady=5)

    # Start the GUI loop
    app_state.root.mainloop()