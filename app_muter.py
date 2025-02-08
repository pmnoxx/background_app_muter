import os
import sys
import winreg
import keyboard
import psutil
import pyuac
import toml
import win32gui
import win32process

from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
from tkinter import Tk, Listbox, Button, Label, END, Checkbutton, IntVar, Scale

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
        # Load configuration
        config = read_config("config.toml")
        self.DEFAULT_EXCEPTION_LIST = config.get("DEFAULT_EXCEPTIONS", ["chrome.exe", "firefox.exe", "msedge.exe"])
        self.MUTE_GROUPS = config.get("MUTE_GROUPS", [])

        # Initialize state variables
        self.exceptions_list = []
        self.to_unmute = []
        self.last_foreground_app_pid = None
        self.pressed_keys = set()
        self.zero_cnt = 0

        # Load exceptions
        if not self.load_exceptions():
            self.exceptions_list = self.DEFAULT_EXCEPTION_LIST.copy()

    def load_exceptions(self):
        try:
            with open('exceptions.txt', 'r') as file:
                for line in file:
                    # Remove newline and add to exceptions list
                    self.exceptions_list.append(line.strip())
                return True
        except FileNotFoundError:
            return False

    def save_exceptions(self):
        with open('exceptions.txt', 'w') as file:
            for item in self.exceptions_list:
                file.write("%s\n" % item)

    def add_exception(self, app_name):
        if app_name and app_name not in self.exceptions_list:
            self.exceptions_list.append(app_name)
            self.to_unmute.append(app_name)
            self.save_exceptions()

    def remove_exception(self, app_name):
        if app_name and app_name in self.exceptions_list:
            self.exceptions_list.remove(app_name)
            self.save_exceptions()

# Create global state instance
app_state = AppState()

def on_release(key):
    print(f"Released {key.name}")

def on_key_event(key: keyboard.KeyboardEvent):
    if keyboard.is_pressed("f5") and keyboard.is_pressed('windows'):
        if "f5" not in app_state.pressed_keys:
            app_state.pressed_keys.add("f5")
        else:
            app_state.pressed_keys.remove("f5")
    if keyboard.is_pressed("f6") and keyboard.is_pressed('windows'):
        if "f6" not in app_state.pressed_keys:
            app_state.pressed_keys.add("f6")
        else:
            app_state.pressed_keys.remove("f6")
    if keyboard.is_pressed("f7") and keyboard.is_pressed('windows'):
        if "f7" not in app_state.pressed_keys:
            app_state.pressed_keys.add("f7")
        else:
            app_state.pressed_keys.remove("f7")

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
    if app_state.pressed_keys:
        # this will not work correctly if keys are pressed a few times in 100ms
        if "f5" in app_state.pressed_keys:
            force_mute_fg_var.set(1 - force_mute_fg_var.get())
            print("pressed f5")
        if "f6" in app_state.pressed_keys:
            force_mute_bg_var.set(1 - force_mute_bg_var.get())
            print("pressed f5")
        app_state.pressed_keys.clear()

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
    root.after(100, update_lists)

# Function to mute/unmute applications
def mute_unmute_apps():
    global app_state

    if lock_var.get():
        root.after(100, mute_unmute_apps)
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
            background_volume_value = float(background_volume_var.get()) / 100

            if current_volume is None or abs(current_volume - background_volume_value) > 0.001:
                volume.SetMasterVolume(background_volume_value, None)

            if force_mute_bg_var.get() == 1:
                should_be_muted = True
        else:
            volume_value = float(volume_var.get()) / 100

            if current_volume is None or abs(current_volume - volume_value) > 0.001:
                volume.SetMasterVolume(volume_value, None)

            # Check if the process ID is the foreground process
            if force_mute_fg_var.get() == 1 or (mute_foreground_when_background.get() == 1 and app_state.zero_cnt <= 30):
                should_be_muted = True
            elif is_foreground_process(process_id):
                app_state.last_foreground_app_pid = process_id
                # Unmute the audio if it's in the foreground

                should_be_muted = False
            else:
                should_be_muted = True
                if not mute_last_app.get() and process_id == app_state.last_foreground_app_pid:
                    if not non_zero_other:
                        should_be_muted = False

        if volume.GetMute() == 0 and should_be_muted:
            volume.SetMute(1, None)
            print(f"Muted({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")
        elif volume.GetMute() == 1 and not should_be_muted:
            volume.SetMute(0, None)
            print(f"Unmuted({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")

    app_state.to_unmute.clear()
    root.after(100, mute_unmute_apps)

def is_admin():
    """Check if the current process is running with administrative privileges"""
    try:
        return pyuac.isUserAdmin()
    except:
        return False

def run_as_admin():
    """Run the current script with administrative privileges"""
    if not is_admin():
        print("This script requires administrative privileges to run.")
        try:
            pyuac.runAsAdmin(wait=False)
        except:
            pass
        sys.exit(0)
    else:
        # Your code here
        print("Running with administrative privileges.")

if __name__ == "__main__":
    run_as_admin()

    keyboard.hook(on_key_event)
    APP_GROUPS = [
        ["steam.exe", "steamwebhelper.exe"]

    ]

    # Create the main window
    root = Tk()
    root.title("App Muter")


    # Create the listboxes and labels
    label_exceptions = Label(root, text="Exceptions (Not Muted)")
    label_exceptions.pack()
    lb_exceptions = Listbox(root)
    lb_exceptions.pack()

    label_non_exceptions = Label(root, text="Non-Exceptions (Muted)")
    label_non_exceptions.pack()
    lb_non_exceptions = Listbox(root)
    lb_non_exceptions.pack()

    # Create the add and remove buttons
    btn_add = Button(root, text="Add to Exceptions", command=lambda: app_state.add_exception(lb_non_exceptions.get(lb_non_exceptions.curselection())))
    btn_add.pack()

    btn_remove = Button(root, text="Remove from Exceptions", command=lambda: app_state.remove_exception(lb_exceptions.get(lb_exceptions.curselection())))
    btn_remove.pack()


    def load_from_winreg():
        try:
            # Load dictionary back from registry
            loaded_dict = {}
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\AppMuter")
            for i in range(winreg.QueryInfoKey(key)[1]):
                value_name = winreg.EnumValue(key, i)[0]
                value = winreg.EnumValue(key, i)[1]
                loaded_dict[value_name] = value
            return loaded_dict
        except Exception as e:
            print(e)
            return {}


    def save_to_winreg(dict_to_save):
        # Save dictionary to registry
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\AppMuter")
        for key_name, value in dict_to_save.items():
            winreg.SetValueEx(key, key_name, 0, winreg.REG_DWORD, value)


    params = load_from_winreg() or {}
    print(params)

    # Checkbox for muting the last opened app
    mute_last_app = IntVar(value=params.get("mute_last_app") or 0)
    cb_mute_last_app = Checkbutton(root, text="Mute Last Opened App", variable=mute_last_app)
    cb_mute_last_app.pack()

    force_mute_fg_var = IntVar(value=params.get("force_mute_fg") or 0)
    cb_force_mute_fg = Checkbutton(root, text="Force mute fg", variable=force_mute_fg_var)
    cb_force_mute_fg.pack()

    force_mute_bg_var = IntVar(value=params.get("force_mute_bg") or 0)
    cb_force_mute_bg = Checkbutton(root, text="Force mute bg", variable=force_mute_bg_var)
    cb_force_mute_bg.pack()

    lock_var = IntVar(value=params.get("lock") or 0)
    cb_lock = Checkbutton(root, text="Lock Mode", variable=lock_var)
    cb_lock.pack()

    mute_foreground_when_background = IntVar(value=params.get("mute_foreground_when_background") or 0)
    cb_mute_forground_when_background = Checkbutton(root, text="Mute foreground when background is playing",
                                                    variable=mute_foreground_when_background)
    cb_mute_forground_when_background.pack()

    background_volume_var = IntVar(
        value=int(params.get("background_volume")) if params.get("background_volume") is not None else 100)

    background_volume_scale = Scale(root, from_=0, to=100, orient='horizontal', variable=background_volume_var)
    background_volume_scale.pack()

    volume_var = IntVar(value=int(params.get("volume")) if params.get("volume") is not None else 100)

    volume_scale = Scale(root, from_=0, to=100, orient='horizontal', variable=volume_var)
    volume_scale.pack()


    def update_params():
        params = {
            "mute_last_app": mute_last_app.get(),
            "force_mute_fg": force_mute_fg_var.get(),
            "force_mute_bg": force_mute_bg_var.get(),
            "volume": int(volume_var.get()),
            "background_volume": int(background_volume_var.get()),
            "mute_forgeround_when_background": mute_foreground_when_background.get(),
            "lock": 0, # int(lock_var.get()),
        }
        print("writing params", params)
        save_to_winreg(params)


    mute_last_app.trace("w", lambda *args: update_params())
    force_mute_fg_var.trace("w", lambda *args: update_params())
    force_mute_bg_var.trace("w", lambda *args: update_params())
    volume_var.trace("w", lambda *args: update_params())
    mute_foreground_when_background.trace("w", lambda *args: update_params())

    # Schedule the first update of the lists
    root.after(100, update_lists)
    root.after(100, mute_unmute_apps)

    # Start the GUI loop
    root.mainloop()