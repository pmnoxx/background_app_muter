from tkinter import Tk, Listbox, Button, Label, END
from pycaw.pycaw import AudioUtilities, IAudioMeterInformation
import psutil
import os
import win32gui
import win32process
import winreg
from tkinter import Checkbutton, IntVar

from comtypes import CLSCTX_ALL
import comtypes

DEFAULT_EXCEPTION_LIST = ["chrome.exe", "firefox.exe", "msedge.exe"]

# List to hold the names of processes that should not be muted
exceptions_list = []
to_unmute = []


# Add a variable to keep track of the last foreground app's PID
last_foreground_app_pid = None


# Function to check if a specific process ID is the foreground window
def is_foreground_process(pid):
    # Get the handle to the foreground window
    foreground_window = win32gui.GetForegroundWindow()
    # Get the process id of the foreground window
    _, foreground_pid = win32process.GetWindowThreadProcessId(foreground_window)
    
    # Check if the process ID matches and return accordingly
    return pid == foreground_pid

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
            if process_exe_name in exceptions_list:
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
        global last_foreground_app_pid, skip_mute_last_app
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
                if process_name in exceptions_list:
                    if volume is not None:
                        audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                        peak_value = audio_meter.GetPeakValue()
                        if peak_value > 0:
                            non_zero_other = True


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

                # Unmute items removed from exceptions
                if process_name in to_unmute and volume.GetMute() == 1:
                    print(f"Unmuted({process_id}): {process_name} [{process_name}]")
                    volume.SetMute(0, None)

                # Skip muting Chrome windows
                if process_name in exceptions_list:
                    continue

                if volume is not None:
                    # current_volume = volume.GetMasterVolume()

                    # Query the IAudioMeterInformation interface
                    audio_meter = session._ctl.QueryInterface(IAudioMeterInformation)
                    peak_value = audio_meter.GetPeakValue()

                    # Check if the process ID is the foreground process
                    if is_foreground_process(process_id) and force_mute.get() == 0:
                        last_foreground_app_pid = process_id
                        # Unmute the audio if it's in the foreground
                        if volume.GetMute() == 1:
                            volume.SetMute(0, None)
                            print(f"Unmuted({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")
                    else:
                        if skip_mute_last_app.get() and process_id == last_foreground_app_pid:
                            if non_zero_other is False:
                                if restore_unmuted.get() == 1 and volume.GetMute() == 1:
                                    volume.SetMute(0, None)
                                    print(f"Muted(coliding sound)({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")
                                # Skip muting the last foreground app
                                continue


                        # Mute the audio if it's not the foreground
                        if volume.GetMute() == 0:
                            volume.SetMute(1, None)
                            print(f"Muted({process_id}): {process_name} [{process_name}] PeakValue: {peak_value}")



        to_unmute.clear()
        root.after(100, mute_unmute_apps)

# Function to save the exceptions to a file
def save_exceptions():
    with open('exceptions.txt', 'w') as file:
        for item in exceptions_list:
            file.write("%s\n" % item)

# Function to load the exceptions from a file
def load_exceptions():
    try:
        with open('exceptions.txt', 'r') as file:
            for line in file:
                # Remove newline and add to exceptions list
                exceptions_list.append(line.strip())
            return True
    except FileNotFoundError:
        # If the file does not exist, we can pass as we have an empty exceptions list
        pass
    return False

# Function to add an exception
def add_exception():
    if len(lb_non_exceptions.curselection()) == 0:
        return
    selected = lb_non_exceptions.get(lb_non_exceptions.curselection())
    if selected and selected not in exceptions_list:
        exceptions_list.append(selected)
        to_unmute.append(selected)
        save_exceptions()

# Function to remove an exception
def remove_exception():
    if len(lb_exceptions.curselection()) == 0:
        return
    print(lb_exceptions.curselection())
    selected = lb_exceptions.get(lb_exceptions.curselection())
    if selected and selected in exceptions_list:
        exceptions_list.remove(selected)
        save_exceptions()

# Load the exceptions when the application starts
if not load_exceptions():
    exceptions_list = DEFAULT_EXCEPTION_LIST

# Create the main window
root = Tk()
root.title("App Muter")


skip_mute_last_app = IntVar(value=1)

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
btn_add = Button(root, text="Add to Exceptions", command=add_exception)
btn_add.pack()

btn_remove = Button(root, text="Remove from Exceptions", command=remove_exception)
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
skip_mute_last_app = IntVar(value=params.get("skip_mute_last_app") or 1)
cb_skip_mute_last_app = Checkbutton(root, text="Don't Mute Last Opened App", variable=skip_mute_last_app)
cb_skip_mute_last_app.pack()

restore_unmuted = IntVar(value=params.get("restore_unmuted") or 0)
cb_restore_unmuted = Checkbutton(root, text="Restore unmuted", variable=restore_unmuted)
cb_restore_unmuted.pack()

force_mute = IntVar(value=params.get("force_mute") or 0)
cb_force_mute = Checkbutton(root, text="Force mute", variable=force_mute)
cb_force_mute.pack()


params = {
	"skip_mute_last_app": skip_mute_last_app.get(),
	"restore_unmuted": restore_unmuted.get(),
	"force_mute": force_mute.get(),
}





save_to_winreg(params)


def update_params():
  params = {
    "skip_mute_last_app": skip_mute_last_app.get(),
    "restore_unmuted": restore_unmuted.get(),
    "force_mute": force_mute.get()
  }
  print("writing params", params)
  save_to_winreg(params)

skip_mute_last_app.trace("w", lambda *args: update_params())
restore_unmuted.trace("w", lambda *args: update_params())
force_mute.trace("w", lambda *args: update_params())

# Schedule the first update of the lists
root.after(100, update_lists)
root.after(100, mute_unmute_apps)

# Start the GUI loop
root.mainloop()