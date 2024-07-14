import time
import ctypes
import win32gui

def is_hsr_open():
    hwnd = win32gui.GetForegroundWindow()  # 根据当前活动窗口获取句柄
    Text = win32gui.GetWindowText(hwnd)
    warn_game = False
    cnt = 0
    return Text == "崩坏：星穹铁道" or Text == "Honkai: Star Rail"

PROCESS_PER_MONITOR_DPI_AWARE = 2

ctypes.windll.shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)

from pynput.mouse import  Listener as  MouseListener, Controller as MouseController
from pynput.keyboard import Key, Listener as KeyboardListener, Controller as KeyboardController

# List to store recorded events (combined mouse and keyboard)
all_events = []

def on_mouse_event(x, y, button, pressed):
    all_events.append((time.time(), f"Mouse {button} pressed at ({x}, {y})", (x, y, button, pressed)))

def on_keyboard_event(key):
    print(key)
    all_events.append((time.time(), f"Key pressed: {key}", (key, True)))

def on_keyboard_event_release(key):
    print(key)
    all_events.append((time.time(), f"Key pressed: {key}", (key, False)))


# Set up mouse and keyboard event listeners
mouse_listener = MouseListener(on_click=on_mouse_event)
keyboard_listener = KeyboardListener(on_press=on_keyboard_event, on_release=on_keyboard_event_release)

start_time = time.time()  # Record the start time
mouse_listener.start()
keyboard_listener.start()

# Wait for 15 seconds before starting the recording
time.sleep(25)


# Stop recording
mouse_listener.stop()
keyboard_listener.stop()

print("Replaying...")

# Replay events
mouse = MouseController()
keyboard = KeyboardController()

now = time.time()

for timestamp, event, param in all_events:
    time.sleep(max(0, (timestamp - start_time) - (time.time() - now)))  # Preserve timing
    if "Mouse" in event:
        x, y, button, pressed = param

        mouse.position = (x, y)
        if pressed:
            mouse.press(button)
        else:
            mouse.release(button)
        #if button == "left":
        #    mouse.press(MouseController().BUTTON_LEFT)
        #    mouse.release(MouseController().BUTTON_LEFT)
        #elif button == "right":
        #    mouse.press(MouseController().BUTTON_RIGHT)
        #    mouse.release(MouseController().BUTTON_RIGHT)
    else:
        print(param)
        key, pressed = param
        if key == Key.tab:
            continue
        if hasattr(key, 'char'):
            if pressed:
                keyboard.press(key.char)
            else:
                keyboard.release(key.char)
        else:
            if pressed:
                keyboard.press(key)
            else:
                keyboard.release(key)

print("Replay complete")
