# App Muter for Windows

App Muter is a simple Python utility for Windows that allows you to automatically mute all applications except for those on an exceptions list. By default, it is configured to keep Google Chrome (`chrome.exe`) unmuted, but you can easily modify the exceptions list through the user interface.

## Features

- Automatically mutes all applications not on the exceptions list.
- Allows for real-time updating of exceptions.
- Provides a simple GUI to manage exceptions.
- Saves exceptions between sessions.

## Prerequisites

Before you can run App Muter, make sure you have the following installed:

- Python 3.x
- `pycaw` library
- `psutil` library
- `pywin32` library

You can install the required libraries using `pip`:

```bash
pip install pycaw psutil pywin32
```

## Usage

To start using App Muter, simply run the script:

```bash
python app_muter.py
```

The GUI will appear with two lists: "Exceptions (Not Muted)" and "Non-Exceptions (Muted)". Applications will automatically be muted unless they are added to the exceptions list.

### Adding an Exception

1. Select an application from the "Non-Exceptions (Muted)" list.
2. Click the "Add to Exceptions" button.

### Removing an Exception

1. Select an application from the "Exceptions (Not Muted)" list.
2. Click the "Remove from Exceptions" button.

## Notes

- The muting functionality is based on the executable name of the application (e.g., `chrome.exe` for Google Chrome).
- The application must be running for it to appear in the lists.
- The script only works on Windows OS.

## License

This project is open-source and available under the [MIT License](LICENSE).

## Disclaimer

This software is provided "as is", without warranty of any kind. Use at your own risk.