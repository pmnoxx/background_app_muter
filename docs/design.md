# App Muter Design Document

## Current Issues
1. **Monolithic Structure**
   - All code is in a single file
   - AppState class handles too many responsibilities
   - UI and business logic are tightly coupled

2. **Configuration Management**
   - Multiple config files (runtime.toml, config.toml)
   - Inconsistent config loading and saving
   - No validation of config values

3. **Error Handling**
   - Inconsistent error handling
   - Many silent failures
   - Limited user feedback

## Proposed Architecture

### 1. Module Structure
python
app_muter/
├── src/
│ ├── init.py
│ ├── core/
│ │ ├── init.py
│ │ ├── audio.py # Audio control logic
│ │ ├── process.py # Process management
│ │ └── window.py # Window management
│ ├── config/
│ │ ├── init.py
│ │ ├── settings.py # Settings management
│ │ └── validators.py # Config validation
│ ├── ui/
│ │ ├── init.py
│ │ ├── main_window.py # Main window UI
│ │ ├── volume_control.py # Volume control window
│ │ └── options.py # Options window
│ └── utils/
│ ├── init.py
│ ├── icons.py # Icon generation
│ └── logging.py # Logging utilities
└── tests/
├── init.py
├── test_audio.py
├── test_process.py
└── test_window.py
