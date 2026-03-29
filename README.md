# WMM - Webcam Movement Manager

A Windows desktop application for saving and restoring DirectShow webcam settings, including PTZ (Pan/Tilt/Zoom) preset positions. Built as a single-file C# WinForms application with zero external dependencies.

## What It Does

Webcam Movement Manager lets you:

- **Save** all camera settings (brightness, contrast, exposure, focus, PTZ positions, etc.) to named profiles
- **Restore** those settings instantly -- either for all cameras or a specific one
- **Adjust** camera properties in real-time with sliders
- **Generate .bat files** that restore a profile with a double-click -- no GUI needed
- **Manage multiple cameras** -- each camera gets its own tab
- **Check for updates** directly from GitHub Releases

This is especially useful for PTZ cameras (like the DJI Osmo Pocket 3) where you need to return to exact Pan/Tilt/Zoom positions between sessions, or for any webcam setup where Windows resets your preferred settings after a reboot.

## Supported Camera Properties

### PTZ Controls
| Property | Description |
|----------|-------------|
| Pan | Horizontal rotation |
| Tilt | Vertical rotation |
| Zoom | Optical/digital zoom level |

### Camera Control
| Property | Description |
|----------|-------------|
| Exposure | Shutter speed / exposure time |
| Focus | Focus distance |
| Iris | Aperture opening |
| Roll | Rotation around optical axis |

### Video Processing
| Property | Description |
|----------|-------------|
| Brightness | Image brightness |
| Contrast | Tonal contrast |
| Hue | Color hue shift |
| Saturation | Color saturation |
| Sharpness | Image sharpness |
| Gamma | Gamma correction curve |
| White Balance | Color temperature |
| Backlight Compensation | Backlight correction |
| Gain | Signal amplification |
| Color Enable | Enable/disable color |

Each property supports **Auto** and **Manual** modes. Not all cameras support all properties -- the app only shows what your camera actually supports.

## How to Use

### GUI Mode

1. Run `WebcamSettingsManager.exe`
2. Each connected camera appears as a tab
3. Adjust sliders to change settings in real-time
4. Click **Save All** or **Save Selected** to save a profile
5. Select a profile from the dropdown and click **Restore All** or **Restore Selected** to apply it
6. Use **Generate .bat** to create a batch file for quick one-click restoring

### Toolbar Buttons

| Button | Description |
|--------|-------------|
| Refresh | Re-detect connected cameras |
| Save All | Save all camera settings to a new profile |
| Save Selected | Save only the current tab's camera |
| Restore All | Apply the selected profile to all matching cameras |
| Restore Selected | Apply the selected profile to the current camera only |
| Profile dropdown | Select which saved profile to restore |
| Delete | Delete the selected profile |
| Generate .bat | Create a batch file that restores the selected profile |

### Batch File Mode (Headless)

You can restore profiles without opening the GUI by using command-line arguments. This is what the generated `.bat` files use.

```
WebcamSettingsManager.exe --restore "dji"
```

Restore only a specific camera:
```
WebcamSettingsManager.exe --restore "dji" --device "OsmoPocket3"
```

List all saved profiles:
```
WebcamSettingsManager.exe --list-profiles
```

List all connected cameras:
```
WebcamSettingsManager.exe --list-devices
```

Show help:
```
WebcamSettingsManager.exe --help
```

#### Example .bat file

```bat
@echo off
"C:\Path\To\WebcamSettingsManager.exe" --restore "dji"
exit
```

Double-click this file and your camera settings are restored instantly. You can also add it to Windows Task Scheduler or your startup folder to auto-apply settings on boot.

## Profiles

Profiles are stored as JSON files in:
```
%APPDATA%\WebcamSettings\profiles\
```

You can open this folder from **File > Open Profiles Folder** in the menu. Profiles can also be imported/exported via **File > Import Profile** and **File > Export Profile**.

Cameras are matched by their device path (a unique hardware identifier), so profiles work reliably even if you have multiple cameras of the same model.

## Updates

The app can check for updates from GitHub Releases via **Help > Check for Updates**. If a new version is available, it shows the release notes and offers to download and install it automatically.

## Building from Source

No Visual Studio or SDK required. The app compiles with the C# compiler built into Windows 10/11:

```
build.bat
```

This calls `csc.exe` from `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\` and produces `WebcamSettingsManager.exe`. The entire application is a single `.cs` file with no external dependencies.

### Requirements

- Windows 10 or later
- .NET Framework 4.x (pre-installed on Windows 10/11)
- A DirectShow-compatible webcam

## License

This project is open source.

## Acknowledgements

This application was written with the help of AI (Claude by Anthropic).
