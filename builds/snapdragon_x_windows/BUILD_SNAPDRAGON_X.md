# Building PowerPoint Tracker for Snapdragon X Windows ARM64

This guide explains how to build a Snapdragon X Windows ARM64 executable for the PowerPoint Tracker application.

## Overview

The PowerPoint Tracker application will be compiled into a single ARM64 Windows executable (.exe) that can run on Windows 11 devices with Snapdragon X ARM64 processors.

## Prerequisites

### System Requirements
- **macOS** (required for cross-compilation to ARM64 Windows)
- Python 3.8 or higher
- Internet connection for downloading dependencies

### Hardware Requirements
- Minimum 4GB RAM
- At least 2GB free disk space for build process

## Quick Start

### 1. Run the Build Script

```bash
# Navigate to the build directory
cd builds/snapdragon_x_windows

# Run the build script
./build_snapdragon_x.sh
```

The build script will:
- Create/activate a virtual environment
- Install all required dependencies
- Cross-compile the application for Snapdragon X Windows ARM64
- Create a single executable file
- Generate a deployment package

### 2. Find Your Executable

After successful build, your executable will be located at:
```
builds/snapdragon_x_windows/dist/PowerPointTracker_SnapdragonX
```

## Manual Build Process

If you prefer to build manually or need to customize the process:

### 1. Set Up Environment

```bash
# Navigate to project root
cd /path/to/que-pilot

# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install -r builds/snapdragon_x_windows/requirements_snapdragon_x.txt
```

### 2. Build with PyInstaller

```bash
# Navigate to build directory
cd builds/snapdragon_x_windows

# Build using the spec file
pyinstaller --clean --noconfirm powerpoint_tracker_snapdragon_x.spec
```

### 3. Alternative Direct Build

```bash
# Build directly without spec file
pyinstaller \
    --clean \
    --onefile \
    --windowed \
    --target-arch arm64 \
    --name "PowerPointTracker_SnapdragonX" \
    --add-data "presentation:presentation" \
    --hidden-import "presentation_detector" \
    --hidden-import "tkinter" \
    --hidden-import "pptx" \
    --hidden-import "cv2" \
    --hidden-import "PIL" \
    --hidden-import "numpy" \
    --hidden-import "pytesseract" \
    --hidden-import "psutil" \
    --exclude-module "matplotlib" \
    --exclude-module "scipy" \
    --exclude-module "pandas" \
    --exclude-module "jupyter" \
    --exclude-module "IPython" \
    --exclude-module "test" \
    --exclude-module "unittest" \
    --exclude-module "doctest" \
    --exclude-module "Cocoa" \
    --exclude-module "Quartz" \
    --exclude-module "Foundation" \
    --exclude-module "AppKit" \
    --exclude-module "pyobjc" \
    presentation/main_app.py
```

## Build Configuration

### PyInstaller Spec File

The `powerpoint_tracker_snapdragon_x.spec` file contains:
- **Target Architecture**: ARM64 (Snapdragon X)
- **Output Format**: Single executable file
- **Hidden Imports**: All required modules including presentation_detector
- **Data Files**: Dependencies and resources
- **Exclusions**: Unnecessary modules to reduce size
- **Platform Specific**: Windows ARM64 only, excludes macOS modules

### Key Build Parameters

- `--target-arch arm64`: Specifies ARM64 architecture for Snapdragon X
- `--onefile`: Creates a single executable
- `--windowed`: Hides console window (GUI only)
- `--clean`: Cleans previous builds
- `--noconfirm`: Overwrites without confirmation

## Dependencies Included

The build includes these core dependencies:
- **python-pptx**: PowerPoint file handling
- **opencv-python**: Computer vision and image processing
- **pytesseract**: OCR text extraction
- **Pillow**: Image manipulation
- **numpy**: Numerical computing
- **psutil**: System process monitoring
- **tkinter**: GUI framework (built into Python)
- **presentation_detector**: Core application package

## Platform-Specific Handling

### Snapdragon X ARM64 Features
- Uses `pywin32-ctypes` for Windows API access
- Includes Windows-specific window detection
- Compatible with Windows 11 on Snapdragon X ARM64

### Excluded Dependencies
- `pyobjc-framework-Cocoa`: macOS only
- `pyobjc-framework-Quartz`: macOS only
- `Foundation`, `AppKit`: macOS only
- Platform-specific modules not needed for Windows

## Troubleshooting

### Common Issues

#### 1. Build Fails with Import Errors
```bash
# Solution: Add missing modules to hidden-imports
pyinstaller --hidden-import "missing_module" your_app.py
```

#### 2. Executable Too Large
```bash
# Solution: Exclude unnecessary modules
pyinstaller --exclude-module "large_module" your_app.py
```

#### 3. ARM64 Target Not Supported
```bash
# Solution: Update PyInstaller
pip install --upgrade pyinstaller>=5.13.0
```

#### 4. Cross-Compilation Issues
- Ensure you're running on macOS
- Check PyInstaller version compatibility
- Verify ARM64 target support

#### 5. Presentation Package Not Found
- Ensure the presentation directory exists
- Check the spec file paths are correct
- Verify the project structure

### Build Logs

Check build output for:
- Missing dependencies
- Import errors
- Architecture compatibility
- File size warnings
- Path resolution issues

## Deployment

### Target System Requirements

**Windows 11 on Snapdragon X ARM64:**
- Windows 11 (ARM64 version)
- Snapdragon X processor
- Minimum 2GB RAM
- Visual C++ Redistributable for ARM64 (may be required)

### Installation

1. Copy `PowerPointTracker_SnapdragonX` to target device
2. Run the executable directly (no installation required)
3. Ensure PowerPoint is installed on target system
4. Grant necessary permissions for window detection

### Testing

Test the executable on target Snapdragon X Windows device:
- Launch PowerPoint with a presentation
- Run the PowerPoint Tracker executable
- Verify auto-detection works
- Test slide navigation and sync features

## Performance Considerations

### Executable Size
- Typical size: 50-100MB (includes all dependencies)
- Single file deployment
- No external dependencies required

### Runtime Performance
- Optimized for ARM64 architecture
- Efficient memory usage
- Fast startup time
- Snapdragon X optimized

## Advanced Configuration

### Custom Icon
Add to spec file:
```python
exe = EXE(
    # ... other parameters ...
    icon='path/to/icon.ico',
)
```

### Version Information
Add to spec file:
```python
exe = EXE(
    # ... other parameters ...
    version='version_info.txt',
)
```

### Console Mode
For debugging, change in spec file:
```python
exe = EXE(
    # ... other parameters ...
    console=True,  # Show console window
)
```

## Support

For issues with Snapdragon X Windows builds:
1. Check PyInstaller documentation
2. Verify ARM64 target support
3. Test on actual Snapdragon X Windows device
4. Check Windows compatibility requirements

## Build Output

Successful build creates:
- `dist/PowerPointTracker_SnapdragonX` - Main executable
- `build/` - Temporary build files
- `PowerPointTracker_SnapdragonX_Package/` - Deployment package
- Build logs and error messages in terminal

The executable is ready for deployment on Snapdragon X Windows devices!
