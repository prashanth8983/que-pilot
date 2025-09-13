# Building Windows .exe using Virtual Machine

This approach uses a Windows virtual machine to build the .exe file locally.

## Prerequisites

- Virtualization software (VMware Fusion, Parallels Desktop, or VirtualBox)
- Windows 11 ARM64 virtual machine
- At least 8GB RAM allocated to the VM
- 20GB free disk space

## Setup Steps

### 1. Create Windows 11 ARM64 VM

**VMware Fusion:**
```bash
# Download Windows 11 ARM64 ISO
# Create new VM with Windows 11 ARM64
# Allocate at least 8GB RAM and 20GB disk space
```

**Parallels Desktop:**
```bash
# Use Parallels Desktop for Mac with Apple Silicon
# Create Windows 11 ARM64 VM
# Enable nested virtualization if needed
```

### 2. Install Python in VM

```powershell
# Download Python 3.11 for Windows ARM64
# Install with "Add to PATH" option
# Verify installation
python --version
pip --version
```

### 3. Transfer Project to VM

**Option A: Shared Folder**
- Set up shared folder between macOS and Windows VM
- Copy project files to shared folder

**Option B: Git Clone**
```powershell
# Install Git in Windows VM
git clone <your-repository-url>
cd que-pilot
```

### 4. Build in Windows VM

```powershell
# Navigate to project directory
cd que-pilot\builds\snapdragon_x_windows

# Install dependencies
pip install pyinstaller>=5.13.0
pip install python-pptx>=0.6.21
pip install opencv-python>=4.8.0
pip install pytesseract>=0.3.10
pip install Pillow>=10.0.0
pip install numpy>=1.24.0
pip install psutil>=5.9.0
pip install pywin32>=306

# Build the executable
pyinstaller --clean --noconfirm powerpoint_tracker_windows_exe.spec
```

### 5. Transfer Back to macOS

- Copy `dist/PowerPointTracker_SnapdragonX.exe` from VM
- Transfer via shared folder or network

## Advantages

- ✅ True Windows .exe file
- ✅ Full control over build environment
- ✅ Can test on Windows ARM64
- ✅ No external dependencies

## Disadvantages

- ❌ Requires Windows license
- ❌ High resource usage
- ❌ Complex setup
- ❌ Slow build times

## Alternative: Use GitHub Actions

For most users, the GitHub Actions approach is simpler and more reliable:

1. Push code to GitHub
2. Run the workflow
3. Download the .exe file

This avoids the complexity of setting up a Windows VM while still getting a true Windows .exe file.
