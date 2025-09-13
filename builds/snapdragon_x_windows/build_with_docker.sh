#!/bin/bash

# Build Windows .exe using Docker with Windows containers
# This requires Docker Desktop with Windows containers support

set -e  # Exit on any error

echo "🐳 Building Windows .exe using Docker"
echo "====================================="

# Check if Docker is available
if ! command -v docker &> /dev/null; then
    echo "❌ Error: Docker is not installed"
    echo "   Please install Docker Desktop first"
    exit 1
fi

# Check if Docker Desktop is running
if ! docker info &> /dev/null; then
    echo "❌ Error: Docker Desktop is not running"
    echo "   Please start Docker Desktop first"
    exit 1
fi

# Get the project root directory
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
BUILD_DIR="$PROJECT_ROOT/builds/snapdragon_x_windows"

echo "📁 Project root: $PROJECT_ROOT"
echo "📁 Build dir: $BUILD_DIR"

# Check if we're on macOS
if [[ "$OSTYPE" != "darwin"* ]]; then
    echo "❌ Error: This script is designed for macOS"
    echo "   Docker Windows containers require macOS with Docker Desktop"
    exit 1
fi

# Check if Windows containers are available
echo "🔍 Checking Docker Windows container support..."
if docker version --format '{{.Server.Os}}' 2>/dev/null | grep -q "windows"; then
    echo "✅ Windows containers detected"
else
    echo "⚠️  Windows containers not detected"
    echo "   You may need to switch Docker Desktop to Windows containers mode"
    echo "   Right-click Docker Desktop → Switch to Windows containers"
    read -p "Continue anyway? (y/N): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Yy]$ ]]; then
        exit 1
    fi
fi

# Create a temporary directory for the build context
TEMP_DIR=$(mktemp -d)
echo "📁 Temporary build directory: $TEMP_DIR"

# Copy necessary files to temp directory
echo "📋 Preparing build context..."
cp -r "$PROJECT_ROOT/presentation" "$TEMP_DIR/"
cp "$BUILD_DIR/powerpoint_tracker_windows_exe.spec" "$TEMP_DIR/"
cp "$BUILD_DIR/requirements_snapdragon_x.txt" "$TEMP_DIR/"
cp "$BUILD_DIR/Dockerfile.windows" "$TEMP_DIR/Dockerfile"

# Build the Docker image
echo "🔨 Building Docker image..."
cd "$TEMP_DIR"
docker build -t powerpoint-tracker-windows-builder .

# Run the container to build the executable
echo "🚀 Building Windows .exe..."
docker run --rm -v "$TEMP_DIR/dist:/app/dist" powerpoint-tracker-windows-builder

# Check if the build was successful
if [ -f "$TEMP_DIR/dist/PowerPointTracker_SnapdragonX.exe" ]; then
    echo ""
    echo "✅ Windows .exe build successful!"
    echo "📁 Executable location: $TEMP_DIR/dist/PowerPointTracker_SnapdragonX.exe"
    
    # Copy the executable to the build directory
    mkdir -p "$BUILD_DIR/dist_docker"
    cp "$TEMP_DIR/dist/PowerPointTracker_SnapdragonX.exe" "$BUILD_DIR/dist_docker/"
    
    echo "📁 Copied to: $BUILD_DIR/dist_docker/PowerPointTracker_SnapdragonX.exe"
    echo "📊 Size: $(du -h "$BUILD_DIR/dist_docker/PowerPointTracker_SnapdragonX.exe" | cut -f1)"
    
    # Create deployment package
    echo "📦 Creating deployment package..."
    DEPLOY_DIR="$BUILD_DIR/dist_docker/PowerPointTracker_SnapdragonX_Docker_Package"
    mkdir -p "$DEPLOY_DIR"
    cp "$BUILD_DIR/dist_docker/PowerPointTracker_SnapdragonX.exe" "$DEPLOY_DIR/"
    
    cat > "$DEPLOY_DIR/README.txt" << EOF
PowerPoint Tracker for Snapdragon X Windows ARM64
=================================================

This package contains the PowerPoint Tracker application as a Windows .exe file
compiled for Snapdragon X Windows ARM64 using Docker.

Files:
- PowerPointTracker_SnapdragonX.exe: Main Windows executable

System Requirements:
- Windows 11 on Snapdragon X ARM64 processor
- PowerPoint installed on target system
- Visual C++ Redistributable for ARM64 (may be required)

Usage:
1. Copy PowerPointTracker_SnapdragonX.exe to your Snapdragon X Windows device
2. Double-click the .exe file to run (no installation required)
3. Ensure PowerPoint is installed on target system
4. Grant necessary permissions for window detection

Features:
- Automatic PowerPoint window detection
- Real-time slide tracking
- Cross-platform compatibility
- Self-contained Windows executable

Built using Docker with Windows containers.
EOF
    
    echo "📦 Deployment package created: $DEPLOY_DIR"
    
else
    echo "❌ Build failed!"
    echo "   Check the Docker build logs above for details"
    exit 1
fi

# Clean up
echo "🧹 Cleaning up..."
rm -rf "$TEMP_DIR"

echo "🎉 Docker build completed!"
echo ""
echo "📋 Next steps:"
echo "1. Test the .exe file on a Snapdragon X Windows device"
echo "2. Verify PowerPoint integration works correctly"
echo "3. Check window detection functionality"
