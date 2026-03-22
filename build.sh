#!/bin/bash
set -e
BUILD_TYPE=${1:-Release}
BUILD_DIR=build-linux
echo "=== OpenSheet ET Linux Build ==="
echo "Build type: $BUILD_TYPE"

# Install dependencies if needed
if ! pkg-config --exists Qt6Core 2>/dev/null; then
    echo "Installing Qt6 dependencies..."
    sudo apt-get install -y \
        qt6-base-dev qt6-base-dev-tools \
        libgl1-mesa-dev libglu1-mesa-dev libglx-dev libegl1-mesa-dev \
        libxkbcommon-dev libxkbcommon-x11-dev \
        cmake build-essential 2>/dev/null || true
fi

rm -rf "$BUILD_DIR"
cmake -B "$BUILD_DIR" \
      -DCMAKE_BUILD_TYPE="$BUILD_TYPE" \
      -DCMAKE_INSTALL_PREFIX=/usr/local

cmake --build "$BUILD_DIR" --parallel "$(nproc)"
echo ""
echo "Build complete! Run: ./$BUILD_DIR/OpenSheetET"
