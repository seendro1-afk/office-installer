#!/bin/bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")" && pwd)"
MACOS="$ROOT/macos"
DIST="$MACOS/dist"
APP="$DIST/Office Installer.app"
STAGING="$DIST/dmg-content"
DMG="$DIST/OfficeInstaller-macOS-arm64-v1.0.dmg"
ICON="$MACOS/assets/OfficeInstaller.icns"

echo "=== Office Installer: macOS Build ==="

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "Error: macOS is required."
    exit 1
fi

if [[ ! -f "$ICON" ]]; then
    echo "Error: application icon not found: $ICON"
    exit 1
fi

cd "$MACOS"

echo "[1/6] Building release..."
swift build -c release

BIN_DIR="$(swift build -c release --show-bin-path)"
BIN="$BIN_DIR/OfficeInstaller"

if [[ ! -f "$BIN" ]]; then
    echo "Error: executable not found."
    exit 1
fi

echo "[2/6] Creating application bundle..."
mkdir -p "$DIST"
rm -rf "$APP" "$STAGING"

mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"

cp "$BIN" "$APP/Contents/MacOS/OfficeInstaller"
chmod +x "$APP/Contents/MacOS/OfficeInstaller"

echo "[3/6] Copying resources and icon..."
find "$BIN_DIR" -maxdepth 1 -name '*.bundle' -exec cp -R {} "$APP/Contents/Resources/" \;
cp "$ICON" "$APP/Contents/Resources/OfficeInstaller.icns"

cat > "$APP/Contents/Info.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>OfficeInstaller</string>
    <key>CFBundleIdentifier</key>
    <string>com.seendro.officeinstaller</string>
    <key>CFBundleName</key>
    <string>Office Installer</string>
    <key>CFBundleDisplayName</key>
    <string>Office Installer</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0.0</string>
    <key>CFBundleVersion</key>
    <string>1</string>
    <key>CFBundleIconFile</key>
    <string>OfficeInstaller</string>
    <key>LSMinimumSystemVersion</key>
    <string>14.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
PLIST

echo "[4/6] Validating and signing application..."
plutil -lint "$APP/Contents/Info.plist"

codesign --force --sign - --timestamp=none "$APP"
codesign --verify --deep --strict --verbose=2 "$APP"

echo "[5/6] Creating DMG..."
mkdir -p "$STAGING"

ditto "$APP" "$STAGING/Office Installer.app"
ln -s /Applications "$STAGING/Applications"

hdiutil create \
    -volname "Office Installer" \
    -srcfolder "$STAGING" \
    -format UDZO \
    -ov \
    "$DMG"

echo "[6/6] Verifying DMG..."
hdiutil verify "$DMG"

echo ""
echo "=== BUILD COMPLETE ==="
ls -lh "$DMG"
