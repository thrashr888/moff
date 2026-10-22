#!/bin/bash
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
package="$root/apps/macos/MoffCompanion"
if [[ -z "${DEVELOPER_DIR:-}" && -d /Applications/Xcode.app/Contents/Developer ]]; then
  export DEVELOPER_DIR=/Applications/Xcode.app/Contents/Developer
fi

cd "$root"
cargo build -p moff-cli
npm --prefix harness run build
swift build --package-path "$package"
bin="$(swift build --package-path "$package" --show-bin-path)"
app="$package/.build/Moff Companion.app"
mkdir -p "$app/Contents/MacOS"
cp "$bin/MoffCompanion" "$app/Contents/MacOS/MoffCompanion"
cp "$root/apps/macos/Info.plist" "$app/Contents/Info.plist"
plutil -lint "$app/Contents/Info.plist"
printf '\nBuilt local development app: %s\n' "$app"
