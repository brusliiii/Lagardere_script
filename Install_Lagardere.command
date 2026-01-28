#!/bin/zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"

notify() {
  /usr/bin/osascript -e "display dialog \"$1\" buttons {\"OK\"} default button \"OK\"" >/dev/null
}

find_python() {
  for v in 3.12 3.11 3.10; do
    if [[ -x "/Library/Frameworks/Python.framework/Versions/$v/bin/python3" ]]; then
      echo "/Library/Frameworks/Python.framework/Versions/$v/bin/python3"
      return 0
    fi
  done
  return 1
}

PYTHON="$(find_python || true)"

if [[ -z "$PYTHON" ]]; then
  notify "Python (3.12/3.11) is required. Installer will download and install it now."

  BASE="https://www.python.org/ftp/python"
  INDEX=$(curl -fsSL "$BASE/" || true)
  if [[ -z "$INDEX" ]]; then
    notify "Cannot reach python.org. Check internet and try again."
    exit 1
  fi

  VERSION=""
  PKG_URL=""
  VERSIONS=$(echo "$INDEX" | grep -oE '3\.(12|11)\.[0-9]+' | sort -Vr | uniq)
  for v in $VERSIONS; do
    url="$BASE/$v/python-$v-macos11.pkg"
    if curl -fsI "$url" >/dev/null; then
      VERSION="$v"
      PKG_URL="$url"
      break
    fi
  done

  if [[ -z "$VERSION" ]]; then
    notify "No macOS installer found for Python 3.12/3.11."
    exit 1
  fi

  TMPDIR=$(mktemp -d)
  cleanup() { rm -rf "$TMPDIR"; }
  trap cleanup EXIT

  curl -fL "$PKG_URL" -o "$TMPDIR/python.pkg"

  /usr/bin/osascript -e "do shell script \"/usr/sbin/installer -pkg '$TMPDIR/python.pkg' -target /\" with administrator privileges" >/dev/null

  PYTHON="$(find_python || true)"
  if [[ -z "$PYTHON" ]]; then
    notify "Python install failed."
    exit 1
  fi
fi

if [[ ! -x "$ROOT_DIR/build_mac_app.command" ]]; then
  notify "Missing build_mac_app.command in $ROOT_DIR"
  exit 1
fi

"$ROOT_DIR/build_mac_app.command"

APP_SRC="$ROOT_DIR/dist/Lagardere.app"
if [[ ! -d "$APP_SRC" ]]; then
  notify "Build failed: $APP_SRC not found."
  exit 1
fi

if [[ -f "$ROOT_DIR/dist/build_info.txt" ]]; then
  BUILD_INFO=$(cat "$ROOT_DIR/dist/build_info.txt")
else
  BUILD_INFO="(unknown)"
fi

APP_DST="/Applications/Lagardere.app"
if [[ -d "$APP_DST" ]]; then
  rm -rf "$APP_DST"
fi

if ! cp -R "$APP_SRC" "$APP_DST" 2>/dev/null; then
  APP_DST="$HOME/Applications/Lagardere.app"
  mkdir -p "$HOME/Applications"
  rm -rf "$APP_DST"
  cp -R "$APP_SRC" "$APP_DST"
fi

notify "Installed to: $APP_DST"
notify "Build info:\n$BUILD_INFO"
open "$APP_DST"
