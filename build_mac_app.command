#!/bin/zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

PYTHON=""
for v in 3.12 3.11 3.10; do
  if [[ -x "/Library/Frameworks/Python.framework/Versions/$v/bin/python3" ]]; then
    PYTHON="/Library/Frameworks/Python.framework/Versions/$v/bin/python3"
    break
  fi
done

if [[ -z "$PYTHON" ]]; then
  PYTHON="python3"
fi

if [[ ! -d ".venv" ]]; then
  "$PYTHON" -m venv .venv
fi

source .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python - <<'PY'
from PIL import Image, ImageDraw

size = 1024
img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# Face
face_color = (255, 214, 10, 255)
outline = (230, 180, 0, 255)
margin = 80
bbox = (margin, margin, size - margin, size - margin)
d.ellipse(bbox, fill=face_color, outline=outline, width=12)

# Eyes
eye_y = int(size * 0.38)
eye_x_offset = int(size * 0.18)
eye_w = int(size * 0.10)
eye_h = int(size * 0.12)
left_eye = (
    size // 2 - eye_x_offset - eye_w // 2,
    eye_y,
    size // 2 - eye_x_offset + eye_w // 2,
    eye_y + eye_h,
)
right_eye = (
    size // 2 + eye_x_offset - eye_w // 2,
    eye_y,
    size // 2 + eye_x_offset + eye_w // 2,
    eye_y + eye_h,
)
d.ellipse(left_eye, fill=(60, 40, 0, 255))
d.ellipse(right_eye, fill=(60, 40, 0, 255))

# Smile
smile_bbox = (
    size * 0.25,
    size * 0.40,
    size * 0.75,
    size * 0.82,
)
d.arc(smile_bbox, start=20, end=160, fill=(60, 40, 0, 255), width=28)

img.save("icon.png")
PY

rm -rf icon.iconset
mkdir icon.iconset
for size in 16 32 64 128 256 512; do
  sips -z $size $size icon.png --out icon.iconset/icon_${size}x${size}.png >/dev/null
  sips -z $((size*2)) $((size*2)) icon.png --out icon.iconset/icon_${size}x${size}@2x.png >/dev/null
done
iconutil -c icns icon.iconset -o icon.icns

rm -rf build dist Lagardere.spec
pyinstaller --noconfirm --clean --windowed --name "Lagardere" --icon icon.icns desktop_app.py

echo "\nГотово: $ROOT_DIR/dist/Lagardere.app"
