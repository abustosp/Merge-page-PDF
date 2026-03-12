#!/usr/bin/env bash
set -euo pipefail

project_root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$project_root"

dist_path="$project_root/Ejecutable"
work_path="$project_root/temp_build"

pyinstaller="pyinstaller"
if [ -x "$project_root/venv/bin/pyinstaller" ]; then
  pyinstaller="$project_root/venv/bin/pyinstaller"
fi

if ! command -v "$pyinstaller" >/dev/null 2>&1; then
  echo "Error: pyinstaller no encontrado. Activa el venv o instala pyinstaller." >&2
  exit 1
fi

echo "=== Compilando Consolidar PDF ==="

icon_path="$project_root/bin/ABP-blanco-en-fondo-negro.ico"

"$pyinstaller" \
  --noconfirm \
  --clean \
  --onefile \
  --windowed \
  --distpath "$dist_path" \
  --workpath "$work_path" \
  --specpath "$work_path" \
  --name "ConsolidarPDF" \
  --icon "$icon_path" \
  "$project_root/gui_consolidar.py"

echo "=== Copiando archivos adicionales ==="

install -d "$dist_path/bin"
cp -a "$project_root/bin/." "$dist_path/bin/"

echo "=== Build completado ==="
echo "Ejecutable en: $dist_path/ConsolidarPDF"
