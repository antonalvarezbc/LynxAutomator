#!/bin/sh
set -eu
repo_dir=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
# Editors packaged as Flatpak see a different Python/font environment.
if [ -n "${FLATPAK_ID:-}" ] && [ "${LYNX_ON_HOST:-0}" != 1 ]; then
    exec flatpak-spawn --host env LYNX_ON_HOST=1 sh "$repo_dir/scripts/run_ubuntu.sh"
fi
cd "$repo_dir"
python_bin=/usr/bin/python3
if ! "$python_bin" -c 'import tkinter, venv' >/dev/null 2>&1; then
    echo 'Faltan las dependencias gráficas nativas de Ubuntu.'
    echo 'Ejecuta en una terminal: sudo apt install python3-tk python3-venv'
    exit 1
fi
# Separate from .venv-ubuntu: that environment may use a portable Tk without Xft.
if [ ! -x .venv-ubuntu-native/bin/python ]; then
    "$python_bin" -m venv .venv-ubuntu-native
fi
if [ ! -f .venv-ubuntu-native/.lynx-requirements ] || ! cmp -s requirements.txt .venv-ubuntu-native/.lynx-requirements; then
    .venv-ubuntu-native/bin/python -m pip install -r requirements.txt
    cp requirements.txt .venv-ubuntu-native/.lynx-requirements
fi
exec .venv-ubuntu-native/bin/python LynxAutomator_v001alpha.py
