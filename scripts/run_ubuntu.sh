#!/bin/sh
set -eu
repo_dir=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
cd "$repo_dir"
if [ ! -x .venv-ubuntu/bin/python ]; then
    if ! python3 -c 'import tkinter' >/dev/null 2>&1; then
        echo 'Instala primero: sudo apt install python3-tk python3-venv'
        exit 1
    fi
    python3 -m venv .venv-ubuntu
    .venv-ubuntu/bin/python -m pip install -r requirements.txt
fi
exec .venv-ubuntu/bin/python LynxAutomator_v001alpha.py
