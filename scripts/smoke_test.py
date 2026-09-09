"""Exercise startup in a fresh process, also for the frozen distribution."""
import argparse
import os
from pathlib import Path
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--packaged", action="store_true")
    args = parser.parse_args()
    if args.packaged:
        if sys.platform == "darwin":
            binary = ROOT / "dist/LynxAutomator.app/Contents/MacOS/LynxAutomator"
        elif sys.platform == "win32":
            binary = ROOT / "dist/LynxAutomator/LynxAutomator.exe"
        else:
            binary = ROOT / "dist/LynxAutomator/LynxAutomator"
        command = [str(binary)]
    else:
        command = [sys.executable, str(ROOT / "LynxAutomator_v001alpha.py")]
    with tempfile.TemporaryDirectory() as folder:
        marker = Path(folder) / "started.txt"
        env = dict(os.environ, LYNXAUTOMATOR_SMOKE_TEST=str(marker))
        subprocess.run(command, cwd=folder, env=env, timeout=90, check=True)
        if not marker.is_file():
            raise RuntimeError("The GUI did not finish initializing.")


if __name__ == "__main__":
    main()
