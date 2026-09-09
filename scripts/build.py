"""Build a native desktop distribution on the current OS."""
import platform
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def main():
    subprocess.run([sys.executable, "-c", "import tkinter; import customtkinter"], check=True)
    command = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
               "--onedir", "--windowed", "--name", "LynxAutomator",
               "--collect-all", "customtkinter", "--hidden-import", "openpyxl",
               "--add-data", f"{ROOT / 'logo.png'}{';' if sys.platform == 'win32' else ':'}."]
    if sys.platform == "win32":
        command += ["--icon", str(ROOT / "favicon.ico")]
    command.append(str(ROOT / "LynxAutomator_v001alpha.py"))
    subprocess.run(command, cwd=ROOT, check=True)
    artifacts = ROOT / "artifacts"
    artifacts.mkdir(exist_ok=True)
    name = artifacts / f"LynxAutomator-{platform.system()}-{platform.machine()}"
    if sys.platform == "darwin":
        # ditto preserves the .app symlinks and executable permissions.
        subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent",
                        str(ROOT / "dist/LynxAutomator.app"), str(name) + ".zip"], check=True)
    else:
        shutil.make_archive(str(name), "zip" if sys.platform == "win32" else "gztar",
                            root_dir=ROOT / "dist", base_dir="LynxAutomator")


if __name__ == "__main__":
    main()
