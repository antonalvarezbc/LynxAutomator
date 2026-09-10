# LynxAutomator - Modular Application

## Description

This application has been developed to automate and facilitate various processes in wildlife monitoring, specifically for the Iberian lynx, but it can also be used for other species. The application is modular and consists of modules offering specific data management and analysis functionalities.

## Modules

### 1. Wildbook

This module includes features that facilitate the generation of the **bulk import** file for Wildbook, a platform used for identifying and tracking individuals through photographs.

### 2. Wildlife Insights

This module helps you to download images from the **Wildlife Insights** platform.

### 3. Wildlife Insights-Wildbook

This module automatically generates an Excel file for **bulk import** into Wildbook using the data and images obtained from Wildlife Insights, efficiently integrating both systems.
This module is also available online at [lynxautomator-wi-wb.streamlit.app](https://lynxautomator-wi-wb.streamlit.app/).

### 4. Iberian Lynx

This module is designed to facilitate the generation of Excel files necessary for tracking the Iberian lynx, allowing for more efficient organization and analysis of the data.

### 5. Functionalities

This module contains several useful tools for managing and correcting data related to camera traps:

- **Date adjustment:** If a camera trap was set up with an incorrect date, this feature allows you to adjust the capture dates of the images.
- **Frame extraction:** Enables you to extract specific frames from videos while retaining the original capture date.
- **Image renaming:** A tool for bulk renaming images according to specific patterns.


## Download 

To download the application, follow these steps:

1. Visit the [Releases](https://github.com/antonalvarezbc/LynxAutomator/releases/) page of the repository on GitHub.
2. Look for the latest version or the one you want to download.
3. Click on the selected version to access the download files.

## Installation

This application has two options:
- LynxAutomator.exe does not need to be installed
- LynxAutomator-Setup.exe is the installer (the same app, just with an installer).

If security software flags a downloaded build, verify its source and review the detection before running it.

## User guide

- [Manual actualizado en español](docs/manual.md)
- [Updated English manual](docs/manual.en.md)
- [Manual em português](docs/manual.pt.md)
- [Single-application architecture and Qt migration (Spanish)](docs/architecture.md)
- [Original DOCX guide (historical)](WIP%20LynxAutomator%20GUIDE%20.docx)

The Markdown manuals describe the single application on `feature/qt-migration`;
older release executables may behave differently.

## Linux, macOS and Windows builds

GitHub Actions is configured in `.github/workflows/build.yml`. On manual runs it tests and builds the application for:

| Runner | Package |
| --- | --- |
| Ubuntu 22.04, x86_64 | `LynxAutomator-Linux-x86_64.tar.gz` |
| Windows 2022, x86_64 | `LynxAutomator-Windows-AMD64.zip` |
| macOS 15, Apple Silicon | `LynxAutomator-Darwin-arm64.zip` |
| macOS 15, Intel | `LynxAutomator-Darwin-x86_64.zip` |

After pushing these changes to GitHub, open **Actions → Build desktop apps**.
Select a successful run and download the package under **Artifacts**. Manual
runs are available via **Run workflow** once the workflow is on the default
branch. Artifacts expire after 30 days; this workflow does not publish Releases.
Each build includes its resolved Python dependency versions in `dependencies.txt`.

Extract the complete archive before opening the executable; keep the accompanying
files together. Linux packages are built on Ubuntu 22.04 and need a compatible
graphical Linux environment. macOS packages target macOS 15 or newer, with
separate downloads for Intel and Apple Silicon. The macOS apps have no Developer
ID signature or notarization; public distribution through Gatekeeper requires a
separate signing setup with Apple credentials.

### Run from source

Use Python 3.12 with Tk support. On Ubuntu/Debian install `python3-venv` and
`python3-tk` through the system package manager. On macOS use a Python installation
that includes Tk. Check it with `python3 -m tkinter` before installing dependencies.

From this repository's directory, on Linux/macOS:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python LynxAutomator_v001alpha.py
```

On Windows, activate the environment with `.venv\Scripts\Activate.ps1` in
PowerShell. The `pywin32` dependency is installed only on Windows. OpenCV's
headless package handles video decoding; the desktop interface uses Tk.

The Wildlife Insights downloader additionally requires Google Cloud CLI with
`gsutil` available on `PATH` and credentials authorized to access the input
`gs://` locations. It accepts JPEG images, keeps existing output files, and reports
failed downloads. Cancel waits for the current transfer, with a five-minute timeout
per transfer. Google Cloud credentials are not bundled into the app or CI.

### Dates and data handling

Capture dates in JPEG EXIF are shifted while preserving unrelated metadata. On
Windows, file creation, modification and access times are updated. On Linux and
macOS, modification and access times are updated; file creation time is not
changed. When EXIF is unavailable, Unix uses modification time as the reference,
never `ctime` (which is a metadata-change time). Other file types retain their
embedded metadata and only their filesystem timestamps are adjusted.

CSV conversion rejects missing/duplicate deployment matches and invalid dates
instead of silently removing or multiplying images. Temporal grouping separates
projects as well as deployments. See `docs/logic-review.md` for the review and
remaining improvements.

### Build and test locally

```sh
python -m pip install -r requirements-build.txt
python -m unittest discover -s tests -v
python scripts/smoke_test.py
python scripts/build.py
python scripts/smoke_test.py --packaged
```

On Linux without a display, run the smoke tests using `xvfb-run -a`. Packaging must
run on the target OS. The build script includes the logo and CustomTkinter assets.
The supported entry point is `LynxAutomator_v001alpha.py`. Alpha mini has been
removed on `feature/qt-migration`; it remains available in Git history. The goal
is one stable application for all three platforms. The current interface still
uses Tk while migration to PySide6 is prepared; see the [migration plan](docs/interface-plan.md).


## Responsive background tasks

Both desktop editions share UI-independent processing services. Video extraction,
Excel processing and saving, folder scans, date correction, renaming and downloads
run in a worker thread. The bottom task panel shows progress and **Cancel**; input
controls are restored after completion, cancellation or failure. Closing requests
cancellation and waits for the worker to finish safely.

Cancellation is cooperative between frames, files or processing stages. A pending
Excel read/write, copy, decoder call or gsutil transfer must return first. Completed
outputs remain; original-file changes are not rolled back. New frames, copies and
Excel exports use temporary output to avoid publishing incomplete files.

`python scripts/smoke_test.py` now also runs real Tk responsiveness/cancellation
and application integration tests. On headless Linux use `xvfb-run -a`. Pure processing
and task tests remain runnable without Tk via `python -m unittest discover -s tests -v`.

### Compilación manual

Actions sólo se ejecuta mediante **Run workflow**. Ubuntu es la opción
predeterminada; se puede elegir otro sistema o todos. No se ejecuta por push ni
pull request y no publica pre-releases. Para trabajar localmente, véase
[la guía de Ubuntu](docs/ubuntu-local.md).

### Fotografías desde Camtrap DP

**Bulk Import → Camtrap DP** abre paquetes ZIP locales, muestra especies mediante
casillas y obtiene sus imágenes locales o remotas con progreso y cancelación.
No requiere elegir un Wildbook. Véase el [recorrido y ejemplo sintético](docs/manual.md#camtrap-dp-seleccionar-especies-y-obtener-fotografías).


Photo acquisition is integrated into the first **Bulk Import** tab for Wildlife Insights and Camtrap DP. Choose **Local photos** or **Download photographs**. **Download access** supports Google browser sign-in through Google Cloud CLI for WI, and in-memory Agouti API keys/Bearer tokens or Trapper API tokens for protected DP photo URLs. See the manuals for account permissions, scoped credentials and retry behavior.


Bulk Import starts with **Wildlife Insights** and also offers **Create from folder**, **Catalog**, **Camtrap DP**, **Agouti API** and **Trapper API**. The API sources load project exports into the same species → photographs → Excel workflow. Authorization is optional in the interface; private projects may require server-issued credentials. See the manuals for server and project ID inputs.

El editor Bulk Import propone nombres Excel con ubicación común y fecha/hora, ofrece **Previsualizar** en ventana y carga metadatos relacionados al elegir **Obtener valor de**. `MarkedIndividual.individualID` comienza desmarcado. **Agouti API (alpha)** y **Trapper API (alpha)** incluyen autorización y filtros antes de cargar; el diálogo distingue filtros remotos y locales. Consulte las limitaciones y el flujo con ZIP filtrado en los [manuales](docs/manual.md#exportación-y-selección-de-datos).

**Configurar Excel** abre una ventana emergente común a los seis orígenes. Carpeta y Catálogo permiten interpretar niveles de subcarpetas y asignar coordenadas por estación. Las API consultan valores reales para ofrecer filtros con multiselección: Agouti lee primero los despliegues; Trapper necesita el ZIP de metadatos. Las fotografías sólo se obtienen en el paso explícito **Usar fotos locales** / **Descargar fotos seleccionadas**.
