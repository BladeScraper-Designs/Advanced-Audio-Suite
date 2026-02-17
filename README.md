# Advanced Audio Suite for RC Transmitters

Desktop GUI app using Qt and Azure Speech SDK.

![](docs/img/AAS.png)

## Repository structure

- `AAS.py` - main GUI application
- `AAS.spec` - PyInstaller build spec
- `requirements.txt` - Python dependencies
- `in/` - input CSV files
- `samples/` - prebuilt sample packs
- `docs/` - docs/images

Runtime-generated folders (`config/`, `data/`, `out/`, `build/`, `dist/`) are created locally and ignored by git.

## FrSky audio CSV source

For FrSky users, default audio pack CSV files are available on FrSky's GitHub in the `audio` folder, under each language subfolder:

- https://github.com/FrSkyRC/ETHOS-Feedback-Community/blob/1.6/audio/

English example:

- https://github.com/FrSkyRC/ETHOS-Feedback-Community/blob/1.6/audio/en/en.csv

## Setup

From repo root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

From repo root:

```powershell
python AAS.py
```

On first run, enter your Azure Speech key and region when prompted.

## Build EXE

From repo root:

```powershell
.\.venv\Scripts\Activate.ps1
pip install pyinstaller
pyinstaller AAS.spec
```

Build output:

- `dist/AdvancedAudioSuite.exe`

## Build for Windows, Linux, and macOS (GitHub Actions)

This repo includes a GitHub Actions workflow at `.github/workflows/build-cross-platform.yml`.

- Run it manually from the **Actions** tab (`Build cross-platform binaries`), or
- publish a new GitHub Release to trigger it automatically.

Each run builds on:

- `windows-latest`
- `ubuntu-latest`
- `macos-latest`

Artifacts are uploaded per platform as:

- `AdvancedAudioSuite-windows-latest`
- `AdvancedAudioSuite-ubuntu-latest`
- `AdvancedAudioSuite-macos-latest`
