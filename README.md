# Word to PDF Converter (Windows)

This is a small Windows desktop utility that converts Microsoft Word files (`.doc` and `.docx`) into PDF files.

## Why this project exists

This repository was created as a practical experiment to see how Codex works in a real coding workflow: generating code, refining behavior, and iterating from feedback.

## How the application works

1. You start the app (`app.py` or the packaged `.exe`).
2. You add one or more Word files with **Add Files**, or load all supported files from a directory with **Add Folder**.
3. The app keeps selected files in a queue visible in the interface.
4. When you press **Convert**, each file is opened through Microsoft Word automation (`win32com`).
5. For every source file, Windows shows a **Save As** dialog so you choose where the output PDF will be written.
6. The app updates progress and writes status messages to the activity log until all queued files finish.

## Requirements

- Windows
- Microsoft Word installed (required for COM automation)
- Python 3.11+ (only if running from source)

## Run from source

```bash
python -m pip install -r requirements.txt
python app.py
```

## Build an executable (Windows)

Run:

```bat
build_exe.bat
```

The build script:

- installs dependencies from `requirements.txt`
- installs PyInstaller
- builds a single-file Windows executable named `WordToPdfConverter.exe`
- writes build output and errors to `build.log` in your selected output directory

## Typical usage flow

1. Add files/folder.
2. Click **Convert**.
3. Choose PDF destinations in the Save As dialogs.
4. Wait for completion confirmation.
