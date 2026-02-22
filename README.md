# Word to PDF Converter (Windows)

A Windows desktop app that converts `.doc` and `.docx` files to PDF using Microsoft Word automation.

## Features

- Rich GUI queue for multiple documents
- Add files or load an entire folder
- Remove selected entries or clear the queue
- Uses the Windows **Save As** dialog for each output PDF
- Activity log and progress bar
- Packaged as a single executable with PyInstaller

## Important requirement

This app uses Microsoft Word COM automation (`win32com`).
**Microsoft Word must be installed on the machine where conversion runs.**

## Run from source

```bash
python -m pip install -r requirements.txt
python app.py
```

## Build single executable (Windows)

Run:

```bat
build_exe.bat
```

The script prompts for an output folder (for example your Downloads folder).
Press Enter to use the default output:

`dist\WordToPdfConverter.exe`

The build now writes a detailed log to `build.log` in the chosen output folder, prints the last log lines on failure, and pauses before closing so you can read errors.

If build dependencies for `win32com` are missed, the script adds required hidden imports and searches both the chosen output folder and the default `dist` folder before reporting failure.

Tip: run `build_exe.bat --no-pause` if you don't want it to pause at the end.

## Usage

1. Click **Add Files** or **Add Folder**.
2. Click **Convert**.
3. For each input file, choose destination in the Windows **Save As** dialog.
4. Wait for progress completion and success message.
