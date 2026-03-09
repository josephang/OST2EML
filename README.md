# OST to EML Extractor

A standalone Python application featuring a modern graphical and dark mode UI designed to extract messages from proprietary Microsoft Outlook OST files, saving them locally as portable and standard `.eml` files. 

This repository provides an open-source, cost-free alternative to commercial OST to PST conversion pipelines, dumping emails with their original headers and preserving folder hierarchy.

## Features

- **Direct Extraction**: Uses the open-source [libpff](https://github.com/libyal/libpff) C-library bindings to read the `.ost` structures directly without needing Microsoft Outlook or expensive Aspose licenses.
- **Modern User Interface**: Built entirely in Python using [CustomTkinter](https://customtkinter.tomschimansky.com/) to deliver a beautiful, responsive, thread-safe dashboard.
- **Robust Exception Handling**: Employs error-safe `try-except` parsing for emails missing standard HTML structural blocks.
- **Long Path Safety**: Automatically truncates ridiculously long message subjects to abide by the Windows `MAX_PATH` character limits (Errno 2).

## Requirements

### Windows
- **Python 3.10+** (Added to your System PATH)
- **Microsoft C++ Build Tools** 
  - *Why?* The underlying extraction library (`libpff-python`) requires compilation from C-source on Windows systems because pre-built wheels aren't always provided.
  - *Download location*: [Visual Studio C++ Build Tools](https://visualstudio.microsoft.com/visual-cpp-build-tools/). 
  - *Selection*: When installing, select the **"Desktop development with C++"** workload.

## Installation & Build

Because the UI architecture (`tkinter`) and extraction libraries rely heavily on OS-level binaries, the final standalone executable should be compiled directly on your Windows PC.

1. **Clone** or download this directory to your Windows machine.
2. Ensure Python 3 and the Microsoft C++ Build Tools are installed.
3. Double-click the included `build.bat` script.
   * This script will automatically:
      - Install `pyinstaller`, `customtkinter` and `libpff-python`.
      - Compile the C-extensions for Python.
      - Bundle the user interface and logic into a portable Windows executable.
4. When the script completes, look inside the newly generated `dist/app/` folder.
5. Your standalone application will be located there as `OST_Extractor.exe`.

## How to Use

1. Launch **`OST_Extractor.exe`**.
2. **Select OST File**: Click `Browse...` and find your target `.ost` file on your filesystem. 
   *(Always use the Browse button to ensure absolute path structures are correctly passed to the extraction engine.)*
3. **Select Output Folder**: Click `Browse...` to choose an empty destination folder.
4. Click **Extract Messages**.

The application will begin mirroring your original Outlook folder structure into the output directory, iterating through each email, extracting the text/headers, and saving them as individually enumerated `.eml` files.

An active logging progress bar will display the operation's status in the UI safely via asynchronous multithreading.

## License

This software is released under the GNU General Public License v3.0 (GPLv3). Please see the `LICENSE` file for more details.
