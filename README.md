# 📄 Document Reader

A lightweight Python GUI application for reading `.docx` and legacy `.doc` files. Built with **tkinter** — no heavy frameworks required.

![Python](https://img.shields.io/badge/Python-3.7%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey)

---

## Features

- Open and read `.docx` and `.doc` files through a clean dark-themed GUI
- Threaded file loading — UI stays responsive on large documents
- Word and line count displayed in the status bar
- Keyboard shortcuts (`Ctrl+O` to open, `Ctrl+Q` to quit)
- Fallback chain for legacy `.doc` support (antiword → LibreOffice)

## Screenshots

> *Add a screenshot of the app here after running it:*
>
> ```
> ![screenshot](https://github.com/zrnge/doc-reader/blob/main/Doc-Reader.png)
> ```

## Prerequisites

- Python 3.7 or higher
- `tkinter` (included with most Python installations)

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/zrnge/doc-reader.git
   cd doc-reader
   ```

2. **Install Python dependencies:**

   ```bash
   pip install python-docx
   ```

3. **(Optional) For legacy `.doc` file support, install one of the following:**

   **Linux:**

   ```bash
   sudo apt install antiword
   ```

   **Windows / macOS:**

   Install [LibreOffice](https://www.libreoffice.org/download/) and make sure it is available in your system `PATH`.

## Usage

```bash
python doc_reader.py
```

Click **Open File** (or press `Ctrl+O`) and select a `.docx` or `.doc` file. The content will be displayed in the text area.

## Keyboard Shortcuts

| Shortcut   | Action     |
|------------|------------|
| `Ctrl + O` | Open file  |
| `Ctrl + Q` | Quit app   |

## Project Structure

```
document-reader/
├── doc_reader.py   # Main application
├── README.md
└── LICENSE
```

## How It Works

| Format  | Method                                                                 |
|---------|------------------------------------------------------------------------|
| `.docx` | Parsed with `python-docx`                                             |
| `.doc`  | Extracted via `antiword` CLI, with LibreOffice headless as a fallback  |

## Contributing

Contributions are welcome. Fork the repo, create a branch, and open a pull request.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
