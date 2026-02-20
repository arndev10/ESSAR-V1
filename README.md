# ESSAR Generator

Desktop application that generates Excel reports from a template and input files. The template is bundled into the executable — no Python installation required for end users.

## Features

- GUI for selecting input/output folders
- Batch processing of Excel input files
- Template-based report generation
- Standalone .exe via PyInstaller (Windows)

## Tech Stack

**Language:** Python 3.10+
**Libraries:** pandas, openpyxl
**Packaging:** PyInstaller

## Usage (executable)

1. Build the executable (once): see [Build](#build-executable)
2. Open `Generador Essar.exe`
3. Select input folder (Excel files) and output folder
4. Click Generate

Output files are saved as `ESAR__{col1}__{col2}__{col3}.xlsx`.

## Usage (development)

```bash
pip install -r requirements.txt
python app.py
```

## Build Executable

```bash
pip install -r requirements.txt
pip install pyinstaller
build.bat
```

The .exe is created in the project root with the template bundled in.

## Project Structure

```
├── app.py                # GUI application
├── generador.py          # Excel generation logic
├── plantilla.xlsx        # Template (bundled in .exe)
├── GeneradorEssar.spec   # PyInstaller config
├── build.bat             # Build script
└── requirements.txt
```

## License

Open source — available for personal and educational use.
