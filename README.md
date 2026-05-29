# GeoDIN GGF Date Tool

A small helper tool to update sampling dates inside GeoDIN `.GGF` graphics/layout files using dates from an Excel file.

The tool reads an Excel file, filters rows using the `Ionenbilanz` column, converts valid dates to GeoDIN format `YYYYMMDD`, and replaces the fixed `$SMPDATE$` slots inside a prepared 60-slot `.GGF` template.

Unused slots are filled with the dummy date:

```text
19000101
```

This keeps the `.GGF` file size unchanged and avoids breaking GeoDIN's internal layout structure.

## Main features

- Select `.GGF` template file
- Select Excel `.xlsx` file
- Auto-detect date and filter columns
- Default filter: `-5 <= Ionenbilanz <= 5`
- Warn about duplicate dates
- Fill unused GeoDIN date slots with dummy date
- Create log/report after processing
- Simple GUI for colleagues without Python knowledge

## Installation for development

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

This project uses a `src/` layout.

Before running the tool during development, install it in editable mode:

```bash
pip install -e .
```

## Run GUI

```bash
python main.py
```

## Run backend directly

```bash
python -m geodin_ggf_tool.backend
```

If you do not install the project, Python may show:
```
ModuleNotFoundError: No module named 'geodin_ggf_tool' 
```


## Package as EXE later

A basic PyInstaller command:

```bash
pyinstaller --onefile --windowed --name GeoDIN_GGF_Date_Tool main.py
```

The final `.exe` will appear in the `dist/` folder.

## Planned features

- Drag-and-drop support for `.GGF` and `.xlsx` files
- English/German interface language switch
- German and English logs/messages
- Better duplicate-date review before writing
- Excel preview before processing
- PyInstaller `.exe` build to run without Python
