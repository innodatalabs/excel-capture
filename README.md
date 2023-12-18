# excel-capture

Captures Excel workbook as a set of PNG images

## Usage

```bash
python -m excel_capture [-f/--fast] [-w <max-width>] <excel-filename>
```

Tool will create images in the same directory as `<excel-filename>` - one for each worksheet.

By default, it will try to strip trsailing empty rows and columns. This process is somewhat slow as we have to
examine the content of many cells.
Flag `-f` can be used to skip this step and capture conservatively large area. This may
cause some whitespace at the bottom and/or right of the image.

## Building executable

```cmd
python -m venv .venv
call .\venv\Scripts\activate.bat
pip install .
pip install pyinstaller

pyinstaller excel_capture.py
```
