import os
from excel2img.excel2img import ExcelFile
from PIL import ImageGrab

def column_index_to_address(index):
    a = 'ABCDEFGHIJKLMNOPQRSTUVQXYZ'[index % 26]
    if index >= 26:
        a = column_index_to_address(index // 26 - 1) + a
    return a

def is_non_empty_cell(cell):
    return cell.Value is not None and cell.Text != ''

COPY_MODE_SCREEN = 1
COPY_MODE_PRINT = 2

COPY_AS_VECTOR = -4147
COPY_AS_BITMAP = 2

def main(filename, fast=False):

    name, _ = os.path.splitext(filename)

    with ExcelFile.open(filename) as excel:
        num_sheets = excel.workbook.Sheets.Count
        for i in range(1, num_sheets + 1):
            sheet = excel.workbook.Sheets.Item(i)
            print(f'Capturing {i}/{num_sheets}: {sheet.Name}')
            last_used_column = sheet.UsedRange.Columns.Count
            last_used_row = sheet.UsedRange.Rows.Count

            print(f'\tlast_used_row={last_used_row}, last_used_column={last_used_column}')
            if not fast:
                while last_used_column > 1:
                    if any(is_non_empty_cell(sheet.Cells(row, last_used_column)) for row in range(1, last_used_row + 1)):
                        break
                    last_used_column -= 1
                print(f'\tafter shrinking columns: last_used_column={last_used_column}')

                while last_used_row > 1:
                    if any(is_non_empty_cell(sheet.Cells(last_used_row, col)) for col in range(1, last_used_column + 1)):
                        break
                    last_used_row -= 1
                print(f'\tafter shrinking rows: last_used_row={last_used_row}')

            cell_range = f'A1:{column_index_to_address(last_used_column)}{last_used_row}'
            print(f'\tcapturing cell range {cell_range}')

            sheet.Range(cell_range).CopyPicture(COPY_MODE_SCREEN, COPY_AS_BITMAP)
            im = ImageGrab.grabclipboard()
            img_filename = f'{name}-{i:02d}-{sheet.Name}.png'
            im.save(img_filename, 'PNG')
            print(f'\twritten sheet as {img_filename}')


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Intelligently captures Excel workbook as images (one per sheet)')
    parser.add_argument('filename', help='Excel workbook filename')
    parser.add_argument('--fast', '-f', action='store_true', help='Skip searching for the tight image boundaries (for speed)')

    args = parser.parse_args()

    main(args.filename, args.fast)
