from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from openpyxl.styles import NamedStyle, Font, Border, Side

highlight = NamedStyle(name="highlight")
highlight.font = Font(name='Arial', size=24)

math_test = [
    '12 + 13 =', 
    '12 + 13 =',
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =',
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =',
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =',
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =', 
    '12 + 13 =',
    '12 + 13 =', 
    '12 + 13 =', 
    ]

wb = Workbook()
wb.add_named_style(highlight)

dest_filename = 'tmp.xlsx'

ws = wb.active
ws.title = "math"

ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = True
ws.print_options.gridLines = True

ws.print_title_cols = 'A:B' # the first two cols
ws.print_title_rows = '1:1' # the first row

col_idx, row_idx = 0, 0
for col in ws.iter_cols(min_row=1, max_col=9, max_row=20):
    row_idx = 0
    if col_idx % 3 == 0:
        for cell in col:
            cell.style = highlight
            cell.value = math_test[row_idx]
            row_idx += 1
    col_idx += 1
    

wb.save(filename = dest_filename)

print('done')