from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

mystyle = NamedStyle(name="mystyle")
mystyle.font = Font(name='Arial', size=20)
mystyle.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)

test1 = '12 + 13 ='
math_test1 = []
for i in range(20):
    math_test1.append(test1)


wb = Workbook()
wb.add_named_style(mystyle)

dest_filename = 'tmp.xlsx'

ws = wb.active
ws.title = "math"

ws.page_setup.fitToWidth = 1

ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = True
ws.print_options.gridLines = True


for col_idx in range(12):
    ws.column_dimensions[get_column_letter(col_idx+1)].width = 7

def test1():
    col_idx, row_idx = 0, 0
    for col in ws.iter_cols(min_row=1, max_col=9, max_row=22):
        row_idx = 0
        if col_idx % 3 == 0:
            for cell in col:
                cell.style = mystyle
                cell.value = math_test1[row_idx]
                row_idx += 1
        col_idx += 1

test2 = ['12', '+', '13', '=', ' ', ' ']
math_test2 = []
for i in range(30):
    math_test2.append(test2)
def test2(s):
    for row_idx, row in enumerate(ws.iter_rows(min_col=s, max_row=30, max_col=s+5)):
        for col_idx, cell in enumerate(row):
            cell.style = mystyle
            cell.value = math_test2[row_idx][col_idx]

test2(1)
test2(7)

wb.save(filename = dest_filename)

print('done')