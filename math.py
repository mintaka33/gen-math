from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
import numpy as np

dest_filename = 'out.xlsx'
math_group_rows = 25
math_group_cols = 6
math_group_num = 2
total_cols = math_group_cols * math_group_num
total_rows = math_group_rows
font_size = 20
column_width = 7

mystyle = NamedStyle(name="mystyle")
mystyle.font = Font(name='Arial', size=font_size)
mystyle.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)

wb = Workbook()
wb.add_named_style(mystyle)

ws = wb.active

ws.title = "math2"
ws.page_setup.fitToWidth = 1
ws.print_options.horizontalCentered = True
ws.print_options.verticalCentered = True
ws.print_options.gridLines = True

for col_idx in range(total_cols):
    ws.column_dimensions[get_column_letter(col_idx+1)].width = column_width

math_testset = []
def gen_math():
    for i in range(30):
        for j in range(30):
            test = [str(i), '+', str(j), '=', ' ', ' ']
            math_testset.append(test)
gen_math()

def gen_math_group(group_idx):
    start = group_idx * math_group_cols + 1
    end = start + math_group_cols - 1
    test_array = np.random.randint(900, size=(total_rows))
    for row_idx, row in enumerate(ws.iter_rows(min_col=start, max_row=total_rows, max_col=end)):
        test = math_testset[test_array[row_idx]]
        for col_idx, cell in enumerate(row):
            cell.style = mystyle
            cell.value = test[col_idx]

for i in range(math_group_num):
    gen_math_group(i)

wb.save(filename = dest_filename)

print('done')