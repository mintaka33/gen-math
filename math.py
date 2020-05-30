from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
import numpy as np

math_group_rows = 25
math_group_cols = 6
math_group_num = 2

class MathTests():
    def __init__(self):
        self.testset_add = []
        self.testset_sub = []
        for i in range(30):
            for j in range(30):
                test = [str(i), '+', str(j), '=', ' ', ' ']
                self.testset_add.append(test)
        for i in range(30):
            for j in range(30):
                if (i - j) > 0:
                    test = [str(i), '-', str(j), '=', ' ', ' ']
                    self.testset_sub.append(test)

    def get_tests_add(self, num):
        tests = []
        test_array = np.random.randint(len(self.testset_add), size=(num))
        for i in range(num):
            tests.append(self.testset_add[test_array[i]])
        return tests

    def get_tests_sub(self, num):
        tests = []
        test_array = np.random.randint(len(self.testset_sub), size=(num))
        for i in range(num):
            tests.append(self.testset_sub[test_array[i]])
        return tests

class XmlWriter():
    def __init__(self):
        self.dest_filename = 'out.xlsx'
        self.total_cols = math_group_cols * math_group_num
        self.total_rows = math_group_rows
        self.font_size = 20
        self.column_width = 7
        self.mystyle = NamedStyle(name="mystyle")
        self.mystyle.font = Font(name='Arial', size=self.font_size)
        self.mystyle.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        self.wb = Workbook()
        self.wb.add_named_style(self.mystyle)
        self.ws = self.wb.active
        self.ws.title = "math"
        self.ws.page_setup.fitToWidth = 1
        self.ws.print_options.horizontalCentered = True
        self.ws.print_options.verticalCentered = True
        self.ws.print_options.gridLines = True

        for col_idx in range(self.total_cols):
            self.ws.column_dimensions[get_column_letter(col_idx+1)].width = self.column_width

    def write_group(self, group_idx, tests):
        start = group_idx * math_group_cols + 1
        end = start + math_group_cols - 1
        for row_idx, row in enumerate(self.ws.iter_rows(min_col=start, max_row=self.total_rows, max_col=end)):
            test = tests[row_idx]
            for col_idx, cell in enumerate(row):
                cell.style = self.mystyle
                cell.value = test[col_idx]

    def save_to_file(self):
        self.wb.save(self.dest_filename)

def gen_tests_xml():
    math = MathTests()
    xml = XmlWriter()
    # get random tests
    add_tests = math.get_tests_add(math_group_rows)
    sub_tests = math.get_tests_sub(math_group_rows)
    # write tests to xml file
    xml.write_group(0, add_tests)
    xml.write_group(1, sub_tests)
    # save xml file
    xml.save_to_file()

gen_tests_xml()

print('done')