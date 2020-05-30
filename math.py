from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment
import numpy as np
from datetime import datetime

math_group_rows = 25*7
math_group_cols = 6
math_group_num = 2
min1, max1 = 5, 25
min2, max2 = 5, 25

def gen_outfile_name():
    n = datetime.now()
    time_str = 'math_'
    time_str += str(n.year) + '-' + str(n.month).zfill(2) + '-' + str(n.day).zfill(2)
    time_str += '_' + str(n.hour).zfill(2) + '-' + str(n.minute).zfill(2) + '-' + str(n.second).zfill(2)
    time_str += '_' + str(n.microsecond).zfill(6)
    time_str += '.xlsx'
    return time_str

class MathTests():
    def __init__(self):
        self.testset_add = []
        self.testset_sub = []
        for i in range(min1, max1):
            for j in range(min2, max2):
                test = [str(i), '+', str(j), '=', ' ', ' ']
                self.testset_add.append(test)
        print('INFO: total', len(self.testset_add), 'tests in ADD set')
        for i in range(min1, max1):
            for j in range(min2, max2):
                if (i - j) > 0:
                    test = [str(i), '-', str(j), '=', ' ', ' ']
                    self.testset_sub.append(test)
        print('INFO: total', len(self.testset_sub), 'tests in SUB set')

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
        self.dest_filename = gen_outfile_name()
        self.total_cols = math_group_cols * math_group_num
        self.total_rows = math_group_rows
        self.font_size = 21
        self.column_width = 7
        self.mystyle = NamedStyle(name="mystyle")
        self.mystyle.font = Font(name='Arial', size=self.font_size)
        self.mystyle.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        self.wb = Workbook()
        self.wb.add_named_style(self.mystyle)
        self.ws = self.wb.active
        self.ws.title = "math"
        self.ws.page_margins.left = 0.5
        self.ws.page_margins.right = 0.5
        self.ws.page_margins.top = 0.75
        self.ws.page_margins.bottom = 0.75
        self.ws.page_setup.fitToWidth = 1
        self.ws.print_options.horizontalCentered = True
        self.ws.print_options.verticalCentered = True
        self.ws.print_options.gridLines = True
        # set column width
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
gen_tests_xml()

print('done')