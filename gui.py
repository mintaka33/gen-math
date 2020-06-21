import sys
from PyQt5.QtWidgets import (QWidget, QGridLayout, QPushButton, QApplication, QLabel, QComboBox, QCheckBox)

class MathGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        combos = []
        for n in range(32):
            combo = QComboBox()
            for i in range(1, 50):
                if i % 5 == 0:
                    combo.addItem(str(i))
            combos.append(combo)

        combos2 = []
        for n in range(16):
            combo2 = QComboBox()
            for i in range(1, 50):
                combo2.addItem(str(i))
            combos2.append(combo2)

        grid.addWidget(QCheckBox(''), 0, 0)
        grid.addWidget(QLabel('加法:'), 0, 1)
        grid.addWidget(QLabel('A + B = []'), 0, 2)
        grid.addWidget(QLabel('加数A范围:'), 0, 3)
        grid.addWidget(combos[0], 0, 4)
        grid.addWidget(combos[1], 0, 5)
        grid.addWidget(QLabel('加数B范围:'), 0, 6)
        grid.addWidget(combos[2], 0, 7)
        grid.addWidget(combos[3], 0, 8)
        grid.addWidget(QLabel('页数:'), 0, 9)
        grid.addWidget(combos2[0], 0, 10)

        grid.addWidget(QCheckBox(''), 1, 0)
        grid.addWidget(QLabel('减法:'), 1, 1)
        grid.addWidget(QLabel('A - B = []'), 1, 2)
        grid.addWidget(QLabel('减法A范围:'), 1, 3)
        grid.addWidget(combos[4], 1, 4)
        grid.addWidget(combos[5], 1, 5)
        grid.addWidget(QLabel('减法B范围:'), 1, 6)
        grid.addWidget(combos[6], 1, 7)
        grid.addWidget(combos[7], 1, 8)
        grid.addWidget(QLabel('页数:'), 1, 9)
        grid.addWidget(combos2[1], 1, 10)

        grid.addWidget(QCheckBox(''), 2, 0)
        grid.addWidget(QLabel('加补1:'), 2, 1)
        grid.addWidget(QLabel('A + [] = C'), 2, 2)
        grid.addWidget(QLabel('加数A范围:'), 2, 3)
        grid.addWidget(combos[8], 2, 4)
        grid.addWidget(combos[9], 2, 5)
        grid.addWidget(QLabel('结果C范围:'), 2, 6)
        grid.addWidget(combos[10], 2, 7)
        grid.addWidget(combos[11], 2, 8)
        grid.addWidget(QLabel('页数:'), 2, 9)
        grid.addWidget(combos2[2], 2, 10)

        grid.addWidget(QCheckBox(''), 3, 0)
        grid.addWidget(QLabel('加补2:'), 3, 1)
        grid.addWidget(QLabel('[] + B = C'), 3, 2)
        grid.addWidget(QLabel('加数B范围:'), 3, 3)
        grid.addWidget(combos[12], 3, 4)
        grid.addWidget(combos[13], 3, 5)
        grid.addWidget(QLabel('结果C范围:'), 3, 6)
        grid.addWidget(combos[14], 3, 7)
        grid.addWidget(combos[15], 3, 8)
        grid.addWidget(QLabel('页数:'), 3, 9)
        grid.addWidget(combos2[3], 3, 10)

        grid.addWidget(QCheckBox(''), 4, 0)
        grid.addWidget(QLabel('减补1:'), 4, 1)
        grid.addWidget(QLabel('A - [] = C'), 4, 2)
        grid.addWidget(QLabel('减数A范围:'), 4, 3)
        grid.addWidget(combos[16], 4, 4)
        grid.addWidget(combos[17], 4, 5)
        grid.addWidget(QLabel('结果C范围:'), 4, 6)
        grid.addWidget(combos[18], 4, 7)
        grid.addWidget(combos[19], 4, 8)
        grid.addWidget(QLabel('页数:'), 4, 9)
        grid.addWidget(combos2[4], 4, 10)

        grid.addWidget(QCheckBox(''), 5, 0)
        grid.addWidget(QLabel('减补2:'), 5, 1)
        grid.addWidget(QLabel('[] - B = C'), 5, 2)
        grid.addWidget(QLabel('减数B范围:'), 5, 3)
        grid.addWidget(combos[20], 5, 4)
        grid.addWidget(combos[21], 5, 5)
        grid.addWidget(QLabel('结果C范围:'), 5, 6)
        grid.addWidget(combos[22], 5, 7)
        grid.addWidget(combos[23], 5, 8)
        grid.addWidget(QLabel('页数:'), 5, 9)
        grid.addWidget(combos2[5], 5, 10)

        self.move(300, 150)
        self.setWindowTitle('Math')
        self.show()

def main():
    app = QApplication(sys.argv)
    ex = MathGUI()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()



