"""
Entry point for application.

You may run this without a GUI by passing in the absolute path to the Template document and the Variables workbook
as command line arguments.
"""
import sys

from PyQt5.QtWidgets import QApplication

from gui.gui import MainWindow
from automation.automate import Automate

if __name__ == '__main__':
    if sys.argv[1:]:
        if len(sys.argv[1:]) != 2:
            print('To run this without a GUI, please supply the absolute paths to the Template and Variables files '
                  'enclosed in quotes. For example: ')
            print(r'python main.py "C:\path\to\template.docx" "C:\path\to\variables.xls"')
        else:
            template_file, vars_file = sys.argv[1], sys.argv[2]
            automation = Automate(in_gui=False)
            automation.run(template_file, vars_file)
    else:
        app = QApplication(sys.argv)
        main = MainWindow()
        main.show()

        sys.exit(app.exec_())
