"""
PyQt GUI classes
"""


import logging
import os
import sys
from functools import partial

from PyQt5.QtCore import Qt, QSize, QThread, pyqtSlot, pyqtSignal
from PyQt5.QtGui import QIcon, QFont, QColor
from PyQt5.QtWidgets import (
    QWidget,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QProgressBar,
    QLabel,
    QFileDialog,
)

from automation.automate import Automate


class FileDrop(QListWidget):
    """
    File drop that users may click or drag files to upload.
    """
    def __init__(self, parent, which_file):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.which_file = which_file
        self.cur_file = None

        self.setViewMode(QListWidget.IconMode)
        self.setFixedWidth(260)
        self.setFixedHeight(260)
        self.setItemAlignment(Qt.AlignHCenter)
        self.setIconSize(QSize(128, 128))
        self.setWordWrap(True)
        self.setDefaultContents()


    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[-1]
            if not url.isLocalFile() or not url.toLocalFile().endswith(self.which_file):
                event.setDropAction(Qt.IgnoreAction)
                event.ignore()
            else:
                event.accept()
        else:
            event.setDropAction(Qt.IgnoreAction)
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[-1]
            if not url.isLocalFile() or not url.toLocalFile().endswith(self.which_file):
                event.setDropAction(Qt.IgnoreAction)
                event.ignore()
            else:
                event.accept()
        else:
            event.setDropAction(Qt.IgnoreAction)
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)

            url = event.mimeData().urls()[-1]
            if not url.isLocalFile() or not url.toLocalFile().endswith(self.which_file):
                event.ignore()
            else:
                if self.which_file == '.docx':
                    icon = QIcon(os.path.join(sys.path[1], 'assets\\word_icon.png'))
                else:
                    icon = QIcon(os.path.join(sys.path[1], 'assets\\excel_icon.png'))

                file_name = str(url.toLocalFile()).split('/')[-1]
                self.cur_file = str(url.toLocalFile())
                self.clear()

                item = QListWidgetItem(icon, file_name)
                item.setFont(QFont(None, 12))
                item.setSizeHint(QSize(256, 256))
                item.setTextAlignment(Qt.AlignCenter)
                self.addItem(item)
        else:
            event.ignore()

    def setDefaultContents(self):
        self.clear()
        self.cur_file = None

        if self.which_file == '.docx':
            item = QListWidgetItem(QIcon(os.path.join(sys.path[1], 'assets\\word_icon.png')),
                                   '\nDrag and drop or Click to upload Template Document')
        else:
            item = QListWidgetItem(QIcon(os.path.join(sys.path[1], 'assets\\excel_icon.png')),
                                   '\nDrag and drop or Click to upload Variables Workbook')
        self.addItem(item)
        item.setSizeHint(QSize(256, 256))
        item.setTextAlignment(Qt.AlignCenter)



class MainWindow(QWidget):
    """
    Main window for uploading Template and Variables files.
    """
    def __init__(self):
        super(MainWindow, self).__init__()

        self.automation_thread = None
        self.threads = []

        self.setWindowTitle('AutoReport')
        self.setWindowIcon(QIcon(os.path.join(sys.path[1], 'assets\\small_logo.ico')))
        self.setMaximumSize(600, 400)

        self.TemplateDrop = FileDrop(self, '.docx')
        self.TemplateDrop.clicked.connect(partial(self.upload, '.docx'))

        self.VarsDrop = FileDrop(self, '.xls')
        self.VarsDrop.clicked.connect(partial(self.upload, '.xls'))

        vbox = QVBoxLayout()
        hbox = QHBoxLayout()

        hbox.addWidget(self.TemplateDrop, 0, Qt.AlignHCenter)
        hbox.addWidget(self.VarsDrop, 0, Qt.AlignHCenter)
        vbox.addLayout(hbox)

        self.submitBtn = QPushButton('Submit')
        self.submitBtn.clicked.connect(self.submit)

        vbox.addWidget(self.submitBtn)

        self.setLayout(vbox)
        self.setAcceptDrops(True)

    def upload(self, which_file):
        """
        Handles the upload of a file from a file-upload dialog. This function is ran after a user clicks 'Open'
        after selecting a template or vars file.

        Parameters
        ----------
        which_file : str
            Which file (Template or Variables) is being uploaded.
        """
        if which_file == '.xls':
            file_path = QFileDialog().getOpenFileName(filter='*.xls')[0]
            icon_name = os.path.join(sys.path[1], 'assets\\excel_icon.png')
            listbox = self.VarsDrop
        else:
            file_path = QFileDialog().getOpenFileName(filter='*.docx')[0]
            icon_name = os.path.join(sys.path[1], 'assets\\word_icon.png')
            listbox = self.TemplateDrop

        if not file_path:
            listbox.setDefaultContents()
        else:
            file_name = file_path.split('/')[-1]
            listbox.cur_file = file_path
            listbox.clear()
            item = QListWidgetItem(QIcon(icon_name), file_name)
            item.setFont(QFont(None, 12))
            item.setSizeHint(QSize(256, 256))
            item.setTextAlignment(Qt.AlignCenter)
            listbox.addItem(item)

    def submit(self):
        """
        Submit the inputted template and vars file to the automation thread.
        This code is ran upon press of the submit button.
        """
        if not self.TemplateDrop.cur_file or not self.VarsDrop.cur_file:
            return

        self.dialog = ProgressDialog()
        self.automation_thread = QThread()
        self.threads.append(self.automation_thread)
        automation = Automate()
        automation.moveToThread(self.automation_thread)
        automation.guiUpdater.moveToThread(self.automation_thread)
        automation.guiUpdater.logging_signal.connect(self.log)

        # When the thread starts, the run method of the Automation object will be ran with the template and vars files
        self.automation_thread.started.connect(
            partial(automation.run, self.TemplateDrop.cur_file, self.VarsDrop.cur_file)
        )

        self.automation_thread.start()
        self.dialog.exec()

    @pyqtSlot(int, str, int)
    def log(self, level, message, progress=None):
        """
        Add a log message to the progress dialog box.

        Parameters
        ----------
        level : int
            The level (i.e. logging level, info, warning, or critical) of the message.
        message : str
            The contents of the message.
        progress : int, optional
            The progress level from 0 to 100 of the message.
        """
        if progress is not None:
            self.dialog.progressBar.setValue(progress)

        if level in {logging.CRITICAL, logging.ERROR}:
            self.dialog.setWindowTitle('Error!')

        log = QListWidgetItem(message)
        color = {
            logging.WARNING: QColor(255, 204, 0),  # Yellow
            logging.ERROR: QColor(166, 68, 82),   # Red
            logging.CRITICAL: QColor(166, 68, 82)   # Red
        }.get(level)   # If level is not Warning, Error, or Critical, there will be a default (white) background.

        if color:
            log.setBackground(color)

        self.dialog.logsList.addItem(log)
        self.dialog.logsList.scrollToBottom()


class ProgressDialog(QDialog):
    """
    Progress dialog box that updates the user on the automation process.
    """

    def __init__(self):
        super(ProgressDialog, self).__init__()
        self.setWindowTitle('Running....')

        vbox = QVBoxLayout()

        self.progressBar = QProgressBar()
        self.progressBar.setAlignment(Qt.AlignCenter)
        self.progressBar.setValue(1)
        self.logsList = QListWidget()

        vbox.addWidget(self.progressBar)
        vbox.addWidget(self.logsList)

        self.setLayout(vbox)
