import sys
from PyQt5.QtCore import Qt, QThread
from PyQt5.QtWidgets import (QMainWindow, QLabel, QPushButton, QCheckBox, QMessageBox,
                             QFileDialog, QApplication, QGridLayout, QWidget, QLineEdit)
import ntpath
import pandas as pd


# Utility function to export SAS dataset using pandas
def export_sas(fileimport, fileexport):
    df = pd.read_sas(fileimport, format='sas7bdat', encoding='latin1')
    xlwriter = pd.ExcelWriter(fileexport, engine='xlsxwriter')
    df.to_excel(xlwriter, sheet_name="Sheet1", index=False)
    xlwriter.save()


# Background Thread for exporting SAS (it might take a while, we want to avoid clogging the main UI)
class ExportSASThread(QThread):

    def __init__(self, impfile, expfile):
        super().__init__()
        self.import_file = impfile
        self.export_file = expfile

    def run(self):
        export_sas(self.import_file, self.export_file)


# Main UI will handle user inputs
class Window(QMainWindow):

    def __init__(self):
        super().__init__()
        self.fileName = ''
        self.importFile = ''
        self.exportFile = ''
        self.initUI()

    # Create the UI
    def initUI(self):
        # Central Widget
        centralWidget = QWidget()
        # Create a grid layout
        grid = QGridLayout()
        # Select file input
        self.importLabel = QLabel('Choose SAS file to export:')
        grid.addWidget(self.importLabel, 0, 0, Qt.AlignTop)
        self.importLineEdit = QLineEdit()
        self.importLineEdit.setReadOnly(True)
        grid.addWidget(self.importLineEdit, 0, 1, Qt.AlignTop)
        self.importButton = QPushButton('Browse')
        self.importButton.clicked.connect(self.showImportDialog)
        grid.addWidget(self.importButton, 0, 2, Qt.AlignTop)
        # Select export path
        self.exportLabel = QLabel('Export to Excel:')
        grid.addWidget(self.exportLabel, 1, 0, Qt.AlignTop)
        self.exportLineEdit = QLineEdit()
        self.exportLineEdit.setReadOnly(True)
        self.exportLineEdit.setDisabled(True)
        grid.addWidget(self.exportLineEdit, 1, 1, Qt.AlignTop)
        self.exportButton = QPushButton('Browse')
        self.exportButton.setDisabled(True)
        self.exportButton.clicked.connect(self.showExportDialog)
        grid.addWidget(self.exportButton, 1, 2, Qt.AlignTop)
        # Checkbox to export to same place and name than SAS file
        self.checkBox = QCheckBox('Export with same path and name (but .xlsx)')
        self.checkBox.setChecked(True)
        self.checkBox.stateChanged.connect(self.changeCheckBox)
        grid.addWidget(self.checkBox, 2, 0, Qt.AlignTop)
        # Save Button
        self.saveBtn = QPushButton('Export')
        self.saveBtn.clicked.connect(self.save)
        grid.addWidget(self.saveBtn, 3, 1, Qt.AlignTop)
        # Thread
        self.export_sas_thread = ExportSASThread(self.importFile, self.exportFile)
        # Basic
        self.setWindowTitle('Export SAS')
        self.setGeometry(100, 300, 800, 100)
        # self.setLayout(grid)
        # Set the central widget
        centralWidget.setLayout(grid)
        self.setCentralWidget(centralWidget)
        # Display it
        self.show()

    # Display the dialog to chose a file to open
    def showImportDialog(self):
        fname = QFileDialog.getOpenFileName(self, 'Select SAS file to export', '/home', "SAS (*.sas7bdat)")[0]
        if fname:
            self.fileName = ntpath.basename(fname)
            self.importFile = fname
            self.importLineEdit.setText(fname)

    # Display a dialog to chose a file to save to (path and name)
    def showExportDialog(self):
        displayName = self.fileName.replace('sas7bdat', 'xlsx')
        fname = QFileDialog.getSaveFileName(self, 'Select file to export', displayName, "Excel (*.xlsx)")[0]
        if fname:
            self.exportFile = fname
            self.exportLineEdit.setText(fname)

    # If checkbox we don't want the user to select an export path
    def changeCheckBox(self):
        if not self.checkBox.isChecked():
            self.exportLineEdit.setDisabled(False)
            self.exportButton.setDisabled(False)
        else:
            self.exportLineEdit.setDisabled(True)
            self.exportButton.setDisabled(True)

    # Start the export process
    def save(self):
        if self.checkBox.isChecked():
            if self.importFile != '':
                self.exportFile = self.importFile.replace('sas7bdat', 'xlsx')
                self.export_sas_thread = ExportSASThread(self.importFile, self.exportFile)
                self.export_sas_thread.started.connect(self.thread_start)
                self.export_sas_thread.finished.connect(self.thread_done)
                self.export_sas_thread.start()
            else:
                QMessageBox.warning(self, 'Missing SAS File', 'Please select a SAS file to export')
        else:
            if self.importFile == '' or self.exportFile == '':
                QMessageBox.warning(self, 'Missing SAS File', 'Please select a SAS file')
            else:
                self.export_sas_thread = ExportSASThread(self.importFile, self.exportFile)
                self.export_sas_thread.started.connect(self.thread_start)
                self.export_sas_thread.finished.connect(self.thread_done)
                self.export_sas_thread.start()

    # Don't allow multiple threads, display to the user that their request is being processed
    def thread_start(self):
        self.saveBtn.setDisabled(True)
        self.statusBar().showMessage('Export in progress...')

    # Let the user know that the export is done
    def thread_done(self):
        self.saveBtn.setDisabled(False)
        self.statusBar().showMessage('Export successful!')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Window()
    sys.exit(app.exec_())
