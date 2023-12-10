import sys
import os
import os.path
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QErrorMessage
from PyQt5.uic import loadUi
from PyQt5.QtCore import QRegExp
from PyQt5.QtGui import QRegExpValidator


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()

        for dirpath, dirnames, filenames in os.walk("."):
            for filename in [f for f in filenames if f=="GUI.ui"]:
                GUI_Path = os.path.join(dirpath,filename)

        loadUi(GUI_Path,self)
        
        self.BrowseButton.clicked.connect(self.BrowseFiles)
        self.SaveButton.clicked.connect(self.BrowseDirectory)
        #self.ConvertButton.clicked.connect(self.Start)
        self.ClearButton.clicked.connect(self.Clear)

        FileFormat = ('.xlsx','.xls')
        self.Format_ComboBox.addItems(FileFormat)

        # Create a regular expression pattern that does not allow spaces
        regex = QRegExp("[^\\s]*")
        validator = QRegExpValidator(regex)

        # Set the validator for the QLineEdit widgets
        self.Save_Path.setValidator(validator)
        self.File_Path.setValidator(validator)
        self.File_Name.setValidator(validator)

    def BrowseFiles(self):
        FilePath = QFileDialog.getOpenFileName(self,'Select an Excel file','C:/_GIT/BOM_PYEXCEL','Excel Files (*.xlsx *.xls)')
        self.File_Path.setText(FilePath[0])
        self.Save_Path.setText(FilePath[0].rsplit('/',1)[0] + '/')

    def BrowseDirectory(self):
        FolderPath = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Folder') + '/'
        self.Save_Path.setText(FolderPath)

    def Start(self):
        self.FolderPath = self.Save_Path.text()
        self.FilePath = self.File_Path.text()
        self.FileName = self.File_Name.text()
        
        if (self.FolderPath == '') or (self.FilePath == '') or (self.FileName == ''):
            ErrorMsg = QErrorMessage()
            ErrorMsg.showMessage('All should be filled!')
        else:
            pass

    def Clear(self):
        self.Save_Path.setText('')
        self.File_Path.setText('')
        self.File_Name.setText('')
        #self.ProgressBar_Progression(100,0)
        self.ProgressBar.setValue(0)

    def ProgressBar_Progression(self,Start,End):
        animation = QtCore.QPropertyAnimation(self.ProgressBar, "setValue")  # Create a QPropertyAnimation object for the progress bar value
        animation.setDuration(3000)  # Set the duration of the animation to 3000 milliseconds
        animation.setStartValue(Start)  # Set the start value of the animation
        animation.setEndValue(End)  # Set the end value of the animation
        animation.start()  # Start the animation

if __name__ == '__main__':
    app = QApplication(sys.argv)
    GUI = MainWindow()
    GUI.show()
    app.exec_()