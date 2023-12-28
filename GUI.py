import os
import sys
import os.path
from PyQt5.uic import loadUi
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import QRegExp
from PyQt5.QtGui import QRegExpValidator
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()

        '''
        for dirpath, dirnames, filenames in os.walk("."):
            for filename in [f for f in filenames if f=="GUI.ui"]:
                GUI_Path = os.path.join(dirpath,filename)

        loadUi(GUI_Path,self)
        '''
        
        loadUi('GUI.ui',self)
        
        self.BrowseButton.clicked.connect(self.BrowseFiles)
        self.SaveButton.clicked.connect(self.BrowseDirectory)
        self.ClearButton.clicked.connect(self.Clear)

        FileFormat = ('.xlsx','.xls')
        self.Format_ComboBox.addItems(FileFormat)

        # Create a regular expression pattern that does not allow spaces
        regex = QRegExp("[^\\s\\(\\)\\/\\:\\*\\?\"\\<\\>\\|]*")
        validator = QRegExpValidator(regex)

        # Set the validator for the QLineEdit widgets
        self.Save_Path.setValidator(validator)
        self.File_Path.setValidator(validator)
        self.File_Name.setValidator(validator)

    def BrowseFiles(self):
        DesktopPath = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')  # Get the path to the desktop directory
        FilePath = QFileDialog.getOpenFileName(self,'Select an Excel file',DesktopPath,'Excel Files (*.xlsx *.xls)')
        self.File_Path.setText(FilePath[0])
        self.Save_Path.setText(FilePath[0].rsplit('/',1)[0] + '/')

    def BrowseDirectory(self):
        FolderPath = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Folder') + '/'
        self.Save_Path.setText(FolderPath)

    def PopUp_Message(self,Title,Message,Icon=QMessageBox.Information):
        ErrorMsg = QMessageBox()
        ErrorMsg.setWindowTitle(Title)
        ErrorMsg.setText(Message)
        ErrorMsg.setIcon(Icon)
        ErrorMsg.exec_()  # Ensure that the error message dialog remains open

    def Clear(self):
        self.Save_Path.setText('')
        self.File_Path.setText('')
        self.File_Name.setText('')
        self.CheckBox.setChecked(False)

        if self.ProgressBar.value() != 0: 
            self.ProgressBar_Progression(100,0,350)

    def ProgressBar_Progression(self, Start, End, Time):
        self.animation = QtCore.QPropertyAnimation(self.ProgressBar, b"value")  # Create a QPropertyAnimation object for the progress bar value
        self.animation.setDuration(Time)  # Set the duration of the animation in milliseconds
        self.animation.setStartValue(Start)  # Set the start value of the animation
        self.animation.setEndValue(End)  # Set the end value of the animation
        self.animation.start()  # Start the animation

if __name__ == '__main__':
    app = QApplication(sys.argv)
    GUI = MainWindow()
    GUI.show()
    app.exec_()