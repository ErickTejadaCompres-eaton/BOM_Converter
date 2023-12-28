import os
import sys
from GUI import MainWindow
from BOM_Template import Template
from PyQt5.QtWidgets import QApplication, QMessageBox

def Convert():
    FolderPath = GUI.Save_Path.text()
    FilePath = GUI.File_Path.text()
    FileName = GUI.File_Name.text()

    if (FolderPath != '') or (FilePath != '') or (FileName != ''):
        Save_Path = FolderPath + FileName + str(GUI.Format_ComboBox.currentText())
        
        Excel = Template(FilePath)
        Excel.DeleteTabs()
        Excel.DeleteColumns()
        Excel.Change_RowHeight()
        Excel.Change_ColumnWidth()
        Excel.MergeCells()

        Excel.SaveDocument(Save_Path)

        GUI.ProgressBar_Progression(0,100,2000)

        if GUI.CheckBox.isChecked():
            os.system(f'start excel "{Save_Path}"')
    else:
        GUI.PopUp_Message("Error","All fields should be filled.",QMessageBox.Critical)

if __name__ == '__main__':    
    app = QApplication(sys.argv)
    GUI = MainWindow()
    GUI.show()

    GUI.ConvertButton.clicked.connect(Convert)

    app.exec_()