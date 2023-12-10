import sys
from GUI import MainWindow
from BOM_Template import Template
from PyQt5.QtWidgets import QApplication

def Convert():
    FolderPath = GUI.Save_Path.text()
    FilePath = GUI.File_Path.text()
    FileName = GUI.File_Name.text()

    if (FolderPath != '') or (FilePath != '') or (FileName != ''):
        Save_Path = FolderPath + FileName + str(GUI.Format_ComboBox.currentText())
        
        Excel = Template(FilePath)
        Excel.DeleteTabs()
        Excel.DeleteColumn()
        Excel.AddHeader1()
        Excel.Change_RowHeight()
        Excel.Change_ColumnWidth()
        Excel.mergecells()

        Excel.SaveDocument(Save_Path)
        #GUI.ProgressBar_Progression(0,100)
        GUI.ProgressBar.setValue(100)

if __name__ == '__main__':    
    app = QApplication(sys.argv)
    GUI = MainWindow()
    GUI.show()

    GUI.ConvertButton.clicked.connect(Convert)

    app.exec_()