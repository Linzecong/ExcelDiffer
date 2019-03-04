#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget, QMainWindow, QApplication, QTableWidgetItem,QTableWidget,QHBoxLayout,QVBoxLayout
import sys
from ExcelWidget import ExcelWidget
 
class ViewWidget(QWidget):
    def __init__(self):
        super(ViewWidget,self).__init__()

        self.OldTableWidget = ExcelWidget()
        self.NewTableWidget = ExcelWidget()
        self.MainLayout = QHBoxLayout()
        self.MainLayout.addWidget(self.OldTableWidget)
        self.MainLayout.addWidget(self.NewTableWidget)
        self.setLayout(self.MainLayout)
    
    def setOldTable(self,data):
        self.OldTableWidget.setData(data)
        
    def setNewTable(self,data):
        self.NewTableWidget.setData(data)

if __name__=="__main__":
    app = QApplication(sys.argv)
    main = ViewWidget()
    main.show()
    sys.exit(app.exec_())