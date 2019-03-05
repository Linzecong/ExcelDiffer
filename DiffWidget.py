#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget, QTabWidget, QHBoxLayout,QVBoxLayout,QListWidget,QSplitter
from PyQt5.QtCore import Qt,QSize
import sys

class CellDiffWidget(QWidget):
    def __init__(self):
        super(CellDiffWidget,self).__init__()

        self.CellDiffListWidget = QListWidget()
        self.MergeDiffListWidget = QListWidget()

        self.MainLayout = QHBoxLayout()
        self.CellTabWidget = QTabWidget()

        self.CellTabWidget.addTab(self.CellDiffListWidget,"单元格改动")
        self.CellTabWidget.addTab(self.MergeDiffListWidget,"合并区域改动")
        self.MainLayout.addWidget(self.CellTabWidget)
        self.setLayout(self.MainLayout)
    
    def sizeHint(self):
        return QSize(100,200)

    def setData(self,data):

        self.CellDiffListWidget.clear()
        self.MergeDiffListWidget.clear()

        for nm in data["new_merge"]:
            self.MergeDiffListWidget.addItem("新增区域 -> "+str(nm))
        for nm in data["del_merge"]:
            self.MergeDiffListWidget.addItem("删除区域 -> "+str(nm))
        for diff in data["change_cell"]:
            self.CellDiffListWidget.addItem(str(diff[0]) + " -- " + str(diff[2][0]) +"\n         改为\n" + str(diff[1]) + " -- " + str(diff[2][1]))

 
class RowDiffWidget(QWidget):
    def __init__(self):
        super(RowDiffWidget,self).__init__()
        self.RowDelListWidget = QListWidget()
        self.RowAddListWidget = QListWidget()

        self.MainLayout = QHBoxLayout()
        self.RowTabWidget = QTabWidget()

        self.RowTabWidget.addTab(self.RowAddListWidget,"行增加")
        self.RowTabWidget.addTab(self.RowDelListWidget,"行删除")

        self.MainLayout.addWidget(self.RowTabWidget)
        self.setLayout(self.MainLayout)

    def sizeHint(self):
        return QSize(100,200)

    def setData(self,data):
        self.RowDelListWidget.clear()
        self.RowAddListWidget.clear()

        for row in data["add_row"]:
            self.RowAddListWidget.addItem("新增行 -> "+str(row))
        for row in data["del_row"]:
            self.RowDelListWidget.addItem("删除行 -> "+str(row))
        
class ColDiffWidget(QWidget):
    def __init__(self):
        super(ColDiffWidget,self).__init__()
        self.ColDelListWidget = QListWidget()
        self.ColAddListWidget = QListWidget()

        self.MainLayout = QHBoxLayout()
        self.ColTabWidget = QTabWidget()

        self.ColTabWidget.addTab(self.ColAddListWidget,"列增加")
        self.ColTabWidget.addTab(self.ColDelListWidget,"列删除")

        self.MainLayout.addWidget(self.ColTabWidget)
        self.setLayout(self.MainLayout)

    def sizeHint(self):
        return QSize(100,200)

    def setData(self,data):
        self.ColDelListWidget.clear()
        self.ColAddListWidget.clear()

        for col in data["add_col"]:
            self.ColAddListWidget.addItem("新增列 -> "+str(col))
        for col in data["del_col"]:
            self.ColDelListWidget.addItem("删除列 -> "+str(col))
        
if __name__=="__main__":
    app = QApplication(sys.argv)
    main = RowDiffWidget()
    main.show()
    sys.exit(app.exec_())