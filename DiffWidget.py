#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget, QTabWidget, QHBoxLayout,QVBoxLayout,QListWidget,QSplitter
from PyQt5.QtCore import Qt,QSize
import sys

class CellDiffWidget(QWidget):
    def __init__(self):
        super(CellDiffWidget,self).__init__()

        self.CellDiffListWidget = QListWidget()
        self.MergeAddDiffListWidget = QListWidget()
        self.MergeDelDiffListWidget = QListWidget()

        self.MainLayout = QHBoxLayout()
        self.CellTabWidget = QTabWidget()

        self.CellTabWidget.addTab(self.CellDiffListWidget,"单元格改动")
        self.CellTabWidget.addTab(self.MergeAddDiffListWidget,"合并区域增加")
        self.CellTabWidget.addTab(self.MergeDelDiffListWidget,"合并区域删除")
        self.MainLayout.addWidget(self.CellTabWidget)
        self.setLayout(self.MainLayout)
    
    def sizeHint(self):
        return QSize(100,200)

    def clear(self):
        self.CellDiffListWidget.clear()
        self.MergeAddDiffListWidget.clear()
        self.MergeDelDiffListWidget.clear()

    def setData(self,data):
        self.clear()

        for nm in data["new_merge"]:
            self.MergeAddDiffListWidget.addItem("新增区域 -> "+str(nm))
        for nm in data["del_merge"]:
            self.MergeDelDiffListWidget.addItem("删除区域 -> "+str(nm))
        for diff in data["change_cell"]:
            self.CellDiffListWidget.addItem(str(diff[0]) + " -- " + str(diff[2][0]) +"\n" + str(diff[1]) + " -- " + str(diff[2][1]))

 
class RowDiffWidget(QWidget):
    def __init__(self):
        super(RowDiffWidget,self).__init__()
        self.RowDelListWidget = QListWidget()
        self.RowAddListWidget = QListWidget()
        self.RowExcListWidget = QListWidget()

        self.MainLayout = QHBoxLayout()
        self.RowTabWidget = QTabWidget()

        self.RowTabWidget.addTab(self.RowAddListWidget,"行增加")
        self.RowTabWidget.addTab(self.RowDelListWidget,"行删除")
        self.RowTabWidget.addTab(self.RowExcListWidget,"行交换")

        self.MainLayout.addWidget(self.RowTabWidget)
        self.setLayout(self.MainLayout)

    def sizeHint(self):
        return QSize(100,200)

    def clear(self):
        self.RowDelListWidget.clear()
        self.RowAddListWidget.clear()
        self.RowExcListWidget.clear()

    def setData(self,data):
        self.clear()

        for row in data["add_row"]:
            self.RowAddListWidget.addItem("新增行 -> "+str(row))
        for row in data["del_row"]:
            self.RowDelListWidget.addItem("删除行 -> "+str(row))
        for row in data["row_exchange"]:
            self.RowExcListWidget.addItem("交换行 -> "+str(row[0]) +" -- " +str(row[1]))
        
class ColDiffWidget(QWidget):
    def __init__(self):
        super(ColDiffWidget,self).__init__()
        self.ColDelListWidget = QListWidget()
        self.ColAddListWidget = QListWidget()
        self.ColExcListWidget = QListWidget()

        self.MainLayout = QHBoxLayout()
        self.ColTabWidget = QTabWidget()

        self.ColTabWidget.addTab(self.ColAddListWidget,"列增加")
        self.ColTabWidget.addTab(self.ColDelListWidget,"列删除")
        self.ColTabWidget.addTab(self.ColExcListWidget,"列交换")

        self.MainLayout.addWidget(self.ColTabWidget)
        self.setLayout(self.MainLayout)

    def sizeHint(self):
        return QSize(100,200)

    def clear(self):
        self.ColDelListWidget.clear()
        self.ColAddListWidget.clear()
        self.ColExcListWidget.clear()

    def setData(self,data):
        self.clear()

        for col in data["add_col"]:
            self.ColAddListWidget.addItem("新增列 -> "+str(col))
        for col in data["del_col"]:
            self.ColDelListWidget.addItem("删除列 -> "+str(col))
        for col in data["col_exchange"]:
            self.ColExcListWidget.addItem("列交换 -> "+str(col[0])+" -- "+str(col[1]))
        
if __name__=="__main__":
    app = QApplication(sys.argv)
    main = RowDiffWidget()
    main.show()
    sys.exit(app.exec_())