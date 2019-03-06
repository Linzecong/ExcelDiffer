#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget, QTabWidget, QHBoxLayout,QVBoxLayout,QListWidget,QSplitter,QListWidgetItem
from PyQt5.QtCore import Qt,QSize,QSettings
from PyQt5.QtGui import QColor
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
        return QSize(200,300)

    def clear(self):
        self.CellDiffListWidget.clear()
        self.MergeAddDiffListWidget.clear()
        self.MergeDelDiffListWidget.clear()

    def setData(self,data):
        self.clear()
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        add = self.ColorSettings.value("add");
        delcolor = self.ColorSettings.value("delcolor");
        change = self.ColorSettings.value("change");
        for nm in data["new_merge"]:
            a = QListWidgetItem("新增区域 -> "+str(nm))
            a.setForeground(QColor(add))
            self.MergeAddDiffListWidget.addItem(a)
        for nm in data["del_merge"]:
            a = QListWidgetItem("删除区域 -> "+str(nm))
            a.setForeground(QColor(delcolor))
            self.MergeDelDiffListWidget.addItem(a)
        for diff in data["change_cell"]:
            a = QListWidgetItem(str(diff[0]) + " -- " + str(diff[2][0]) +"\n" + str(diff[1]) + " -- " + str(diff[2][1]))
            a.setForeground(QColor(change))
            self.CellDiffListWidget.addItem(a)

 
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
        return QSize(200,300)

    def clear(self):
        self.RowDelListWidget.clear()
        self.RowAddListWidget.clear()
        self.RowExcListWidget.clear()

    def setData(self,data):
        self.clear()
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        add = self.ColorSettings.value("add");
        delcolor = self.ColorSettings.value("delcolor");
        change = self.ColorSettings.value("change");
        
        for row in data["add_row"]:
            a = QListWidgetItem("新增行 -> "+str(row))
            a.setForeground(QColor(add))
            self.RowAddListWidget.addItem(a)
        for row in data["del_row"]:
            a = QListWidgetItem("删除行 -> "+str(row))
            a.setForeground(QColor(delcolor))
            self.RowDelListWidget.addItem(a)
        for row in data["row_exchange"]:
            a = QListWidgetItem("交换行 -> "+str(row[0]) +" -- " +str(row[1]))
            a.setForeground(QColor(change))
            self.RowExcListWidget.addItem(a)
        
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
        return QSize(200,300)

    def clear(self):
        self.ColDelListWidget.clear()
        self.ColAddListWidget.clear()
        self.ColExcListWidget.clear()

    def setData(self,data):
        self.clear()
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        add = self.ColorSettings.value("add");
        delcolor = self.ColorSettings.value("delcolor");
        change = self.ColorSettings.value("change");
        for col in data["add_col"]:
            a = QListWidgetItem("新增列 -> "+str(col))
            a.setForeground(QColor(add))
            self.ColAddListWidget.addItem(a)
        for col in data["del_col"]:
            a = QListWidgetItem("删除列 -> "+str(col))
            a.setForeground(QColor(delcolor))
            self.ColDelListWidget.addItem(a)
        for col in data["col_exchange"]:
            a = QListWidgetItem("交换列 -> "+str(col[0])+" -- "+str(col[1]))
            a.setForeground(QColor(change))
            self.ColExcListWidget.addItem(a)
        
if __name__=="__main__":
    app = QApplication(sys.argv)
    main = RowDiffWidget()
    main.show()
    sys.exit(app.exec_())