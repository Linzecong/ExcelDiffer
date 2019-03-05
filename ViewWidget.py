#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget,QSplitter, QMainWindow, QApplication, QTableWidgetItem,QTableWidget,QHBoxLayout,QVBoxLayout
from PyQt5.QtGui import QBrush,QColor
from PyQt5.QtCore import Qt
import sys
from ExcelWidget import ExcelWidget
 
class ViewWidget(QWidget):
    def __init__(self):
        super(ViewWidget,self).__init__()

        self.OldTableWidget = ExcelWidget()
        self.NewTableWidget = ExcelWidget()
        self.MainLayout = QHBoxLayout()
        self.Splitter = QSplitter(Qt.Horizontal)
        self.Splitter.addWidget(self.OldTableWidget)
        self.Splitter.addWidget(self.NewTableWidget)

        self.MainLayout.addWidget(self.Splitter)
        self.setLayout(self.MainLayout)
    
    def setOldTable(self,data):
        self.OldTableWidget.setData(data)
        
    def setNewTable(self,data):
        self.NewTableWidget.setData(data)

    def ABCToInt(self, s):
        dict0 = {}
        for i in range(26):
            dict0[chr(ord('A')+i)]=i+1
 
        output = 0
        for i in range(len(s)):
            output = output*26+dict0[s[i]]
        
        return output

    def setColor(self,diff):
        oi = self.OldTableWidget.currentIndex()
        ni = self.NewTableWidget.currentIndex()

        for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i,j).setBackground(QBrush(QColor(255, 255, 255)))
        for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j).setBackground(QBrush(QColor(255, 255, 255)))

        for s in diff["add_col"]:
            j = self.ABCToInt(s)
            for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):
                self.NewTableWidget.TableWidgets[ni].item(i,j-1).setBackground(QBrush(QColor("#409EFF")))

        for s in diff["del_col"]:
            j = self.ABCToInt(s)
            for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j-1).setBackground(QBrush(QColor("#F56C6C")))

        for i in diff["add_row"]:
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i-1,j).setBackground(QBrush(QColor("#409EFF")))

        for i in diff["del_row"]:
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i-1,j).setBackground(QBrush(QColor("#F56C6C")))

        for rec in diff["change_cell"]:
            j = self.ABCToInt(rec[0][1])
            self.OldTableWidget.TableWidgets[oi].item(rec[0][0]-1,j-1).setBackground(QBrush(QColor("#E6A23C")))
            j = self.ABCToInt(rec[1][1])
            self.NewTableWidget.TableWidgets[ni].item(rec[1][0]-1,j-1).setBackground(QBrush(QColor("#E6A23C")))

        for rec in diff["new_merge"]:
            for i in range(rec[0],rec[1]):
                for j in range(rec[2],rec[3]):
                    self.NewTableWidget.TableWidgets[ni].item(i,j).setBackground(QBrush(QColor("#409EFF")))

        for rec in diff["del_merge"]:
            for i in range(rec[0],rec[1]):
                for j in range(rec[2],rec[3]):
                    self.OldTableWidget.TableWidgets[oi].item(i,j).setBackground(QBrush(QColor("#F56C6C")))


if __name__=="__main__":
    app = QApplication(sys.argv)
    main = ViewWidget()
    main.show()
    sys.exit(app.exec_())