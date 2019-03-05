#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget,QSplitter, QMainWindow, QApplication, QTableWidgetItem,QTableWidget,QHBoxLayout,QVBoxLayout
from PyQt5.QtGui import QBrush,QColor
from PyQt5.QtCore import Qt
import sys
from ExcelWidget import ExcelWidget
 
class ViewWidget(QWidget):
    def __init__(self):
        super(ViewWidget,self).__init__()
        self.diff = -1
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

    def setHighLight(self,widget,difftype,id):
        """
        0 old
        1 new
        2 both
        """
        self.setColor(self.diff,self.oi,self.ni)
        if widget == 0:
            if difftype == "del_col":
                col = self.ABCToInt(self.diff[difftype][id])
                for i in range(self.OldTableWidget.TableWidgets[self.oi].rowCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i,col-1).setBackground(QBrush(QColor("#909399")))
            if difftype == "del_row":
                row = self.diff[difftype][id]
                for j in range(self.OldTableWidget.TableWidgets[self.oi].columnCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(row-1,j).setBackground(QBrush(QColor("#909399")))
            if difftype == "change_cell":
                rec = self.diff[difftype][id]
                j = self.ABCToInt(rec[0][1])
                self.OldTableWidget.TableWidgets[self.oi].item(rec[0][0]-1,j-1).setBackground(QBrush(QColor("#909399")))
                j = self.ABCToInt(rec[1][1])
                self.NewTableWidget.TableWidgets[self.ni].item(rec[1][0]-1,j-1).setBackground(QBrush(QColor("#909399")))
            if difftype == "del_merge":
                rec = self.diff["del_merge"][id]
                for i in range(rec[0],rec[1]):
                    for j in range(rec[2],rec[3]):
                        self.OldTableWidget.TableWidgets[self.oi].item(i,j).setBackground(QBrush(QColor("#909399")))
            
            if difftype == "row_exchange":
                i = self.diff["row_exchange"][id]
                for j in range(self.OldTableWidget.TableWidgets[self.oi].columnCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i[0]-1,j).setBackground(QBrush(QColor("#909399")))
                for j in range(self.NewTableWidget.TableWidgets[self.ni].columnCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(i[1]-1,j).setBackground(QBrush(QColor("#909399")))
            
            if difftype == "col_exchange":
                s = self.diff["col_exchange"][id]
                j1 = self.ABCToInt(s[0])
                j2 = self.ABCToInt(s[1])
                for i in range(self.OldTableWidget.TableWidgets[self.oi].rowCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i,j1-1).setBackground(QBrush(QColor("#909399")))
                for i in range(self.NewTableWidget.TableWidgets[self.ni].rowCount()):    
                    self.NewTableWidget.TableWidgets[self.ni].item(i,j2-1).setBackground(QBrush(QColor("#909399")))

        elif widget == 1:
            if difftype == "add_col":
                col = self.ABCToInt(self.diff[difftype][id])
                for i in range(self.NewTableWidget.TableWidgets[self.ni].rowCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(i,col-1).setBackground(QBrush(QColor("#909399")))
            if difftype == "add_row":
                row = self.diff[difftype][id]
                for j in range(self.NewTableWidget.TableWidgets[self.ni].columnCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(row-1,j).setBackground(QBrush(QColor("#909399")))
            if difftype == "change_cell":
                rec = self.diff[difftype][id]
                j = self.ABCToInt(rec[0][1])
                self.OldTableWidget.TableWidgets[self.oi].item(rec[0][0]-1,j-1).setBackground(QBrush(QColor("#909399")))
                j = self.ABCToInt(rec[1][1])
                self.NewTableWidget.TableWidgets[self.ni].item(rec[1][0]-1,j-1).setBackground(QBrush(QColor("#909399")))
            if difftype == "new_merge":
                rec = self.diff["new_merge"][id]
                for i in range(rec[0],rec[1]):
                    for j in range(rec[2],rec[3]):
                        self.NewTableWidget.TableWidgets[self.ni].item(i,j).setBackground(QBrush(QColor("#909399")))
            if difftype == "row_exchange":
                i = self.diff["row_exchange"][id]
                for j in range(self.OldTableWidget.TableWidgets[self.oi].columnCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i[0]-1,j).setBackground(QBrush(QColor("#909399")))
                for j in range(self.NewTableWidget.TableWidgets[self.ni].columnCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(i[1]-1,j).setBackground(QBrush(QColor("#909399")))
            if difftype == "col_exchange":
                s = self.diff["col_exchange"][id]
                j1 = self.ABCToInt(s[0])
                j2 = self.ABCToInt(s[1])
                for i in range(self.OldTableWidget.TableWidgets[self.oi].rowCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i,j1-1).setBackground(QBrush(QColor("#909399")))
                for i in range(self.NewTableWidget.TableWidgets[self.ni].rowCount()):    
                    self.NewTableWidget.TableWidgets[self.ni].item(i,j2-1).setBackground(QBrush(QColor("#909399")))
        else:
            pass

    def _setColor(self,e):
        if self.diff == -1:
            return
        self.setColor(self.diff)

    def setColor(self,diff,oi=-1,ni=-1):
        self.diff = diff
        if oi==-1:
            oi = self.OldTableWidget.currentIndex()
            ni = self.NewTableWidget.currentIndex()
        self.oi=oi
        self.ni=ni

        for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i,j).setBackground(QBrush(QColor(255, 255, 255)))
        for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j).setBackground(QBrush(QColor(255, 255, 255)))

        for i in diff["row_exchange"]:
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i[0]-1,j).setBackground(QBrush(QColor("#EBEEF5")))
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i[1]-1,j).setBackground(QBrush(QColor("#EBEEF5")))

        for s in diff["col_exchange"]:
            j1 = self.ABCToInt(s[0])
            j2 = self.ABCToInt(s[1])
            for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j1-1).setBackground(QBrush(QColor("#EBEEF5")))
            for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):    
                self.NewTableWidget.TableWidgets[ni].item(i,j2-1).setBackground(QBrush(QColor("#EBEEF5")))

        for rec in diff["new_merge"]:
            for i in range(rec[0],rec[1]):
                for j in range(rec[2],rec[3]):
                    self.NewTableWidget.TableWidgets[ni].item(i,j).setBackground(QBrush(QColor("#409EFF")))

        for rec in diff["del_merge"]:
            for i in range(rec[0],rec[1]):
                for j in range(rec[2],rec[3]):
                    self.OldTableWidget.TableWidgets[oi].item(i,j).setBackground(QBrush(QColor("#F56C6C")))

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


if __name__=="__main__":
    app = QApplication(sys.argv)
    main = ViewWidget()
    main.show()
    sys.exit(app.exec_())