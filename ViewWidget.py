#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QScrollBar, QWidget,QAction, QSplitter, QMainWindow, QApplication, QTableWidgetItem,QTableWidget,QHBoxLayout,QVBoxLayout
from PyQt5.QtGui import QBrush,QColor,QIcon
from PyQt5.QtCore import Qt,QSettings
import sys
from ExcelWidget import ExcelWidget
 
class ViewWidget(QMainWindow):
    def __init__(self):
        super(ViewWidget,self).__init__()
        self.diff = -1
        self.OldTableWidget = ExcelWidget()
        self.NewTableWidget = ExcelWidget()
        self.MainLayout = QHBoxLayout()
        self.Splitter = QSplitter(Qt.Horizontal)
        self.Splitter.addWidget(self.OldTableWidget)
        self.Splitter.addWidget(self.NewTableWidget)
        self.Splitter.setContentsMargins(5,5,5,5)
        self.setCentralWidget(self.Splitter)
        self.Lock = True
        self.OldTableWidget.currentChanged.connect(lambda x:self.setSame(x,0))
        self.NewTableWidget.currentChanged.connect(lambda x:self.setSame(x,1))

        self.OldTableWidget.cellClicked.connect(lambda x,y:self.setSameCell(x,y,0))
        self.NewTableWidget.cellClicked.connect(lambda x,y:self.setSameCell(x,y,1))
        
        self.OldTableWidget.hbarchange.connect(lambda x:self.NewTableWidget.TableWidgets[self.NewTableWidget.currentIndex()].horizontalScrollBar().setValue(x))
        self.NewTableWidget.vbarchange.connect(lambda x:self.OldTableWidget.TableWidgets[self.OldTableWidget.currentIndex()].verticalScrollBar().setValue(x))

        self.NewTableWidget.hbarchange.connect(lambda x:self.OldTableWidget.TableWidgets[self.OldTableWidget.currentIndex()].horizontalScrollBar().setValue(x))
        self.OldTableWidget.vbarchange.connect(lambda x:self.NewTableWidget.TableWidgets[self.NewTableWidget.currentIndex()].verticalScrollBar().setValue(x))
 
        self.initAction()
        self.initToolbar()
        
        

        # self.MainLayout.addWidget(self.Splitter)
        # self.setLayout(self.MainLayout)

    def setSameCell(self,x,y,type1):
        if self.Lock == False:
            return
        if type1 == 0:
            self.NewTableWidget.currentWidget().setCurrentCell(x,y)
        else:
            self.OldTableWidget.currentWidget().setCurrentCell(x,y)

    def setSame(self,id,type1):
        if self.Lock == False:
            return
        if type1 == 0:
            text = self.OldTableWidget.tabText(id)
            for i in range(self.NewTableWidget.count()):
                if text == self.NewTableWidget.tabText(i):
                    self.NewTableWidget.setCurrentIndex(i)
        else:
            text = self.NewTableWidget.tabText(id)
            for i in range(self.OldTableWidget.count()):
                if text == self.OldTableWidget.tabText(i):
                    self.OldTableWidget.setCurrentIndex(i)
    
    def initToolbar(self):
        self.toolbar = self.addToolBar("tabletool")
        

    def initAction(self):
        self.LockAction = QAction(QIcon("icon/lock.png"),"锁定",self)
        self.LockAction.setStatusTip("锁定表格，使得切换标签页时，新旧两个表格同步，且比较时将比较整个文件！")
        self.LockAction.triggered.connect(self.lockTab)

        self.UnlockAction = QAction(QIcon("icon/unlock.png"),"解锁",self)
        self.UnlockAction.setStatusTip("解锁表格，使得切换标签页时，新旧两个表格不会同步，且只比较选定的标签！")
        self.UnlockAction.triggered.connect(self.unlockTab)

    def lockTab(self):
        self.Lock = True
        self.toolbar.removeAction(self.LockAction)
        self.toolbar.addAction(self.UnlockAction)
    
    def unlockTab(self):
        self.Lock = False
        self.toolbar.removeAction(self.UnlockAction)
        self.toolbar.addAction(self.LockAction)
    
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
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        hightlight = self.ColorSettings.value("hightlight")

        self.setColor(self.diff,self.oi,self.ni)
        if widget == 0:
            if difftype == "del_col":
                col = self.ABCToInt(self.diff[difftype][id])
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(0,col-1)
                for i in range(self.OldTableWidget.TableWidgets[self.oi].rowCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i,col-1).setBackground(QBrush(QColor(hightlight)))

            if difftype == "del_row":
                row = self.diff[difftype][id]
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(row-1,0)
                for j in range(self.OldTableWidget.TableWidgets[self.oi].columnCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(row-1,j).setBackground(QBrush(QColor(hightlight)))
            if difftype == "change_cell":
                rec = self.diff[difftype][id]
                j = self.ABCToInt(rec[0][1])
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(rec[0][0]-1,j-1)
                self.OldTableWidget.TableWidgets[self.oi].item(rec[0][0]-1,j-1).setBackground(QBrush(QColor(hightlight)))
                j = self.ABCToInt(rec[1][1])
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(rec[1][0]-1,j-1)
                self.NewTableWidget.TableWidgets[self.ni].item(rec[1][0]-1,j-1).setBackground(QBrush(QColor(hightlight)))
            
            if difftype == "del_merge":
                rec = self.diff["del_merge"][id]
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(rec[0],rec[2])
                for i in range(rec[0],rec[1]):
                    for j in range(rec[2],rec[3]):
                        self.OldTableWidget.TableWidgets[self.oi].item(i,j).setBackground(QBrush(QColor(hightlight)))
            
            if difftype == "row_exchange":
                i = self.diff["row_exchange"][id]
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(i[0]-1,0)
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(i[1]-1,0)
                for j in range(self.OldTableWidget.TableWidgets[self.oi].columnCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i[0]-1,j).setBackground(QBrush(QColor(hightlight)))
                for j in range(self.NewTableWidget.TableWidgets[self.ni].columnCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(i[1]-1,j).setBackground(QBrush(QColor(hightlight)))
            
            if difftype == "col_exchange":
                s = self.diff["col_exchange"][id]
                j1 = self.ABCToInt(s[0])
                j2 = self.ABCToInt(s[1])
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(0,j1-1)
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(0,j2-1)
                for i in range(self.OldTableWidget.TableWidgets[self.oi].rowCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i,j1-1).setBackground(QBrush(QColor(hightlight)))
                for i in range(self.NewTableWidget.TableWidgets[self.ni].rowCount()):    
                    self.NewTableWidget.TableWidgets[self.ni].item(i,j2-1).setBackground(QBrush(QColor(hightlight)))

        elif widget == 1:
            if difftype == "add_col":
                col = self.ABCToInt(self.diff[difftype][id])
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(0,col-1)
                for i in range(self.NewTableWidget.TableWidgets[self.ni].rowCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(i,col-1).setBackground(QBrush(QColor(hightlight)))
            if difftype == "add_row":
                row = self.diff[difftype][id]
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(row-1,0)
                for j in range(self.NewTableWidget.TableWidgets[self.ni].columnCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(row-1,j).setBackground(QBrush(QColor(hightlight)))
            if difftype == "change_cell":
                rec = self.diff[difftype][id]
                j = self.ABCToInt(rec[0][1])
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(rec[0][0]-1,j-1)
                self.OldTableWidget.TableWidgets[self.oi].item(rec[0][0]-1,j-1).setBackground(QBrush(QColor(hightlight)))
                j = self.ABCToInt(rec[1][1])
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(rec[1][0]-1,j-1)
                self.NewTableWidget.TableWidgets[self.ni].item(rec[1][0]-1,j-1).setBackground(QBrush(QColor(hightlight)))
            if difftype == "new_merge":
                rec = self.diff["new_merge"][id]
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(rec[0],rec[2])
                for i in range(rec[0],rec[1]):
                    for j in range(rec[2],rec[3]):
                        self.NewTableWidget.TableWidgets[self.ni].item(i,j).setBackground(QBrush(QColor(hightlight)))
            if difftype == "row_exchange":
                i = self.diff["row_exchange"][id]
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(i[0]-1,0)
                for j in range(self.OldTableWidget.TableWidgets[self.oi].columnCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i[0]-1,j).setBackground(QBrush(QColor(hightlight)))
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(i[1]-1,0)
                for j in range(self.NewTableWidget.TableWidgets[self.ni].columnCount()):
                    self.NewTableWidget.TableWidgets[self.ni].item(i[1]-1,j).setBackground(QBrush(QColor(hightlight)))
            if difftype == "col_exchange":
                s = self.diff["col_exchange"][id]
                j1 = self.ABCToInt(s[0])
                j2 = self.ABCToInt(s[1])
                self.OldTableWidget.TableWidgets[self.oi].setCurrentCell(0,j1-1)
                for i in range(self.OldTableWidget.TableWidgets[self.oi].rowCount()):
                    self.OldTableWidget.TableWidgets[self.oi].item(i,j1-1).setBackground(QBrush(QColor(hightlight)))
                self.NewTableWidget.TableWidgets[self.ni].setCurrentCell(0,j2-1)    
                for i in range(self.NewTableWidget.TableWidgets[self.ni].rowCount()):    
                    self.NewTableWidget.TableWidgets[self.ni].item(i,j2-1).setBackground(QBrush(QColor(hightlight)))
        else:
            pass

    def setColor(self,diff,oi=-1,ni=-1):
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        hightlight = self.ColorSettings.value("hightlight")
        background = self.ColorSettings.value("background");
        exchange = self.ColorSettings.value("exchange");
        add = self.ColorSettings.value("add");
        delcolor = self.ColorSettings.value("delcolor");
        change = self.ColorSettings.value("change");

        self.diff = diff
        if oi==-1:
            oi = self.OldTableWidget.currentIndex()
            ni = self.NewTableWidget.currentIndex()
        self.oi=oi
        self.ni=ni

        for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i,j).setBackground(QBrush(QColor(background)))
                self.NewTableWidget.TableWidgets[ni].item(i,j).setForeground(QBrush(QColor("#000000")))
        for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j).setBackground(QBrush(QColor(background)))
                self.OldTableWidget.TableWidgets[oi].item(i,j).setForeground(QBrush(QColor("#000000")))

        for i in diff["row_exchange"]:
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i[0]-1,j).setBackground(QBrush(QColor(exchange)))
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i[1]-1,j).setBackground(QBrush(QColor(exchange)))

        for s in diff["col_exchange"]:
            j1 = self.ABCToInt(s[0])
            j2 = self.ABCToInt(s[1])
            for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j1-1).setBackground(QBrush(QColor(exchange)))
            for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):    
                self.NewTableWidget.TableWidgets[ni].item(i,j2-1).setBackground(QBrush(QColor(exchange)))

        for rec in diff["new_merge"]:
            for i in range(rec[0],rec[1]):
                for j in range(rec[2],rec[3]):
                    self.NewTableWidget.TableWidgets[ni].item(i,j).setBackground(QBrush(QColor(add)))

        for rec in diff["del_merge"]:
            for i in range(rec[0],rec[1]):
                for j in range(rec[2],rec[3]):
                    self.OldTableWidget.TableWidgets[oi].item(i,j).setBackground(QBrush(QColor(delcolor)))

        for s in diff["add_col"]:
            j = self.ABCToInt(s)
            for i in range(self.NewTableWidget.TableWidgets[ni].rowCount()):
                self.NewTableWidget.TableWidgets[ni].item(i,j-1).setBackground(QBrush(QColor(add)))

        for s in diff["del_col"]:
            j = self.ABCToInt(s)
            for i in range(self.OldTableWidget.TableWidgets[oi].rowCount()):
                self.OldTableWidget.TableWidgets[oi].item(i,j-1).setBackground(QBrush(QColor(delcolor)))

        for i in diff["add_row"]:
            for j in range(self.NewTableWidget.TableWidgets[ni].columnCount()):
                self.NewTableWidget.TableWidgets[ni].item(i-1,j).setBackground(QBrush(QColor(add)))

        for i in diff["del_row"]:
            for j in range(self.OldTableWidget.TableWidgets[oi].columnCount()):
                self.OldTableWidget.TableWidgets[oi].item(i-1,j).setBackground(QBrush(QColor(delcolor)))

        for rec in diff["change_cell"]:
            j = self.ABCToInt(rec[0][1])
            self.OldTableWidget.TableWidgets[oi].item(rec[0][0]-1,j-1).setBackground(QBrush(QColor(change)))
            j = self.ABCToInt(rec[1][1])
            self.NewTableWidget.TableWidgets[ni].item(rec[1][0]-1,j-1).setBackground(QBrush(QColor(change)))


if __name__=="__main__":
    app = QApplication(sys.argv)
    main = ViewWidget()
    main.show()
    sys.exit(app.exec_())