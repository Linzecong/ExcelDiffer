#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget, QTabWidget, QTableWidgetItem,QTableWidget,QHBoxLayout,QVBoxLayout,QAbstractItemView
from PyQt5.QtGui import QFont
import sys
import xlrd
class ExcelWidget(QTabWidget):
    def __init__(self):
        super(ExcelWidget,self).__init__()
        self.TableWidgets = []

    def intToABC(self,n):
        d = {}
        r = []
        a = ''
        for i in range(1,27):
            d[i] = chr(64+i)
        if n <= 26:
            return d[n]
        if n % 26 == 0:
            n = n/26 - 1
            a ='Z'
        while n > 26:
            s = n % 26
            n = n // 26
            r.append(s)
        result = d[n]
        for i in r[::-1]:
            result+=d[i]
        return result + a


    def setData(self,data):
        self.clear()
        self.TableWidgets.clear()
        
        for sheet in data:
            tableWidget = QTableWidget()
            tableWidget.setStyleSheet("selection-background-color: #DCDFE6");
            hlable = []
            vlable = []
            for i in range(sheet["col"]):
                hlable.append(self.intToABC(i+1))
            for i in range(sheet["row"]):
                vlable.append(str(i+1))
            tableWidget.setRowCount(sheet["row"])
            tableWidget.setColumnCount(sheet["col"])
            tableWidget.setVerticalHeaderLabels(vlable)
            tableWidget.setHorizontalHeaderLabels(hlable)
            tableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
            

            for i in range(sheet["row"]):
                for j in range(sheet["col"]):
                    value = sheet["data"][i][j].value
                    if sheet["data"][i][j].ctype == 3:
                        value = xlrd.xldate.xldate_as_datetime(value,0).strftime('%Y/%m/%d %H:%M:%S')
                    tableWidget.setItem(i,j,QTableWidgetItem(str(value)))
            
            for span in sheet["merged"]:
                tableWidget.setSpan(span[0],span[2],span[1]-span[0],span[3]-span[2])
                
            QTableWidget.resizeColumnsToContents(tableWidget)
            QTableWidget.resizeRowsToContents(tableWidget)
            self.TableWidgets.append(tableWidget)
            self.addTab(tableWidget,sheet["name"])


if __name__=="__main__":
    app = QApplication(sys.argv)
    main = ViewWidget()
    main.show()
    sys.exit(app.exec_())