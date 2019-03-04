#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QMainWindow,QApplication,QMenu,QAction,QFileDialog,QDockWidget
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from ViewWidget import ViewWidget
from DiffWidget import DiffWidget
from Model import MyXlsx
from Algorithm import MyAlg
import sys,math,hashlib
 
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()

        self.CentralWidget = ViewWidget()
        self.setCentralWidget(self.CentralWidget)
        self.DiffWidget = DiffWidget()

        self.DiffDock = QDockWidget("改动区")  # 实例化dockwidget类
        self.DiffDock.setWidget(self.DiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.DiffDock.setFeatures(QDockWidget.DockWidgetFloatable|QDockWidget.DockWidgetMovable)    #  设置dockwidget的各类属性
        self.addDockWidget((Qt.LeftDockWidgetArea or Qt.RightDockWidgetArea or Qt.TopDockWidgetArea or Qt.BottomDockWidgetArea), self.DiffDock)

        self.statusBar()
        self.initMenu()
        self.initToolBar()
        self.Alg = MyAlg()

    def initToolBar(self):
        anaAct = QAction(QIcon('ana.ico'), 'Exit', self)
        anaAct.setShortcut('Ctrl+A')
        anaAct.setToolTip("开始比较")
        anaAct.triggered.connect(self.beginAna)
       
        self.toolbar = self.addToolBar('ana')
        self.toolbar.addAction(anaAct)

    def initMenu(self):
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('文件')
        openMenu = QMenu('打开', self)

        openOldAct = QAction('打开旧文件', self)
        openOldAct.setStatusTip("打开旧版文件用于比较")
        openOldAct.triggered.connect(self.openOldFile)
        
        openNewAct = QAction('打开新文件', self)  
        openNewAct.setStatusTip("打开新的文件用于比较")
        openNewAct.triggered.connect(self.openNewFile)

        openMenu.addAction(openOldAct)
        openMenu.addAction(openNewAct)
        fileMenu.addMenu(openMenu)



    def openOldFile(self):
        fname = QFileDialog.getOpenFileName(self, '打开旧文件')
        if fname[0] != '':
            self.OldXlsx = MyXlsx(fname[0])
            self.CentralWidget.setOldTable(self.OldXlsx.SheetDatas)

    def openNewFile(self):
        fname = QFileDialog.getOpenFileName(self, '打开新文件')
        if fname[0] != '':
            self.NewXlsx = MyXlsx(fname[0])
            self.CentralWidget.setNewTable(self.NewXlsx.SheetDatas)

    def beginAna(self):
        oi = self.CentralWidget.OldTableWidget.currentIndex()
        ni = self.CentralWidget.NewTableWidget.currentIndex()

        self.Alg.setOldData(self.OldXlsx.SheetDatas[oi])
        self.Alg.setNewData(self.NewXlsx.SheetDatas[ni])
        self.SheetDiff = self.Alg.getSheetdiff()
        print(self.SheetDiff)
        self.DiffWidget.setData(self.SheetDiff)

if __name__=="__main__":

    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())