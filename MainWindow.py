#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QMainWindow,QApplication,QMenu,QAction,QFileDialog,QDockWidget,QMessageBox
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
        self.initAction()
        self.initMenu()
        self.initToolBar()
        self.Alg = MyAlg()

    def initAction(self):
        self.anaAct = QAction('开始比较', self)
        self.anaAct.setShortcut('Ctrl+A')
        self.anaAct.setStatusTip("开始比较")
        self.anaAct.triggered.connect(self.beginAna)
        
        self.openOldAct = QAction('打开旧文件', self)
        self.openOldAct.setShortcut('Ctrl+O')
        self.openOldAct.setStatusTip("打开旧版文件用于比较")
        self.openOldAct.triggered.connect(self.openOldFile)
        
        self.openNewAct = QAction('打开新文件', self)
        self.openNewAct.setShortcut('Ctrl+N')
        self.openNewAct.setStatusTip("打开新的文件用于比较")
        self.openNewAct.triggered.connect(self.openNewFile)


    def initToolBar(self):
        self.toolbar = self.addToolBar('tool')

        self.toolbar.addAction(self.anaAct)
        self.toolbar.addAction(self.openOldAct)
        self.toolbar.addAction(self.openNewAct)

    def initMenu(self):
        menubar = self.menuBar()

        fileMenu = menubar.addMenu('文件')
        anaMenu = menubar.addMenu('比较')

        openMenu = QMenu('打开', self)
        openMenu.addAction(self.openOldAct)
        openMenu.addAction(self.openNewAct)
        fileMenu.addMenu(openMenu)

        anaMenu.addAction(self.anaAct)
        



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
        if oi == -1 or ni ==-1:
            QMessageBox.warning(self, "提示", "请先打开Excel！", QMessageBox.Ok)
            return
        reply = QMessageBox.warning(self, "确定比较吗？", "将要比较 "+self.OldXlsx.SheetDatas[oi]["name"]+" 和 "+self.NewXlsx.SheetDatas[ni]["name"], QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
        if reply == QMessageBox.No:
            return
        self.Alg.setOldData(self.OldXlsx.SheetDatas[oi])
        self.Alg.setNewData(self.NewXlsx.SheetDatas[ni])
        self.SheetDiff = self.Alg.getSheetdiff()
        # print(self.SheetDiff)
        self.DiffWidget.setData(self.SheetDiff)
        self.CentralWidget.setColor(self.SheetDiff)

if __name__=="__main__":

    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())