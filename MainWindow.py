#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QMainWindow,QApplication,QMenu,QAction,QFileDialog,QDockWidget,QMessageBox,QDesktopWidget
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt,QSettings
from ViewWidget import ViewWidget
from DiffWidget import RowDiffWidget,CellDiffWidget,ColDiffWidget
from Model import MyXlsx
from Algorithm import MyAlg
import sys,math,hashlib
 
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()

        self.CentralWidget = ViewWidget()
        self.setCentralWidget(self.CentralWidget)

        self.RowDiffWidget = RowDiffWidget()
        self.RowDiffDock = QDockWidget("行改动")  # 实例化dockwidget类
        self.RowDiffDock.setWidget(self.RowDiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.RowDiffDock.setFeatures(QDockWidget.AllDockWidgetFeatures)    #  设置dockwidget的各类属性
        self.RowDiffDock.setAllowedAreas(Qt.AllDockWidgetAreas)
        self.RowDiffDock.setObjectName("rowdiff")
        self.addDockWidget(Qt.LeftDockWidgetArea, self.RowDiffDock)
        

        self.ColDiffWidget = ColDiffWidget()
        self.ColDiffDock = QDockWidget("列改动")  # 实例化dockwidget类
        self.ColDiffDock.setWidget(self.ColDiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.ColDiffDock.setFeatures(QDockWidget.AllDockWidgetFeatures)    #  设置dockwidget的各类属性
        self.ColDiffDock.setAllowedAreas(Qt.AllDockWidgetAreas)
        self.ColDiffDock.setObjectName("coldiff")
        self.addDockWidget(Qt.LeftDockWidgetArea, self.ColDiffDock)

        self.CellDiffWidget = CellDiffWidget()
        self.CellDiffDock = QDockWidget("格改动")  # 实例化dockwidget类
        self.CellDiffDock.setWidget(self.CellDiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.CellDiffDock.setFeatures(QDockWidget.AllDockWidgetFeatures)    #  设置dockwidget的各类属性
        self.CellDiffDock.setAllowedAreas(Qt.AllDockWidgetAreas)
        self.CellDiffDock.setObjectName("celldiff")
        self.addDockWidget(Qt.RightDockWidgetArea, self.CellDiffDock)

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        self.Settings = QSettings("ExcelDiffer", "ExcelDifferInit");
        self.Settings.setValue("geometry", self.saveGeometry());
        self.Settings.setValue("windowState", self.saveState());
        

        self.statusBar()
        self.initAction()
        self.initMenu()
        self.initToolBar()
        self.Alg = MyAlg()
    
    def restoreToSetting(self):
        self.Settings = QSettings("ExcelDiffer", "ExcelDiffer");
        self.restoreGeometry(self.Settings.value("geometry"));
        self.restoreState(self.Settings.value("windowState"));

    def saveToSetting(self):
        self.Settings = QSettings("ExcelDiffer", "ExcelDiffer");
        self.Settings.setValue("geometry", self.saveGeometry());
        self.Settings.setValue("windowState", self.saveState());
    
    def restoreToInit(self):
        self.Settings = QSettings("ExcelDiffer", "ExcelDifferInit");
        self.restoreGeometry(self.Settings.value("geometry"));
        self.restoreState(self.Settings.value("windowState"));


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

        self.showRowDiffWinAct = QAction('打开行改动窗口', self)
        self.showRowDiffWinAct.setStatusTip("打开行改动窗口")
        self.showRowDiffWinAct.triggered.connect(self.RowDiffDock.show)
        self.showColDiffWinAct = QAction('打开列改动窗口', self)
        self.showColDiffWinAct.setStatusTip("打开列改动窗口")
        self.showColDiffWinAct.triggered.connect(self.ColDiffDock.show)
        self.showCellDiffWinAct = QAction('打开格改动窗口', self)
        self.showCellDiffWinAct.setStatusTip("打开格改动窗口")
        self.showCellDiffWinAct.triggered.connect(self.CellDiffDock.show)

        self.saveWinAct = QAction('保存当前窗口状态', self)
        self.saveWinAct.setStatusTip("保存当前窗口状态")
        self.saveWinAct.triggered.connect(self.saveToSetting)

        self.restoreWinAct = QAction('读取窗口状态', self)
        self.restoreWinAct.setStatusTip("读取窗口状态")
        self.restoreWinAct.triggered.connect(self.restoreToSetting)

        self.restoreInitWinAct = QAction('恢复默认窗口状态', self)
        self.restoreInitWinAct.setStatusTip("恢复默认窗口状态")
        self.restoreInitWinAct.triggered.connect(self.restoreToInit)


    def initToolBar(self):
        self.toolbar = self.addToolBar('tool')
        self.toolbar.setObjectName("toolbar")
        self.toolbar.addAction(self.anaAct)
        self.toolbar.addAction(self.openOldAct)
        self.toolbar.addAction(self.openNewAct)

    def initMenu(self):
        menubar = self.menuBar()

        fileMenu = menubar.addMenu('文件')
        anaMenu = menubar.addMenu('比较')
        winMenu = menubar.addMenu('窗口')

        openMenu = QMenu('打开', self)
        openMenu.addAction(self.openOldAct)
        openMenu.addAction(self.openNewAct)
        fileMenu.addMenu(openMenu)

        anaMenu.addAction(self.anaAct)

        winMenu.addAction(self.showRowDiffWinAct)
        winMenu.addAction(self.showColDiffWinAct)
        winMenu.addAction(self.showCellDiffWinAct)
        winMenu.addSeparator()
        winMenu.addAction(self.saveWinAct)
        winMenu.addAction(self.restoreWinAct)
        winMenu.addAction(self.restoreInitWinAct)


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
        self.RowDiffWidget.setData(self.SheetDiff)
        self.ColDiffWidget.setData(self.SheetDiff)
        self.CellDiffWidget.setData(self.SheetDiff)
        self.CentralWidget.setColor(self.SheetDiff)

if __name__=="__main__":

    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())