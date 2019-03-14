#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QMainWindow,QFontDialog, QApplication,QMenu,QAction,QFileDialog,QDockWidget,QMessageBox,QDesktopWidget,QTableWidget
from PyQt5.QtGui import QIcon,QFont,QKeySequence
from PyQt5.QtCore import Qt,QSettings,QThread,pyqtSignal
from ViewWidget import ViewWidget
from DiffWidget import RowDiffWidget,CellDiffWidget,ColDiffWidget
from ChangeColorWidget import ChangeColorWidget
from Model import MyXlsx
from Algorithm import MyAlg
import sys,math,hashlib

class Thread(QThread):

    done = pyqtSignal(dict)

    def __init__(self,Window):
        super().__init__()
        self.window = Window
        

    def run(self):
        self.SheetDiff = self.window.Alg.getSheetdiff()
        self.done.emit(self.SheetDiff)
        # print(self.SheetDiff)
        
 
class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        self.isanaing = False #是否在分析
        self.setWindowTitle("ExcelDiffer")
        self.qss = True
        self.IsAna = False
        self.currentFont = self.font()
        self.initFont = self.font()
        self.ps = self.initFont.pointSize()
        self.CentralWidget = ViewWidget()
        self.setCentralWidget(self.CentralWidget)

        self.RowDiffWidget = RowDiffWidget()
        self.RowDiffDock = QDockWidget("行改动")  # 实例化dockwidget类
        self.RowDiffDock.setWidget(self.RowDiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.RowDiffDock.setFeatures(QDockWidget.AllDockWidgetFeatures)    #  设置dockwidget的各类属性
        self.RowDiffDock.setAllowedAreas(Qt.AllDockWidgetAreas)
        self.RowDiffDock.setObjectName("rowdiff")
        self.RowDiffWidget.RowAddListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(1,"add_row",self.RowDiffWidget.RowAddListWidget.currentRow()))
        self.RowDiffWidget.RowDelListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(0,"del_row",self.RowDiffWidget.RowDelListWidget.currentRow()))
        self.RowDiffWidget.RowExcListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(0,"row_exchange",self.RowDiffWidget.RowExcListWidget.currentRow()))
        
        self.addDockWidget(Qt.LeftDockWidgetArea, self.RowDiffDock)

        
        

        self.ColDiffWidget = ColDiffWidget()
        self.ColDiffDock = QDockWidget("列改动")  # 实例化dockwidget类
        self.ColDiffDock.setWidget(self.ColDiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.ColDiffDock.setFeatures(QDockWidget.AllDockWidgetFeatures)    #  设置dockwidget的各类属性
        self.ColDiffDock.setAllowedAreas(Qt.AllDockWidgetAreas)
        self.ColDiffDock.setObjectName("coldiff")
        self.ColDiffWidget.ColAddListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(1,"add_col",self.ColDiffWidget.ColAddListWidget.currentRow()))
        self.ColDiffWidget.ColDelListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(0,"del_col",self.ColDiffWidget.ColDelListWidget.currentRow()))
        self.ColDiffWidget.ColExcListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(0,"col_exchange",self.ColDiffWidget.ColExcListWidget.currentRow()))
        self.addDockWidget(Qt.LeftDockWidgetArea, self.ColDiffDock)

        self.CellDiffWidget = CellDiffWidget()
        self.CellDiffDock = QDockWidget("格改动")  # 实例化dockwidget类
        self.CellDiffDock.setWidget(self.CellDiffWidget)   # 带入的参数为一个QWidget窗体实例，将该窗体放入dock中
        self.CellDiffDock.setFeatures(QDockWidget.AllDockWidgetFeatures)    #  设置dockwidget的各类属性
        self.CellDiffDock.setAllowedAreas(Qt.AllDockWidgetAreas)
        self.CellDiffDock.setObjectName("celldiff")
        self.CellDiffWidget.CellDiffListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(0,"change_cell",self.CellDiffWidget.CellDiffListWidget.currentRow()))
        self.CellDiffWidget.MergeAddDiffListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(1,"new_merge",self.CellDiffWidget.MergeAddDiffListWidget.currentRow()))
        self.CellDiffWidget.MergeDelDiffListWidget.clicked.connect(lambda:self.CentralWidget.setHighLight(0,"del_merge",self.CellDiffWidget.MergeDelDiffListWidget.currentRow()))
        self.addDockWidget(Qt.RightDockWidgetArea, self.CellDiffDock)

        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

        self.WinSettings = QSettings("ExcelDiffer", "ExcelDifferInit");
        self.WinSettings.setValue("geometry", self.saveGeometry());
        self.WinSettings.setValue("windowState", self.saveState());

        
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        if self.ColorSettings.contains("hightlight") == False:
            self.ColorSettings.setValue("hightlight", "#909399");
            self.ColorSettings.setValue("background", "#FFFFFF");
            self.ColorSettings.setValue("exchange", "#EBEEF5");
            self.ColorSettings.setValue("add", "#409EFF");
            self.ColorSettings.setValue("delcolor", "#F56C6C");
            self.ColorSettings.setValue("change", "#E6A23C");
        self.bar = self.statusBar()
        self.initAction()
        self.initMenu()
        self.initToolBar()
        self.Alg = MyAlg()
    
    def restoreToWinSetting(self):
        reply = QMessageBox.warning(self, "确定恢复上次保存的窗口布局吗？", "确定恢复上次保存的窗口布局吗？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
        if reply == QMessageBox.No:
            return
        self.WinSettings = QSettings("ExcelDiffer", "ExcelDiffer");
        self.restoreGeometry(self.WinSettings.value("geometry"));
        self.restoreState(self.WinSettings.value("windowState"));

    def saveToWinSetting(self):
        self.WinSettings = QSettings("ExcelDiffer", "ExcelDiffer");
        self.WinSettings.setValue("geometry", self.saveGeometry());
        self.WinSettings.setValue("windowState", self.saveState());
    
    def restoreToWinInit(self):
        reply = QMessageBox.warning(self, "确定恢复默认吗？", "确认将窗口恢复默认吗？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
        if reply == QMessageBox.No:
            return
        self.WinSettings = QSettings("ExcelDiffer", "ExcelDifferInit");
        self.restoreGeometry(self.WinSettings.value("geometry"));
        self.restoreState(self.WinSettings.value("windowState"));

    def chooseColor(self):
        self.CW = ChangeColorWidget()
        self.CW.exec_()
        
    
    def restoreToColorInit(self):
        reply = QMessageBox.warning(self, "确定恢复默认颜色吗？", "确定恢复默认颜色吗？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
        if reply == QMessageBox.No:
            return
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        self.ColorSettings.setValue("hightlight", "#909399");
        self.ColorSettings.setValue("background", "#FFFFFF");
        self.ColorSettings.setValue("exchange", "#EBEEF5");
        self.ColorSettings.setValue("add", "#409EFF");
        self.ColorSettings.setValue("delcolor", "#F56C6C");
        self.ColorSettings.setValue("change", "#E6A23C");

        QMessageBox.information(self, "提示", "恢复成功，请重新比较", QMessageBox.Yes);


    def initAction(self):
        self.anaAct = QAction(QIcon("icon/ana.png"),'开始比较', self)
        self.anaAct.setShortcut('Ctrl+A')
        self.anaAct.setStatusTip("开始比较")
        self.anaAct.triggered.connect(self.beginAna)

        self.openOldAct = QAction(QIcon("icon/openold.ico"),'打开旧文件', self)
        self.openOldAct.setShortcut('Ctrl+O')
        self.openOldAct.setStatusTip("打开旧版文件用于比较")
        self.openOldAct.triggered.connect(self.openOldFile)
        
        self.openNewAct = QAction(QIcon("icon/opennew.ico"),'打开新文件', self)
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
        self.saveWinAct.triggered.connect(self.saveToWinSetting)

        self.restoreWinAct = QAction('读取窗口状态', self)
        self.restoreWinAct.setStatusTip("读取窗口状态")
        self.restoreWinAct.triggered.connect(self.restoreToWinSetting)

        self.restoreInitWinAct = QAction('恢复默认窗口状态', self)
        self.restoreInitWinAct.setStatusTip("恢复默认窗口状态")
        self.restoreInitWinAct.triggered.connect(self.restoreToWinInit)


        self.zoomInAct = QAction(QIcon("icon/zoom-in.png"),'增加字体大小', self)
        self.zoomInAct.setShortcut(QKeySequence.ZoomIn)
        self.zoomInAct.setStatusTip("增加字体大小")
        self.zoomInAct.triggered.connect(self.zoomIn)
        
        self.zoomOutAct = QAction(QIcon("icon/zoom-out.png"),'减少字体大小', self)
        self.zoomOutAct.setShortcut(QKeySequence.ZoomOut)
        self.zoomOutAct.setStatusTip("减少字体大小")
        self.zoomOutAct.triggered.connect(self.zoomOut)

        self.chooseFontAct = QAction(QIcon("icon/chafont.png"),'选择表格字体', self)
        self.chooseFontAct.setStatusTip("选择表格字体")
        self.chooseFontAct.triggered.connect(self.chooseFont)

        self.restoreFontAct = QAction(QIcon("icon/refont.png"),'恢复默认字体', self)
        self.restoreFontAct.setStatusTip("恢复默认字体")
        self.restoreFontAct.triggered.connect(lambda : self.initFont.setPointSize(self.ps) or self.setTableFont(self.initFont))

        self.chooseColorAct = QAction(QIcon("icon/chacol.png"),'更改颜色', self)
        self.chooseColorAct.setStatusTip("更改表格颜色设置")
        self.chooseColorAct.triggered.connect(self.chooseColor)

        self.restoreColorAct = QAction(QIcon("icon/recol.png"),'恢复默认颜色', self)
        self.restoreColorAct.setStatusTip("更改表格颜色设置")
        self.restoreColorAct.triggered.connect(self.restoreToColorInit)

        self.qssAct = QAction('切换背景', self)
        self.qssAct.setStatusTip("在白天模式和夜间模式切换")
        self.qssAct.triggered.connect(self.changeQSS)


    def changeQSS(self):
        if self.qss == True:
            f = open('style.qss', 'r')
            self.setStyleSheet(f.read())
            self.qss = False
        else:
            self.setStyleSheet("")
            self.qss = True

    def chooseFont(self):
        a = QFontDialog.getFont(self.currentFont,self)
        if a[1]:
            self.setTableFont(a[0])

    def setTableFont(self,font):
        self.currentFont = font
        for widget in self.CentralWidget.OldTableWidget.TableWidgets:
            widget.setFont(font)
            widget.verticalHeader().setFont(font)
            widget.horizontalHeader().setFont(font)
            QTableWidget.resizeColumnsToContents(widget)
            QTableWidget.resizeRowsToContents(widget)

        for widget in self.CentralWidget.NewTableWidget.TableWidgets:
            widget.setFont(font)
            widget.verticalHeader().setFont(font)
            widget.horizontalHeader().setFont(font)
            QTableWidget.resizeColumnsToContents(widget)
            QTableWidget.resizeRowsToContents(widget)
    
    def zoomIn(self):
        font = self.currentFont
        size = font.pointSize()+1
        if size == 200: size = size - 1
        font.setPointSize(size)
        self.setTableFont(font)
    
    def zoomOut(self):
        font = self.currentFont
        size = font.pointSize()-1
        if size == 1: size = size + 1
        font.setPointSize(size)
        self.setTableFont(font)

    def initToolBar(self):
        self.toolbar = self.addToolBar('tool')
        self.toolbar.setObjectName("toolbar")
        self.toolbar.addAction(self.anaAct)
        self.toolbar.addSeparator()
        self.toolbar.addAction(self.openOldAct)
        self.toolbar.addAction(self.openNewAct)

        self.CentralWidget.toolbar.addAction(self.zoomInAct)
        self.CentralWidget.toolbar.addAction(self.zoomOutAct)
        self.CentralWidget.toolbar.addAction(self.chooseFontAct)
        self.CentralWidget.toolbar.addAction(self.restoreFontAct)
        self.CentralWidget.toolbar.addSeparator()
        self.CentralWidget.toolbar.addAction(self.chooseColorAct)
        self.CentralWidget.toolbar.addAction(self.restoreColorAct)
        self.CentralWidget.toolbar.addSeparator()
        self.CentralWidget.toolbar.addAction(self.CentralWidget.UnlockAction)

    def initMenu(self):
        menubar = self.menuBar()

        fileMenu = menubar.addMenu('文件')
        anaMenu = menubar.addMenu('比较')
        winMenu = menubar.addMenu('窗口')
        formatMenu = menubar.addMenu('格式')

        nightMenu = menubar.addMenu('切换模式')
        nightMenu.addAction(self.qssAct)

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

        formatMenu.addAction(self.zoomInAct)
        formatMenu.addAction(self.zoomOutAct)
        formatMenu.addSeparator()
        formatMenu.addAction(self.chooseFontAct)
        formatMenu.addAction(self.restoreFontAct)
        formatMenu.addSeparator()
        formatMenu.addAction(self.chooseColorAct)
        formatMenu.addAction(self.restoreColorAct)

    def openOldFile(self):
        fname = QFileDialog.getOpenFileName(self, '打开旧文件')
        if fname[0] != '':
            self.OldXlsx = MyXlsx(fname[0])
            self.CentralWidget.setOldTable(self.OldXlsx.SheetDatas)
            self.CellDiffWidget.clear()
            self.RowDiffWidget.clear()
            self.ColDiffWidget.clear()
            self.IsAna = False

    def openNewFile(self):
        fname = QFileDialog.getOpenFileName(self, '打开新文件')
        if fname[0] != '':
            self.NewXlsx = MyXlsx(fname[0])
            self.CentralWidget.setNewTable(self.NewXlsx.SheetDatas)
            self.CellDiffWidget.clear()
            self.RowDiffWidget.clear()
            self.ColDiffWidget.clear()
            self.IsAna = False

    def beginAna(self):
        if self.isanaing == True:
            QMessageBox.warning(self, "提示", "正在分析！请稍等！", QMessageBox.Ok)
            return
        
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
        
        self.isanaing =True
        self.Thread = Thread(self)
        self.Thread.done.connect(self.doneDiff)
        self.Thread.window.Alg.statueSignal.connect(self.bar.showMessage)
        
        self.Thread.start()
        
    def doneDiff(self,diff):
        self.SheetDiff = diff
        self.RowDiffWidget.setData(diff)
        self.ColDiffWidget.setData(diff)
        self.CellDiffWidget.setData(diff)
        self.CentralWidget.setColor(diff)
        self.IsAna = True
        self.isanaing =False
        self.bar.showMessage("分析完毕！")

if __name__=="__main__":

    app = QApplication(sys.argv)
    main = MainWindow()
    main.setWindowIcon(QIcon("icon/opennew.ico"))
    main.show()
    sys.exit(app.exec_())