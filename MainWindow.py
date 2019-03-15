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

    def __init__(self,Window,name="###"):
        super().__init__()
        self.window = Window
        self.name = name

    def run(self):
        self.SheetDiff = self.window.Alg[self.name].getSheetdiff()
        self.SheetDiff["diffname"]=self.name
        self.done.emit(self.SheetDiff)


class OpenFileThread(QThread):

    def __init__(self,file):
        super().__init__()
        self.file = file

    def run(self):
        self.Xlrd = MyXlsx(self.file)

 
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
        self.Alg = {}
        self.Alg["###"] = MyAlg()
    
    def restoreToWinSetting(self):
        reply = QMessageBox.warning(self, "确定恢复上次保存的窗口布局吗？", "确定恢复上次保存的窗口布局吗？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
        if reply == QMessageBox.No:
            return
        try:
            self.WinSettings = QSettings("ExcelDiffer", "ExcelDiffer");
            self.restoreGeometry(self.WinSettings.value("geometry"));
            self.restoreState(self.WinSettings.value("windowState"));
        except:
            pass

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
        if self.isanaing == True:
            QMessageBox.warning(self, "提示", "正在分析！请稍等！", QMessageBox.Ok)
            return
        fname = QFileDialog.getOpenFileName(self, '打开旧文件',"","Excel(*.xlsx *.xls)")
        self.oldfilename = fname[0]
        if fname[0] != '':
            self.OpenOldThread = OpenFileThread(fname[0])
            self.OpenOldThread.finished.connect(self.openOldDone)
            self.OpenOldThread.start()
            self.CellDiffWidget.clear()
            self.RowDiffWidget.clear()
            self.ColDiffWidget.clear()
            self.IsAna = False
            self.SheetDiff = {}
            try:
                self.CentralWidget.OldTableWidget.currentChanged.disconnect(self.setDiff)
            except:
                pass

    def openNewFile(self):
        if self.isanaing == True:
            QMessageBox.warning(self, "提示", "正在分析！请稍等！", QMessageBox.Ok)
            return
        fname = QFileDialog.getOpenFileName(self, '打开新文件',"","Excel(*.xlsx *.xls)")
        self.newfilename = fname[0]
        if fname[0] != '':
            self.OpenNewThread = OpenFileThread(fname[0])
            self.OpenNewThread.finished.connect(self.openNewDone)
            self.OpenNewThread.start()

            self.CellDiffWidget.clear()
            self.RowDiffWidget.clear()
            self.ColDiffWidget.clear()
            self.IsAna = False
            self.SheetDiff = {}
            try:
                self.CentralWidget.OldTableWidget.currentChanged.disconnect(self.setDiff)
            except:
                pass

    def openNewDone(self):
        self.NewXlsx = self.OpenNewThread.Xlrd
        self.CentralWidget.setNewTable(self.NewXlsx.SheetDatas)
    
    def openOldDone(self):
        self.OldXlsx = self.OpenOldThread.Xlrd
        self.CentralWidget.setOldTable(self.OldXlsx.SheetDatas)

    def beginAna(self):
        self.Threads = []
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        addcol = self.ColorSettings.value("add")
        delcol = self.ColorSettings.value("delcolor")

        oi = self.CentralWidget.OldTableWidget.currentIndex()
        ni = self.CentralWidget.NewTableWidget.currentIndex()
        if oi == -1 or ni ==-1:
            QMessageBox.warning(self, "提示", "请先打开Excel！", QMessageBox.Ok)
            return

        if self.oldfilename == self.newfilename:
            QMessageBox.warning(self, "提示", "文件相同，无需比较。", QMessageBox.Ok)
            return

        if self.isanaing == True:
            QMessageBox.warning(self, "提示", "正在分析！请稍等！", QMessageBox.Ok)
            return

        if self.CentralWidget.Lock == False:
            reply = QMessageBox.warning(self, "确定比较吗？", "将要比较 "+self.OldXlsx.SheetDatas[oi]["name"]+" 和 "+self.NewXlsx.SheetDatas[ni]["name"], QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
            if reply == QMessageBox.No:
                return
            self.Alg["###"].setOldData(self.OldXlsx.SheetDatas[oi])
            self.Alg["###"].setNewData(self.NewXlsx.SheetDatas[ni])
            
            self.isanaing =True
            thread = Thread(self)
            thread.done.connect(self.doneDiffOne)
            thread.window.Alg["###"].statueSignal.connect(self.bar.showMessage)
            thread.start()
            self.Threads.append(thread)
        else:
            reply = QMessageBox.warning(self, "确定比较吗？", "将要比较整个文件", QMessageBox.Yes | QMessageBox.No, QMessageBox.No);
            if reply == QMessageBox.No:
                return

            oldcount = self.CentralWidget.OldTableWidget.count()
            newcount = self.CentralWidget.NewTableWidget.count()

            self.SheetDiff = {}
            self.diffcount=0
            self.donecount=0

            self.delsheetcount = 0
            self.addsheetcount = 0

            # 计算删除的标签页
            for i in range(oldcount):
                flag = False
                for j in range(newcount):
                    if self.CentralWidget.OldTableWidget.tabText(i) == self.CentralWidget.NewTableWidget.tabText(j):
                        self.diffcount = self.diffcount+1
                        flag = True
                if flag == False:
                    self.CentralWidget.OldTableWidget.setTabBarColor(i,delcol)
                    self.delsheetcount = self.delsheetcount + 1
            
            # 计算新增的标签页
            for i in range(newcount):
                flag = False
                for j in range(oldcount):
                    if self.CentralWidget.OldTableWidget.tabText(j) == self.CentralWidget.NewTableWidget.tabText(i):
                        flag = True
                if flag == False:
                    self.CentralWidget.NewTableWidget.setTabBarColor(i,addcol)    
                    self.addsheetcount = self.addsheetcount + 1
            
            have = False
            for i in range(oldcount):
                for j in range(newcount):
                    if self.CentralWidget.OldTableWidget.tabText(i) == self.CentralWidget.NewTableWidget.tabText(j):
                        have = True
                        self.Alg[self.CentralWidget.OldTableWidget.tabText(i)] = MyAlg()
                        self.Alg[self.CentralWidget.OldTableWidget.tabText(i)].setOldData(self.OldXlsx.SheetDatas[i])
                        self.Alg[self.CentralWidget.OldTableWidget.tabText(i)].setNewData(self.NewXlsx.SheetDatas[j])
                        self.isanaing =True
                        thread = Thread(self,self.CentralWidget.OldTableWidget.tabText(i))
                        thread.done.connect(self.doneDiffAll)
                        thread.window.Alg[self.CentralWidget.OldTableWidget.tabText(i)].statueSignal.connect(self.bar.showMessage)
                        thread.start()
                        self.Threads.append(thread)
            
            if have == False:
                QMessageBox.information(self, "提示", "没有同名Sheet需要比较！", QMessageBox.Ok)
            
        
    def doneDiffAll(self,diff):
        self.SheetDiff[diff["diffname"]]=diff
        self.donecount = self.donecount + 1    
        if self.donecount == self.diffcount:

            self.IsAna = True
            self.isanaing =False
            self.CentralWidget.OldTableWidget.currentChanged.connect(self.setDiff)

            flag = True
            for diffff in self.SheetDiff:
                for item in self.SheetDiff[diffff]:
                    if len(self.SheetDiff[diffff][item]) != 0 and item != "diffname":
                        flag = False
            if flag == True:
                QMessageBox.information(self, "提示", "分析完毕！无改动！", QMessageBox.Ok)
                return

            oi = self.CentralWidget.OldTableWidget.currentIndex()
            text = self.CentralWidget.OldTableWidget.tabText(oi)
            if text in self.SheetDiff:
                self.RowDiffWidget.setData(self.SheetDiff[text])
                self.ColDiffWidget.setData(self.SheetDiff[text])
                self.CellDiffWidget.setData(self.SheetDiff[text])
                self.CentralWidget.setColor(self.SheetDiff[text])
            QMessageBox.information(self, "提示", "分析完毕，共新增了"+str(self.addsheetcount)+"个标签页，删除了"+str(self.delsheetcount)+"个标签页。\n切换标签页查看差异信息！", QMessageBox.Ok)
            
    
    def setDiff(self,id):
        self.RowDiffWidget.clear()
        self.ColDiffWidget.clear()
        self.CellDiffWidget.clear()
        text = self.CentralWidget.OldTableWidget.tabText(id)
        if text in self.SheetDiff:
            self.RowDiffWidget.setData(self.SheetDiff[text])
            self.ColDiffWidget.setData(self.SheetDiff[text])
            self.CellDiffWidget.setData(self.SheetDiff[text])
            self.CentralWidget.setColor(self.SheetDiff[text])

    def doneDiffOne(self,diff):
        self.SheetDiff = diff
        self.RowDiffWidget.setData(diff)
        self.ColDiffWidget.setData(diff)
        self.CellDiffWidget.setData(diff)
        self.CentralWidget.setColor(diff)
        self.IsAna = True
        self.isanaing =False
        self.bar.showMessage("分析完毕！")
        flag = True
        for item in diff:
            
            if len(diff[item]) != 0 and item != "diffname":
                flag = False
        
        if flag == False:
            QMessageBox.information(self, "提示", "分析完毕！", QMessageBox.Ok)
        else:
            QMessageBox.information(self, "提示", "分析完毕！无改动！", QMessageBox.Ok)

if __name__=="__main__":

    app = QApplication(sys.argv)
    main = MainWindow()
    main.setWindowIcon(QIcon("icon/opennew.ico"))
    main.show()
    sys.exit(app.exec_())