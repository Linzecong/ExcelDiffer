#-*- codingg:utf8 -*-
from PyQt5.QtWidgets import QWidget, QTabWidget, QHBoxLayout,QVBoxLayout,QListWidget,QSplitter,QLabel,QDialog,QPushButton,QColorDialog
from PyQt5.QtCore import Qt,QSize,QSettings
from PyQt5.QtGui import QColor
import sys

class ChangeColorWidget(QDialog):
    def __init__(self):
        super(ChangeColorWidget,self).__init__()
        self.setWindowTitle("选择颜色")
        self.ColorSettings = QSettings("ExcelDiffer", "Color");
        # self.ColorSettings.setValue("hightlight", "#909399");
        # self.ColorSettings.setValue("background", "#FFFFFF");
        # self.ColorSettings.setValue("exchange", "#EBEEF5");
        # self.ColorSettings.setValue("add", "#409EFF");
        # self.ColorSettings.setValue("delcolor", "#F56C6C");
        # self.ColorSettings.setValue("change", "#E6A23C");

        self.MainLayout = QVBoxLayout()
        self.HightlightLayout = QHBoxLayout()
        self.BackgroundLayout = QHBoxLayout()
        self.ExchangeLayout = QHBoxLayout()
        self.AddLayout = QHBoxLayout()
        self.DelLayout = QHBoxLayout()
        self.ChangeLayout = QHBoxLayout()
        self.ButtonLayout = QHBoxLayout()

        Label = QLabel()
        Label.setText("高亮：")
        self.LabelCH = QLabel()
        self.LabelCH.setStyleSheet("background-color:"+self.ColorSettings.value("hightlight"))
        self.LabelCH.setFixedSize(30,22)
        self.LabelCH.mouseReleaseEvent = self.HC  

        self.HightlightLayout.addStretch()
        self.HightlightLayout.addWidget(Label)
        self.HightlightLayout.addWidget(self.LabelCH)
        self.HightlightLayout.addStretch()

        Label = QLabel()
        Label.setText("背景：")
        self.LabelCB = QLabel()
        self.LabelCB.setStyleSheet("background-color:"+self.ColorSettings.value("background"))
        self.LabelCB.setFixedSize(30,22)
        self.LabelCB.mouseReleaseEvent = self.BC  
        
        self.BackgroundLayout.addStretch()
        self.BackgroundLayout.addWidget(Label)
        self.BackgroundLayout.addWidget(self.LabelCB)
        self.BackgroundLayout.addStretch()

        Label = QLabel()
        Label.setText("交换：")
        self.LabelCE = QLabel()
        self.LabelCE.setStyleSheet("background-color:"+self.ColorSettings.value("exchange"))
        self.LabelCE.setFixedSize(30,22)
        self.LabelCE.mouseReleaseEvent = self.EC  
        
        self.ExchangeLayout.addStretch()
        self.ExchangeLayout.addWidget(Label)
        self.ExchangeLayout.addWidget(self.LabelCE)
        self.ExchangeLayout.addStretch()

        Label = QLabel()
        Label.setText("新增：")
        self.LabelCA = QLabel()
        self.LabelCA.setStyleSheet("background-color:"+self.ColorSettings.value("add"))
        self.LabelCA.setFixedSize(30,22)
        self.LabelCA.mouseReleaseEvent = self.AC  
        
        self.AddLayout.addStretch()
        self.AddLayout.addWidget(Label)
        self.AddLayout.addWidget(self.LabelCA)
        self.AddLayout.addStretch()

        Label = QLabel()
        Label.setText("删除：")
        self.LabelCD = QLabel()
        self.LabelCD.setStyleSheet("background-color:"+self.ColorSettings.value("delcolor"))
        self.LabelCD.setFixedSize(30,22)
        self.LabelCD.mouseReleaseEvent = self.DC  
        
        self.DelLayout.addStretch()
        self.DelLayout.addWidget(Label)
        self.DelLayout.addWidget(self.LabelCD)
        self.DelLayout.addStretch()

        Label = QLabel()
        Label.setText("修改：")
        self.LabelCC = QLabel()
        self.LabelCC.setStyleSheet("background-color:"+self.ColorSettings.value("change"))
        self.LabelCC.setFixedSize(30,22)
        self.LabelCC.mouseReleaseEvent = self.CC  
        
        self.ChangeLayout.addStretch()
        self.ChangeLayout.addWidget(Label)
        self.ChangeLayout.addWidget(self.LabelCC)
        self.ChangeLayout.addStretch()

        self.Button = QPushButton("确定")
        self.ButtonLayout.addStretch()
        self.ButtonLayout.addWidget(self.Button)
        self.Button.clicked.connect(self.accept)

        self.MainLayout.addLayout(self.HightlightLayout)
        self.MainLayout.addLayout(self.BackgroundLayout)
        self.MainLayout.addLayout(self.ExchangeLayout)
        self.MainLayout.addLayout(self.AddLayout)
        self.MainLayout.addLayout(self.DelLayout)
        self.MainLayout.addLayout(self.ChangeLayout)
        self.MainLayout.addSpacing(20)
        self.MainLayout.addLayout(self.ButtonLayout)
        self.setLayout(self.MainLayout)

    def HC(self,x):
        col = QColorDialog.getColor(QColor(self.ColorSettings.value("hightlight")))
        if col.isValid():
            self.LabelCH.setStyleSheet("background-color:"+self.ColorSettings.value("hightlight"))
            self.ColorSettings.setValue("hightlight",col.name())
            self.LabelCH.setStyleSheet("background-color:"+self.ColorSettings.value("hightlight")) # force repaint
    
    def BC(self,x):
        col = QColorDialog.getColor(QColor(self.ColorSettings.value("background")))
        if col.isValid():
            self.LabelCB.setStyleSheet("background-color:"+self.ColorSettings.value("background"))
            self.ColorSettings.setValue("background",col.name())
            self.LabelCB.setStyleSheet("background-color:"+self.ColorSettings.value("background")) # force repaint
    
    def AC(self,x):
        col = QColorDialog.getColor(QColor(self.ColorSettings.value("add")))
        if col.isValid():
            self.LabelCA.setStyleSheet("background-color:"+self.ColorSettings.value("add"))
            self.ColorSettings.setValue("add",col.name())
            self.LabelCA.setStyleSheet("background-color:"+self.ColorSettings.value("add")) # force repaint
        
    def DC(self,x):
        col = QColorDialog.getColor(QColor(self.ColorSettings.value("delcolor")))
        if col.isValid():
            self.LabelCD.setStyleSheet("background-color:"+self.ColorSettings.value("delcolor"))
            self.ColorSettings.setValue("delcolor",col.name())
            self.LabelCD.setStyleSheet("background-color:"+self.ColorSettings.value("delcolor")) # force repaint
        
    def EC(self,x):
        col = QColorDialog.getColor(QColor(self.ColorSettings.value("exchange")))
        if col.isValid():
            self.LabelCE.setStyleSheet("background-color:"+self.ColorSettings.value("exchange"))
            self.ColorSettings.setValue("exchange",col.name())
            self.LabelCE.setStyleSheet("background-color:"+self.ColorSettings.value("exchange")) # force repaint
        
    def CC(self,x):
        col = QColorDialog.getColor(QColor(self.ColorSettings.value("change")))
        if col.isValid():
            self.LabelCC.setStyleSheet("background-color:"+self.ColorSettings.value("change"))
            self.ColorSettings.setValue("change",col.name())
            self.LabelCC.setStyleSheet("background-color:"+self.ColorSettings.value("change")) # force repaint
        
if __name__=="__main__":
    app = QApplication(sys.argv)
    main = ChangeColorWidget()
    main.show()
    sys.exit(app.exec_())