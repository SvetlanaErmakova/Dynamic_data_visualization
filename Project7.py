from PySide2 import QtWidgets
from PySide2.QtWidgets import QMainWindow, QApplication, QWidget
from PySide2.QtWidgets import QAction
from PySide2.QtWidgets import QDockWidget, QListWidget, QTabWidget, QLabel, QLineEdit
from PySide2.QtWidgets import QListView, QGroupBox, QComboBox, QStackedWidget, QSpacerItem
from PySide2.QtWidgets import QPushButton
from PySide2.QtWidgets import QFileDialog
from PySide2.QtWidgets import QGridLayout, QVBoxLayout
from PySide2.QtWidgets import QHeaderView,QHBoxLayout,  QSizePolicy, QTableView, QSplitter,QScrollArea

from PySide2 import QtCore
from PySide2.QtCore import QStringListModel
from PySide2.QtCore import (QAbstractTableModel, QDateTime, QModelIndex, QTimeZone, QPointF, Slot,QPoint)
from PySide2.QtCore import Qt

from PySide2 import QtGui
from PySide2.QtGui import QColor, QPainter,QPen, QLinearGradient, QGradient
from PySide2.QtGui import QStandardItemModel, QStandardItem
from PySide2.QtGui import QImageReader, QIcon

import PySide2
from PySide2.QtCharts import QtCharts

import sys
import argparse
import pandas as pd
import copy
import statistics
#import sip
import time

import numpy as np

import pyautogui
#import imutils
#import cv2

class ChartWidget(QWidget):
    def __init__(self, model, tabWidget, parent=None):
        super(ChartWidget, self).__init__(parent)
        self.tabs_model=parent.model
        self.model = model
        self.tabWidget = tabWidget
        layout = QVBoxLayout(self)
        chartGroupBox = QGroupBox("General Settings")
        
        grid = QGridLayout(self)
        grid.addWidget(QLabel("Name"), 0, 0)
        self.nameLineEdit = QLineEdit()
        self.nameLineEdit.setText(model.text())
        self.nameLineEdit.textChanged.connect(self.updateText)

        grid.addWidget(self.nameLineEdit, 0, 1)
        self.typeComboBox = QComboBox()
        self.typeComboBox.activated.connect(self.changeType)
       
        grid.addWidget(self.typeComboBox, 1, 1)

        chartGroupBox.setLayout(grid)                
        
        layout.addWidget(chartGroupBox)

        self.chartWidget = QWidget()
        layout.addWidget(self.chartWidget)

        self.stackedWidget = QStackedWidget(self)
        layout.addWidget(self.stackedWidget)
        
        self.setLayout(layout)

        chartTypesLst = [('Bar Chart', "icons/chart-types/bar-chart.svg", [Graphic_B,Settings_BC]),
             ('Pie Chart', "icons/chart-types/pie-chart.svg", [Graphic_P,Settings_P]),
             ('Spline Chart', "icons/chart-types/spline-chart.svg", [Graphic_S,Settings_S])]
        s = QStandardItemModel()
        for ct in chartTypesLst:
            item = QStandardItem(ct[0])
            item.setIcon(QIcon(ct[1]))
            item.setData(ct[2])
            s.appendRow(item) 
        self.typeComboBox.setModel(s)
    def changeType(self):  
        self.typeComboBox.setEnabled(False)
        if self.stackedWidget.count() != 0: 
            self.stackedWidget.removeWidget(self.stackedWidget.widget(0))

        row = self.typeComboBox.currentIndex()
        idx = self.typeComboBox.model().index(row, 0)
        model = self.typeComboBox.model().itemFromIndex(idx)
        t = model.data()
        if t == [Graphic_B,Settings_BC]:
            widget=Chart(t,0)
            self.stackedWidget.addWidget(widget)
        elif t == [Graphic_S,Settings_S]:
            widget = Chart(t,1)
            self.stackedWidget.addWidget(widget)
        elif t == [Graphic_P,Settings_P]:
            widget = Chart(t,2)
            self.stackedWidget.addWidget(widget)

    def updateText(self):
        name = self.nameLineEdit.text()
        
        self.model.setText(name)
        ind = self.tabWidget.indexOf(self)
        self.tabWidget.setTabText(ind, name)
        
        for i in range(self.tabs_model.rowCount()):
            el=self.tabs_model.itemFromIndex(self.tabs_model.index(i, 0))
            if(el.model_name =='D'):
                el.tab.Settings_D.update_all_charts(self.tabs_model)

class Chart(QWidget):
    def __init__(self,arr,Type): 
        QWidget.__init__(self)
        self.typeGr=arr[0]
        self.typeSt=arr[1]

        self.splitter = QSplitter(QtCore.Qt.Vertical)  
        self.On_DB=False
        self.flag=False
        self.Type=Type

        self.chartView = None
        self.chartView_D =None

        self.button=QtWidgets.QPushButton("Построить график")
        self.button1=QtWidgets.QPushButton("Сохранить график")

        grid = QGridLayout(self)
        self.settingsGrid = QGridLayout(self)
        self.settingsGrid.addWidget(QLabel("Файл:"), 0, 0)
        self.datasetLabel = QLabel("нет данных")
        self.settingsGrid.addWidget(self.datasetLabel, 1, 0, 1, 2)
        self.dataSetButton = QPushButton("Выбрать файл", self)
        self.dataSetButton.setIcon(QIcon('icons/file.svg'))
        self.dataSetButton.clicked.connect(self.chooseDataset)
        self.settingsGrid.addWidget(self.dataSetButton, 0, 1)

        grid.addLayout(self.settingsGrid, 0, 0)
        grid.addWidget(self.splitter, 1, 0)
        self.setLayout(grid)

        self.button.clicked.connect(self.on_clicked_button)
        self.button1.clicked.connect(self.on_clicked_button1)
    def chooseDataset(self):
        self.dataSetButton.setEnabled(False)
        file = QtWidgets.QFileDialog.getOpenFileName(parent=window, 
            caption="Заголовок окна", directory="c:\\", 
            filter="All (*);;Exes (*.exe *.dll)", 
            initialFilter="Exes (*.exe *.dll)") 
        self.filename = file[0] 
        self.datasetLabel.setText(self.filename)
        self.Settings_G = self.typeSt(self) 
        self.splitter.addWidget(self.Settings_G)
    def on_clicked_button(self):
        if((self.Settings_G.criteria_x.text()!=' ') ==True &(self.Settings_G.criteria_y.text()!=' ')==True):
            if(self.Type!=2):
                self.Settings_G.Max_Min()
            self.button.setEnabled(False)
            graphic = self.typeGr(self.Settings_G)
            self.chartView=graphic.chartView
            self.chartView.setMinimumSize(100, 300)
            
            self.Settings_G.main_layout.addLayout(self.Settings_G.layout_B)

            self.chartView_D=QtCharts.QChartView(QtCharts.QChart())
            self.splitter.addWidget(self.chartView)
            self.splitter.setCollapsible(self.splitter.count()-1,False)
            self.flag=True

    def on_clicked_button1(self):
        K1=self.splitter.widget(1).pos()
        size=self.splitter.widget(1).size()
        K2=QPoint(size.width(), size.height())

        size_b=self.Settings_G.Button_Back.size()
        KK1=self.mapToGlobal(K1)
        image = pyautogui.screenshot(region=(KK1.x(),KK1.y()+2*size_b.height(), K2.x(),K2.y()))

        image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        file = QtWidgets.QFileDialog.getSaveFileName(parent=window) 
        path=file[0]
        if(path.endswith('.png')==False):
            path=path+'.png'
        cv2.imwrite(path, image)

        

class Settings(QWidget):
    def __init__(self, Chart):
        super(Settings, self).__init__()
        self.main_layout=QtWidgets.QVBoxLayout()
        Layout=QtWidgets.QHBoxLayout()
        self.layout1=QtWidgets.QVBoxLayout()
        self.layout2=QtWidgets.QVBoxLayout()
        self.layout_B=QtWidgets.QHBoxLayout()
        self.Chart=Chart
        self.Type=Chart.Type
        self.filename=Chart.filename
        self.normalname=self.normalN(self.filename)
        self.xls=None
        self.List_Number=0
        self.func_Data()

        self.Button_Forw=QtWidgets.QPushButton()
        self.Button_Forw.setIcon(QIcon('icons/forward.svg'))
        self.Button_Back=QtWidgets.QPushButton()
        self.Button_Back.setIcon(QIcon('icons/backwards.svg'))
        self.Button_Forw.clicked.connect(self.on_clicked_Button_Forw)
        self.Button_Back.clicked.connect(self.on_clicked_Button_Back)

        if(Chart.Type!=2):
            self.Color=[]

        self.numbers=None 
        self.Columns=None 

        Layout.addLayout(self.layout1)
        Layout.addLayout(self.layout2)
        self.layout_B.addWidget(self.Button_Back)
        self.layout_B.addWidget(self.Button_Forw)

        self.main_layout.addLayout(Layout)
        self.setLayout(self.main_layout)
    def add_predstavlenie_y(self):
        
        self.Colums_model_y = QtCore.QStringListModel(self.numbers)
        self.criteria_y=QLabel(" ")
        self.Arr_Col_y=[]
        self.Colum_predstavl_y = QtWidgets.QListView()
        self.Colum_predstavl_y.setModel(self.Colums_model_y) 

        self.layout1.addWidget(self.Colum_predstavl_y)
        if(self.Type==2):
            label=QLabel("Временной промежуток: ")
        elif(self.Type<2):
            label=QLabel("По Оy: ")
        self.layout1.addWidget(label)
        self.layout1.addWidget(self.criteria_y)

        self.Colum_predstavl_y.clicked.connect(self.Choice_y)
    def add_predstavlenie_x(self):
        self.Colums_model_x = QtCore.QStringListModel(self.Columns) 

        self.criteria_x=QLabel(" ")
        self.Colum_predstavl_x = QtWidgets.QListView()

        self.Colum_predstavl_x.setModel(self.Colums_model_x) 

        self.layout2.addWidget(self.Colum_predstavl_x)
        if(self.Type==2):
            label=QLabel("Название: ")
        elif(self.Type<2):
            label=QLabel("По Оx: ")
        self.layout2.addWidget(label)
        self.layout2.addWidget(self.criteria_x)

        self.Colum_predstavl_x.clicked.connect(self.Choice_x)
    def normalN(self, filename):
        index=filename.rfind('/')
        return filename[index+1:]
    def func_Data(self):
        if(self.filename[-1:]=='x'):
            self.xls = pd.ExcelFile(self.filename) 
            self.Data = pd.read_excel(self.filename, sheet_name=self.xls.sheet_names[self.List_Number])
    def on_clicked_Button_Forw(self):
        self.List_Number=self.List_Number+1
        if(self.List_Number>=len(self.xls.sheet_names)-1):
            self.Button_Forw.setEnabled(False)
        self.Data = pd.read_excel(self.filename, sheet_name=self.xls.sheet_names[self.List_Number])
        self.RePaint_G()
        if(self.Button_Back.isEnabled()==False):
            self.Button_Back.setEnabled(True)
    def on_clicked_Button_Back(self):
        self.List_Number=self.List_Number-1
        if(self.List_Number<=0):
            self.Button_Back.setEnabled(False)
        self.Data = pd.read_excel(self.filename, sheet_name=self.xls.sheet_names[self.List_Number])
        self.RePaint_G()
        if(self.Button_Forw.isEnabled()==False):
            self.Button_Forw.setEnabled(True)

class Graphic_B(QWidget):
    def __init__(self, Settings_G): 
        QWidget.__init__(self)
        self.Settings_G=Settings_G
        self.Data=Settings_G.Data
        self.add_Graphic()
        Settings_G.Chart.settingsGrid.addWidget(Settings_G.Chart.button1, 0, 3)

    def add_Graphic(self):
        Yarr_c = self.Settings_G.Arr_Col_y 
        Yarr_d=[]
        for i in range(len(Yarr_c)):
            Yarr_d.append(self.Data[Yarr_c[i]].tolist())
        X_c = self.Settings_G.Col_x 
        X_d=self.Data[X_c].tolist() 
        self.chart = self.createBarCharts(Yarr_d, Yarr_c, X_c, X_d)
        self.chart.setContentsMargins(-10, -10,-10,-10)

        self.chartView = QtCharts.QChartView(self.chart)
        self.chartView.setRenderHint(QPainter.Antialiasing)
    def createBarCharts(self, Yarr_d, Yarr_c, X_c, X_d): 
        barSet=[]
        stroka=""
        for i in range(len(Yarr_c)):
            barSet.append(QtCharts.QBarSet(Yarr_c[i])) 
            stroka=stroka+Yarr_c[i]+" "
            barSet[len(barSet)-1].append(Yarr_d[i])
            barSet[len(barSet)-1].setColor(self.Settings_G.Color[i]) 
        barSeries1 = QtCharts.QBarSeries()
        for u in range(len(barSet)):
            barSeries1.append(barSet[u])
        chart = QtCharts.QChart()
        chart.addSeries(barSeries1)
        chart.setTitle(self.Settings_G.normalname + " "+self.Settings_G.xls.sheet_names[self.Settings_G.List_Number])

        categories = X_d
        barCategoryAxis = QtCharts.QBarCategoryAxis()
        barCategoryAxis.append(categories)
        chart.createDefaultAxes()
        chart.setAxisX(barCategoryAxis, barSeries1)
                                                          
        chart.axes(Qt.Vertical)[0].setRange(self.Settings_G.MIN,self.Settings_G.MAX)
        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignTop)

        return chart
class Settings_BC(Settings):
    def __init__(self, Chart):
        super(Settings_BC, self).__init__(Chart)
        if(self.xls!= None):
            self.numbers=selection_Numbers_without_NULL(self.Data)
            self.add_predstavlenie_y()
            self.Columns=selection_Object_without_NULL(self.Data)
            self.add_predstavlenie_x()
            Chart.settingsGrid.addWidget(Chart.button, 0, 2)
    def RePaint_G(self):
        S=self.Chart.splitter.sizes()
        g=Graphic_B(self)
        self.Chart.chartView.setChart(g.chart)
        self.Chart.splitter.setSizes(S)
        if(self.Chart.On_DB==True):
            self.Repaint_D()
    def Repaint_D(self):
        g=Graphic_B(self)
        self.Chart.chartView_D.setChart(g.chart)
    def Choice_y(self):
        ff=self.Colum_predstavl_y.currentIndex()
        for i in range(len(self.numbers)):
            if(i==ff.row()):
                self.Color.append(PySide2.QtWidgets.QColorDialog.getColor())
                self.Arr_Col_y.append(self.numbers[i])
                str=self.criteria_y.text()    
                self.criteria_y.setText(str+self.Arr_Col_y[len(self.Arr_Col_y)-1]+" ")
    def Choice_x(self):
        ff=self.Colum_predstavl_x.currentIndex()
        if(self.criteria_x.text()==" "):
            for i in range(len(self.Columns)):
                    if(i==ff.row()):
                        self.Col_x=self.Columns[i]
                        str=self.criteria_x.text()    
                        self.criteria_x.setText(str+self.Col_x+" ")
                        break
    def Max_Min(self):
        Max=[]
        Min=[]
        if(self.Chart.filename[-1:]=='x'):
            xls=pd.ExcelFile(self.Chart.filename)    
            for i in range(len(xls.sheet_names)):
                Data = pd.read_excel(self.Chart.filename, sheet_name=xls.sheet_names[i])
                for i in self.Arr_Col_y:
                    Max.append(max(Data[i].tolist()))
                    Min.append(min(Data[i].tolist()))
        self.MAX=max(Max)
        self.MIN=min(Min)

class Graphic_S(QWidget):
    def __init__(self, Settings_G): 
        QWidget.__init__(self)
        self.Settings_G=Settings_G
        self.Data=Settings_G.Data
        self.add_Graphic()
        Settings_G.Chart.settingsGrid.addWidget(Settings_G.Chart.button1, 0, 3)
    def add_Graphic(self):
        Yarr_c = self.Settings_G.Arr_Col_y 
        Yarr_d=[]
        for i in range(len(Yarr_c)):
            Yarr_d.append(self.Data[Yarr_c[i]].tolist())

        X_c = self.Settings_G.Col_x 
        X_d=self.Data[X_c].tolist() 

        data=self.getSplineChartData(Yarr_c, Yarr_d,X_c,X_d)
        self.chart = QtCharts.QChart()
        self.chart.setTitle(self.Settings_G.normalname + " "+self.Settings_G.xls.sheet_names[self.Settings_G.List_Number])
        axisX = QtCharts.QDateTimeAxis();
        axisX.setFormat("MM-yyyy");
        axisX.setTitleText(X_c);
        self.chart.addAxis(axisX, Qt.AlignBottom);
        axisY = QtCharts.QValueAxis();
        axisY.setTitleText("oY");
        self.chart.addAxis(axisY, Qt.AlignLeft);
        Mi=self.Settings_G.MIN
        Ma=self.Settings_G.MAX
                                                              
        self.chart.axes(Qt.Vertical)[0].setRange(Mi-(Ma-Mi)*0.5, Ma+(Ma-Mi)*0.5)
        valuesToDraw = [self.Data[i].tolist() for i in Yarr_c]
        namesToDraw = Yarr_c        
        for i in range(len(namesToDraw)):
            values = valuesToDraw[i]
            name = namesToDraw[i]
            series = QtCharts.QSplineSeries()
            series.setName(name)
            series.setColor(self.Settings_G.Color[i])

            n = len(values)
            for j in range(n):
                time = data[0][j]                
                series.append(float(time.toMSecsSinceEpoch()), values[j])
            self.chart.addSeries(series)
            series.attachAxis(axisX);        
            series.attachAxis(axisY);
        self.chart.axes(Qt.Horizontal)[0].setRange(data[0][0],data[0][len(data[0])-1])
        self.chart.axes(Qt.Horizontal)[0].setTickCount(10)

        self.chart.setContentsMargins(-10, -10,-10,-10)

        self.chartView = QtCharts.QChartView(self.chart)   
        self.chartView.setRenderHint(QPainter.Antialiasing)  
        
    def getSplineChartData(self,Yarr_c, Yarr_d,X_c,X_d ): 
        xls = pd.ExcelFile(self.Settings_G.filename) 
        data = pd.read_excel(self.Settings_G.filename, sheet_name=xls.sheet_names[len(xls.sheet_names)-1])
        timeStamps = data[X_c]                                         
        timeStamps = [QDateTime.fromString(str(x),"yyyy")
                      for x in timeStamps]                                  
        values = []                                                         
        names = []    

        for x in self.Data.loc[:, Yarr_c[0]:]:                              
            values.append(list(self.Data[x]))                               
            names.append(x)                                                 
        return timeStamps, values, names                                    
class Settings_S(Settings):
    def __init__(self, Chart):
        super(Settings_S, self).__init__(Chart)
        if(self.xls!=None):
            self.numbers=selection_Numbers_without_NULL(self.Data)
            self.add_predstavlenie_y()
            self.Columns=selection_Object_without_NULL(self.Data)
            self.add_predstavlenie_x()
            Chart.settingsGrid.addWidget(Chart.button, 0, 2)
    def RePaint_G(self):
        S=self.Chart.splitter.sizes()
        g=Graphic_S(self)
        self.Chart.chartView.setChart(g.chart)
        self.Chart.splitter.setSizes(S)
        if(self.Chart.On_DB==True):
            self.Repaint_D()
    def Repaint_D(self):
        g=Graphic_S(self)
        self.Chart.chartView_D.setChart(g.chart)
    def Choice_y(self):
        ff=self.Colum_predstavl_y.currentIndex()
        for i in range(len(self.numbers)):
            if(i==ff.row()):
                self.Color.append(PySide2.QtWidgets.QColorDialog.getColor())
                self.Arr_Col_y.append(self.numbers[i])
                str=self.criteria_y.text()    
                self.criteria_y.setText(str+self.Arr_Col_y[len(self.Arr_Col_y)-1]+" ")
    def Choice_x(self):
        ff=self.Colum_predstavl_x.currentIndex()
        if(self.criteria_x.text()==" "):
            for i in range(len(self.Columns)):
                    if(i==ff.row()):
                        self.Col_x=self.Columns[i]
                        str=self.criteria_x.text()    
                        self.criteria_x.setText(str+self.Col_x+" ")
                        break
    def Max_Min(self):
        Max=[]
        Min=[]
        if(self.Chart.filename[-1:]=='x'):
            xls=pd.ExcelFile(self.Chart.filename)    
            for i in range(len(xls.sheet_names)):
                Data = pd.read_excel(self.Chart.filename, sheet_name=xls.sheet_names[i])
                for i in self.Arr_Col_y:
                    Max.append(max(Data[i].tolist()))
                    Min.append(min(Data[i].tolist()))
        self.MAX=max(Max)
        self.MIN=min(Min)

class Graphic_P(QWidget):
    def __init__(self, Settings_G): 
        QWidget.__init__(self)
        self.Settings_G=Settings_G
        self.Data=Settings_G.Data
        self.add_Graphic()
        Settings_G.Chart.settingsGrid.addWidget(Settings_G.Chart.button1, 0, 3)
    def add_Graphic(self):
        Y_c = self.Settings_G.Col_y 
        Y_d=self.Data[Y_c].tolist() 
        X_c = self.Settings_G.Col_x
        X_d=self.Data[X_c].tolist() 

        data=self.getPieChartData(X_c, Y_c)

        series = QtCharts.QPieSeries()
        sliceList = list(map(lambda nv: QtCharts.QPieSlice(nv[0], nv[1]),
                             zip(data[0], data[1])))
        series.append(sliceList)
        for s in series.slices():
            s.setLabelVisible()
        self.chart = QtCharts.QChart();
        self.chart.setTitle(self.Settings_G.normalname + " "+self.Settings_G.xls.sheet_names[self.Settings_G.List_Number])

        self.chart.addSeries(series);
        self.chart.setContentsMargins(-10, -10,-10,-10)

        self.chartView = QtCharts.QChartView(self.chart)
        self.chartView.setRenderHint(QPainter.Antialiasing);
    def getPieChartData(self,X_c, Y_c):
        self.Data = self.Data[[X_c,Y_c]]
        s = self.Data[Y_c].sum()
        threshold = s / 100
        threshold *= 2
        self.Data = self.Data.loc[self.Data[Y_c] > threshold]

        names = list(self.Data[X_c])
        no = list(self.Data[Y_c])
        return names, no
class Settings_P(Settings):
    def __init__(self, Chart):
        super(Settings_P, self).__init__(Chart)
        if(self.xls!=None):
            self.numbers=selection_Numbers(self.Data)
            self.add_predstavlenie_y()
            self.Columns=selection_Object_without_NULL(self.Data)
            self.add_predstavlenie_x()
            Chart.settingsGrid.addWidget(Chart.button, 0, 2)
    def RePaint_G(self):
        S=self.Chart.splitter.sizes()
        g=Graphic_P(self)
        self.Chart.chartView.setChart(g.chart)
        self.Chart.splitter.setSizes(S)
        if(self.Chart.On_DB==True):
            self.Repaint_D()
    def Repaint_D(self):
        g=Graphic_P(self)
        self.Chart.chartView_D.setChart(g.chart)
    def Choice_x(self):
        ff=self.Colum_predstavl_x.currentIndex()
        if(self.criteria_x.text()==" "):
            for i in range(len(self.Columns)):
                if(i==ff.row()):
                    self.Col_x=self.Columns[i]
                    self.criteria_x.setText(self.Col_x)
                    break
    def Choice_y(self):
        ff=self.Colum_predstavl_y.currentIndex()
        if(self.criteria_y.text()==" "):
            for i in range(len(self.numbers)):
                if(i==ff.row()):
                    self.Col_y=self.numbers[i]
                    self.criteria_y.setText(self.Col_y)
                    break

def selection_Numbers_without_NULL(Data):
        arr=[]
        numerics1=['int16', 'int32', 'int64', 'float16', 'float32', 'float64']
        ddd=Data.select_dtypes(include=numerics1) 
        newdf=ddd.isnull().any() 
        for i in range(len(newdf.values)):
            if(newdf.values[i]==False):
                arr.append(ddd.columns.values[i])
        return arr
def selection_Numbers(Data):
        arr=[]
        numerical1=['int16', 'int32', 'int64', 'float16', 'float32', 'float64']
        numbers_null=Data.select_dtypes(include=numerical1)
        for i in range(len(numbers_null.columns.values)):
            column=numbers_null.columns.values[i]
            if(numbers_null[column].count()!=0):
                arr.append(column)
        return arr 
def selection_Object_without_NULL(Data):
        arr=[]
        numerics1=['object','int16', 'int32', 'int64', 'float16', 'float32', 'float64']
        ddd=Data.select_dtypes(include=numerics1)
        newdf=Data.isnull().any() 
        for i in range(len(newdf.values)):
            if(newdf.values[i]==False):
                arr.append(ddd.columns.values[i])
        return arr

class TableWidget(QWidget):
    def __init__(self, model, parent=None):
        super(TableWidget, self).__init__(parent)
        self.model = model

        layout = QVBoxLayout()
        table=Table()
        layout.addWidget(table)

        self.setLayout(layout)   
class Table(QtWidgets.QWidget):
    def __init__(self):
        QtWidgets.QWidget.__init__(self)

        self.main_layout = QVBoxLayout()
        self.splitter = QSplitter()  
        self.splitter.setChildrenCollapsible(False)
        self.widg=QWidget()
        l= QVBoxLayout()
        l.addWidget(self.splitter)
        self.widg.setLayout(l)

        self.button=QtWidgets.QPushButton("Выбрать файл")    
        self.button.setIcon(QIcon('icons/file.svg'))
        self.button.setCheckable(True)

        self.scrollArea=QScrollArea()

        self.scrollArea.setWidget(self.widg)
        self.scrollArea.setWidgetResizable(True)

        self.main_layout.addWidget(self.button)
        self.main_layout.addWidget(self.scrollArea)
        self.setLayout(self.main_layout)

        self.button.clicked.connect(self.on_clicked_button)
    def on_clicked_button(self):
        file = QtWidgets.QFileDialog.getOpenFileName(parent=window, 
            caption="Заголовок окна", directory="c:\\", 
            filter="All (*);;Exes (*.exe *.dll)", 
            initialFilter="Exes (*.exe *.dll)") 
        filename = file[0] 
        if(filename[-1:]=='x'):
            self.button.setEnabled(False)
            xls = pd.ExcelFile(filename) 
            for i in range(len(xls.sheet_names)):
                Data = pd.read_excel(filename, sheet_name=xls.sheet_names[i])
                f=6
                self.build_table(Data)
                f=7
            self.widg.setMinimumSize(len(xls.sheet_names)*len(Data.columns.values)*100, 100)
        elif(filename==''):
            self.button.setChecked(False)
    def build_table(self,Data):
        cc=Data.columns.values
        gg=Data[cc[1]].values
        tv2 = QtWidgets.QTableView() 
        sti2 = QtGui.QStandardItemModel() 
        mas_column=[]
        for r in range(len(cc)):
            for y in range(len(Data)):
                mas_column.append(QtGui.QStandardItem(str(Data[cc[r]].values[y]))) 
            sti2.appendColumn(mas_column)
            mas_column.clear() 
        sti2.setHorizontalHeaderLabels(Data.columns) 
        tv2.setModel(sti2) 
        self.splitter.addWidget(tv2)

class DashboardWidget (QWidget):
    def __init__(self, model, parent=None):
        super(DashboardWidget, self).__init__(parent)
        main_layout = QtWidgets.QVBoxLayout()

        self.model = model
        self.tabs_model=parent.model

        self.splitter = QSplitter(QtCore.Qt.Vertical)  

        self.Settings_D = Settings_Dash(self)
        self.D_Graphics= D_Graphic(self.Settings_D) 
        self.scrollArea=QScrollArea()

        self.scrollArea.setMinimumSize(100, 400)
        self.scrollArea.setWidget(self.D_Graphics)
        self.scrollArea.setWidgetResizable(True)

        self.splitter.addWidget(self.Settings_D)
        self.splitter.addWidget(self.scrollArea)
        self.splitter.setCollapsible(1, False)

        main_layout.addWidget(self.splitter)

        self.setLayout(main_layout) 
class Settings_Dash(QWidget):
    def __init__(self, dashboard): 
        QWidget.__init__(self)
        self.main_layout = QtWidgets.QVBoxLayout()
        self.layout=QtWidgets.QHBoxLayout()

        self.D =dashboard
        self.tabs_model=dashboard.tabs_model
        self.all_charts =QListView()
        self.update_all_charts(self.tabs_model)

        self.Setting_SH=[]

        self.Time=500
        self.combobox=QComboBox()


        model = ['500мс', '1000мс','1500мс','2000мс','2500мс','3000мс','3500мс','4000мс','4500мс','5000мс'] 
        s1 = QtCore.QStringListModel(model) 
        self.combobox.setModel(s1) 
        self.combobox.activated.connect(self.time_for_CB)

        self.button1=QPushButton()
        self.button1.setIcon(QIcon('icons/dynamic_forw.svg'))
        self.button1.setCheckable(True)
        
        self.button2=QPushButton()
        self.button2.setIcon(QIcon('icons/dynamic_back.svg'))
        self.button2.setCheckable(True)

        self.layout.addWidget(self.button2)
        self.layout.addWidget(self.button1)

        self.main_layout.addWidget(self.all_charts)

        self.setLayout(self.main_layout)

        self.button1.clicked.connect(self.on_clicked_button1)
        self.button2.clicked.connect(self.on_clicked_button2)
        self.all_charts.clicked.connect(self.cbo_on_clicked)

    def update_all_charts(self, tabs):
        s = QStandardItemModel()
        for i in range(tabs.rowCount()):
            el=tabs.itemFromIndex(tabs.index(i, 0))
            if(el.model_name =='C'):
                item = QStandardItem(el.text())
                item.setData(el.tab)
                s.appendRow(item) 
        self.all_charts.setModel(s)
    def cbo_on_clicked(self):
        ff=self.all_charts.currentIndex()
        tt=self.all_charts.model().itemFromIndex(ff).data()
        if(tt.stackedWidget.widget(0).flag==True):
            if(tt.stackedWidget.widget(0).On_DB!=True):
                self.Setting_SH.append(tt.stackedWidget.widget(0))
                self.D.D_Graphics.add_Graphic_on_DB(tt.stackedWidget.widget(0))
                tt.stackedWidget.widget(0).On_DB=True
                if(self.main_layout.indexOf(self.button1)==-1):
                    self.main_layout.addLayout(self.layout)
                    self.main_layout.addWidget(self.combobox)
    def on_clicked_button1(self):
        value=self.button1.isChecked()
        if(value==True):
            self.button1.setChecked(True)
            self.button1.setIcon(QIcon('icons/pause.svg'))

            self.timer_id = self.startTimer(self.Time)
            self.combobox.setEnabled(False)
            if(self.button2.isEnabled()==True):
                self.button2.setEnabled(False)
        elif(value==False):

            self.button1.setChecked(False)
            self.button1.setIcon(QIcon('icons/dynamic_forw.svg'))

            self.killTimer(self.timer_id)
            self.combobox.setEnabled(True)

            if(self.button2.isEnabled()==False):
                self.button2.setEnabled(True)

    def on_clicked_button2(self):
        value=self.button2.isChecked()
        if(value==True):
            self.button2.setChecked(True)
            self.button2.setIcon(QIcon('icons/pause.svg'))
            self.timer_id = self.startTimer(self.Time)
            self.combobox.setEnabled(False)
            if(self.button1.isEnabled()==True):
                self.button1.setEnabled(False)
        elif(value==False):
            self.button2.setChecked(False)
            self.button2.setIcon(QIcon('icons/dynamic_back.svg'))

            self.killTimer(self.timer_id)
            self.combobox.setEnabled(True)

            if(self.button1.isEnabled()==False):
                self.button1.setEnabled(True)
    def time_for_CB(self):
        text=self.combobox.currentText()
        index=text.find('м')
        self.Time=float(text[:index])
    def timerEvent(self, event):        
        S1=self.D.D_Graphics.splitter1.sizes()
        S2=self.D.D_Graphics.splitter2.sizes()
        S3=self.D.D_Graphics.splitter3.sizes()
        arr=self.Setting_SH
        for u in arr:
            if(self.button1.isChecked()==True):
                if(u.Settings_G.List_Number>=len(u.Settings_G.xls.sheet_names)-1):
                    u.Settings_G.List_Number=-1
                    u.Settings_G.Button_Forw.setEnabled(True)
                u.Settings_G.Button_Forw.click()
            elif(self.button2.isChecked()==True):
                if(u.Settings_G.List_Number<=0): 
                    u.Settings_G.List_Number=len(u.Settings_G.xls.sheet_names)
                    u.Settings_G.Button_Back.setEnabled(True)
                u.Settings_G.Button_Back.click()
            if(u.Type==0):
                g=Graphic_B(u.Settings_G)
            elif(u.Type==1):
                g=Graphic_S(u.Settings_G)         
            elif(u.Type==2):
                g=Graphic_P(u.Settings_G)          
            u.chartView_D.setChart(g.chart)
        self.D.D_Graphics.splitter1.setSizes(S1)
        self.D.D_Graphics.splitter2.setSizes(S2)
        self.D.D_Graphics.splitter3.setSizes(S3)
class D_Graphic(QWidget):
    def __init__(self, Settings_D): 
        QWidget.__init__(self)
        main_layout = QHBoxLayout()
        self.Settings_D = Settings_D

        self.splitter3 = QSplitter(QtCore.Qt.Vertical) 
        self.splitter1 = QSplitter()
        self.splitter2 = QSplitter()

        self.splitter3.addWidget(self.splitter1)
        self.splitter3.addWidget(self.splitter2)

        main_layout.addWidget(self.splitter3)
        self.setLayout(main_layout)
    def add_Graphic_on_DB(self,grapf):
        if(grapf.Type==0):
            g=Graphic_B(grapf.Settings_G)
        elif(grapf.Type==1):
            g=Graphic_S(grapf.Settings_G)        
        elif(grapf.Type==2):
            g=Graphic_P(grapf.Settings_G)           
        grapf.chartView_D.setChart(g.chart)
        w=grapf.chartView_D
        if(self.splitter1.count()%2==self.splitter2.count()%2):
            self.splitter1.addWidget(w)
        else:
            self.splitter2.addWidget(w)

class TabItemModel(QStandardItem):
    def __init__(self):
        super(TabItemModel, self).__init__()
        self.widget = None
        self.tab = None
        self.model_name=''

class ChartItemModel(TabItemModel):
    def __init__(self):
        super(ChartItemModel, self).__init__()
        self.setIcon(QIcon('icons/chart.svg'))
        self.setEditable(False)
        self.tab = None
        self.model_name = 'C'
class TableItemModel(TabItemModel):
    def __init__(self):
        super(TableItemModel, self).__init__()

        self.setIcon(QIcon('icons/table.svg'))
        self.setEditable(False)

        self.tab = None
        self.model_name = 'T'
class DashboardItemModel(TabItemModel):
    def __init__(self):
        super(DashboardItemModel, self).__init__()

        self.setIcon(QIcon('icons/dashboard.svg'))
        self.setEditable(False)

        self.tab = None
        self.model_name = 'D'

class MainWindowV2(QMainWindow):
    def __init__(self):
        super(MainWindowV2, self).__init__()
        p = self.palette() 
        p.setBrush(self.backgroundRole(), QtGui.QBrush(QColor("#B0C4DE")) )
        self.setPalette(p)

        self.model = QStandardItemModel(self)

        self.tabWidget = QTabWidget(self)        
        self.tabWidget.setTabPosition(QTabWidget.South)
        self.tabWidget.setTabsClosable(True) 
        self.tabWidget.setMovable(True)

        self.tabWidget.tabCloseRequested.connect(self.closeTabFromTabWidget)
        self.setCentralWidget(self.tabWidget)
                    
        logDockWidget = QDockWidget("Tabs", self)
        logDockWidget.setTitleBarWidget(QWidget())
        logDockWidget.setObjectName("LogDockWidget")
        logDockWidget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        logDockWidget.setFeatures(QDockWidget.DockWidgetMovable)
                
        self.navListView = QListView()
        self.navListView.setModel(self.model)

        logDockWidget.setWidget(self.navListView)
        self.addDockWidget(Qt.LeftDockWidgetArea, logDockWidget)

        self.createActions()

        tabsToolbar = self.addToolBar("Tabs")
        tabsToolbar.setObjectName("tabsToolBar")
        tabsToolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon) 
        
        for action in [self.newChartAction, self.newTableAction, self.newDashboardAction]:
            tabsToolbar.addAction(action)

        self.navListView.selectionModel().selectionChanged.connect(self.setActiveTab)
        self.newChart()
    def newChart(self):
        item = ChartItemModel()
        item.setText('New Chart')

        self.model.appendRow(item)
        widget = ChartWidget(item, self.tabWidget, self)
        self.tabWidget.addTab(widget, "New Chart")
        item.tab = widget   

        for i in range(self.model.rowCount()):
            el=self.model.itemFromIndex(self.model.index(i, 0))
            if(el.model_name =='D'):
                el.tab.Settings_D.update_all_charts(self.model)
        pass

    def newTable(self):
        item = TableItemModel()
        item.setText('New Table')
        self.model.appendRow(item)

        widget = TableWidget(item, self)
        self.tabWidget.addTab(widget, "New Table")

        item.tab = widget   
        pass

    def newDashboard(self):
        flag=True
        for i in range(self.model.rowCount()):
            el=self.model.itemFromIndex(self.model.index(i, 0))
            if(el.model_name =='D'):
                flag=False
        if(flag):
            item = DashboardItemModel()
            item.setText('New Dashboard')
            self.model.appendRow(item)
            widget = DashboardWidget(item, self)
            self.tabWidget.addTab(widget, "New Dashboard")

            item.tab = widget     
        
        pass
    def closeTabFromTabWidget(self, i):
        widget = self.tabWidget.widget(i)
        modelIndex = widget.model.index()
        self.navListView.model().removeRow(modelIndex.row())
        self.tabWidget.removeTab(self.tabWidget.indexOf(widget.model.tab))

        for i in range(self.model.rowCount()):
            el=self.model.itemFromIndex(self.model.index(i, 0))
            if(el.model_name =='D'):
                el.tab.Settings_D.update_all_charts(self.model)
                ind=el.tab.Settings_D.Setting_SH.index(widget.stackedWidget.widget(0))  

                el.tab.Settings_D.Setting_SH[ind].chartView_D.setParent(None)
                el.tab.Settings_D.Setting_SH.remove(el.tab.Settings_D.Setting_SH[ind])
                
    def createActions(self):
        self.newChartAction = QAction("New Chart", self,  icon = QIcon('icons/chart.svg'))        
        self.newChartAction.triggered.connect(self.newChart)

        self.newTableAction = QAction("New Table", self, icon = QIcon('icons/table.svg'))        
        self.newTableAction.triggered.connect(self.newTable)

        self.newDashboardAction = QAction("New Dashboard", self,   icon = QIcon('icons/dashboard.svg'))        
        self.newDashboardAction.triggered.connect(self.newDashboard)

    def setActiveTab(self):
        selectedModel = self.model.itemFromIndex(self.navListView.selectedIndexes()[0])
        self.tabWidget.setCurrentWidget(selectedModel.tab)
        
if  __name__== "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindowV2()
    window.setWindowTitle("Приложение для визуализации динамических данных")
    window.resize(900, 600)
    window.show()   
    sys.exit(app.exec_())












