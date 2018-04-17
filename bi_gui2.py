import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
from os import sep, path
import numpy as np
import resource_rc
import sys
import time
from multiprocessing import Pool, cpu_count
from threading import  Thread
import dill
import pyqtgraph as pg

class FileManagement:
    def __init__(self):
        self.workSheet = {'excel': None,
                          'sheets': list(),
                          'df':pd.DataFrame(),
                          'selectedColumns':list(),
                          'selectedRows':list(),
                          'columnsValue': dict(),
                          'currentSelectedFilter':None,
                          'previousCurrentRow':None,
                          'filteredColumns':set(),
                          'dimensions':list(),
                          'measurements':list(),
                          'grouped':{'df':pd.DataFrame(), 'columns':list(), 'filterGrouped':pd.DataFrame(), 'graph':tuple()}}

    def readExcel(self, filePath):
        '''Read excel file.'''
        self.workSheet['excel'] = pd.ExcelFile(filePath)

    def getSheet(self):
        '''Get sheets list.'''
        self.workSheet['sheets'] = self.workSheet['excel'].sheet_names

    def readSheet(self, sheetName):
        '''Read a table from sheetname.'''
        if 'df' in self.workSheet:
            self.workSheet = {'excel':self.workSheet['excel'],
                              'sheets': self.workSheet['sheets'],
                              'df':self.workSheet['excel'].parse(sheetName.text()),
                              'selectedColumns':list(),
                              'selectedRows':list(),
                              'columnsValue':dict(),
                              'currentSelectedFilter':None,
                              'previousCurrentRow':None,
                              'filteredColumns':set(),
                              'dimensions':list(),
                              'measurements':list(),
                              'grouped':{'df':pd.DataFrame(), 'columns':list(), 'filterGrouped':pd.DataFrame(), 'graph':tuple()}}
            self.workSheet['grouped']['filterGrouped'] = self.workSheet['df']
        else:
            self.workSheet['df'] = self.workSheet['excel'].parse(sheetName.text())
            self.workSheet['grouped']['filterGrouped'] = self.workSheet['df']

    def saveFile(self, filePath, data):
        with open(filePath, 'wb') as file:
            dill.dump(data, file)

    def loadFile(self, filePath):
        with open(filePath, 'rb') as file:
            self.workSheet = dill.load(file)

    def isFileExist(self, filePath):
        '''Check whether file exist.'''
        if path.exists(filePath):
           self.loadFile(filePath)

    def toExcel(self, filePath, sheetName, df, graph):
        '''Export to Excel Format.'''
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
        df.to_excel(writer, sheet_name= sheetName)
        # Access the XlsxWriter workbook and worksheet objects from the dataframe.
        workbook = writer.book
        worksheet = writer.sheets[sheetName]
        # Create a chart object.
        chart = workbook.add_chart({'type': 'pie'})
        rowNumbers = df.shape[0]
        columnsNumbers = df.shape[1]
        categoriesRange = '={0}!B1:D{1}'.format(sheetName, columnsNumbers)
        valuesRange = '={0}!B2:D{1}'.format(sheetName, rowNumbers + 2)
        chart.add_series({
            'categories': categoriesRange,
            'values': valuesRange
        })
        # Insert the chart into the worksheet.
        worksheet.insert_chart('E4', chart)
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

class DataOrganization(FileManagement):
    def __init__(self):
        super(DataOrganization, self).__init__()

    def addColumns(self, selectedColumn):
        if 'selectedColumns' in self.workSheet:
            self.workSheet['selectedColumns'].append(selectedColumn)
            # self.getGroupValue(selectedColumn)
        else:
            self.workSheet['selectedColumns'] = [selectedColumn]

    def addRows(self, selectedRows):
        if 'selectedRows' in self.workSheet:
            self.workSheet['selectedRows'].append(selectedRows)
        else:
            self.workSheet['selectedRows'] = [selectedRows]

    def classifyDimensionMeasurement(self, df):
        '''Classify dimension and measurement.'''
        columnsType = {'dimensions':[], 'measurements':[]}

        for eachColumn, eachDataType in zip(df.columns, df.dtypes):
            if eachDataType != 'float64':
                columnsType['dimensions'].append(eachColumn)
            else:
                columnsType['measurements'].append(eachColumn)

        self.workSheet['dimensions'] = columnsType['dimensions']
        self.workSheet['measurements'] = columnsType['measurements']

    def _isDiscrete(self, measurement):
        '''Check whether measurement is discrete value.'''
        return self.workSheet['df'][measurement].dtypes != 'float64'

    def getGroupValue(self, groupedDF):
        self.workSheet['grouped']['columns'] = {}
        groupedIndex = np.array([*groupedDF.index.values])
        for indexNum, eachIndexName in enumerate(self.workSheet['grouped']['df'].index.names):
            self.workSheet['grouped']['columns'][eachIndexName] = list(np.unique(groupedIndex[:, indexNum]))

    def getColumnValue(self, column):
        values = np.array(self.workSheet['df'][column])
        self.workSheet['columnsValue'][column] = list(np.unique(values))

    def filterByColumns(self, df, filterBy, filterValue):
        return df.loc[df[filterBy] == filterValue]

    def filterByIndex(self, groupedDF, filterValue):
        self.workSheet['grouped']['filterGrouped'] = groupedDF.filter(like=filterValue, axis=0)

    def filterGrouped(self, filterGrouped, indexName, indexValue):
        index = list(filterGrouped.index.names).index(indexName)
        boolArray = np.array([*filterGrouped.index.values])[:, index] == indexValue
        dfIndex = filterGrouped.index.values[boolArray]
        self.workSheet['grouped']['filterGrouped'] = filterGrouped.filter(items=dfIndex, axis=0)

    def groupData(self, df, dimensions, measurement):
        if not self._isDiscrete(measurement):
            self.workSheet['grouped'] = {'df':pd.DataFrame(df.groupby(dimensions)[[measurement]].sum())}
        else:
            self.workSheet['grouped'] = {'df': pd.DataFrame(df.groupby(dimensions).size(), columns=['Amount'])}
        self.workSheet['grouped']['filterGrouped'] = self.workSheet['grouped']

    def rangeSelect(self, df, startRow=0, stopRow=None, startColumn=0, stopColumn=None):
        '''Select Columns and Row Range.'''
        tmpDF = df

        if(startRow  != 0  and stopRow == None):
            tmpDF =  df[startRow:]
        elif(startRow != 0 and stopRow != None):
            tmpDF = df[startRow:stopRow]
        elif(startRow == 0 and stopRow != None):
            tmpDF = df[:stopRow]

        if (startColumn != 0 and stopColumn == None):
            tmpDF = df[startRow:]
        elif (startColumn != 0 and stopColumn != None):
            tmpDF = df[startRow:stopRow]
        elif (startColumn == 0 and stopColumn != None):
            tmpDF = df[:stopRow]
        return  tmpDF

    def multiThread(self, worker, args):
         for eacharg in args:
            p = Thread(target=worker, args=(eacharg,))
            p.start()



class Ui_MainWindow(DataOrganization):
    numColumnsList = 0
    numRowsList = 0

    def __init__(self):
        DataOrganization.__init__(self)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1261, 838)
        MainWindow.setWindowTitle('Analytic Tool')
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")

        self.scrollArea = QtWidgets.QScrollArea(self.centralwidget)
        self.scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")

        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 541, 765))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")

        self.gridLayout_2 = QtWidgets.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.gridLayout_2.setColumnMinimumWidth(4, 1)

        self.rowLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.rowLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.rowLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(181, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.rowLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.rowLabel.setScaledContents(False)
        self.rowLabel.setObjectName("rowLabel")
        self.rowLabel.setText("Rows")
        self.gridLayout_2.addWidget(self.rowLabel, 4, 1)

        self.rowListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.rowListWidget.setObjectName("rowListWidget")
        self.rowListWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.rowListWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.rowListWidget.currentItemChanged.connect(self.displayRowsFilter)
        self.gridLayout_2.addWidget(self.rowListWidget, 5, 1)

        self.columnLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.columnLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.columnLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(181, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.columnLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.columnLabel.setScaledContents(False)
        self.columnLabel.setObjectName("columnLabel")
        self.columnLabel.setText("Columns")
        self.gridLayout_2.addWidget(self.columnLabel, 2, 1, 1, 1)

        self.columnListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.columnListWidget.setObjectName("columnListWidget")
        self.columnListWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.columnListWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.columnListWidget.currentItemChanged.connect(self.displayColumnFilter)
        self.gridLayout_2.addWidget(self.columnListWidget, 3, 1)

        self.filterLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.filterLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.filterLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(181, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.filterLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.filterLabel.setScaledContents(False)
        self.filterLabel.setObjectName("filterLabel")
        self.filterLabel.setText("Filter")
        self.gridLayout_2.addWidget(self.filterLabel, 0, 1)

        self.filterListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.filterListWidget.setObjectName("filterListWidget")
        self.gridLayout_2.addWidget(self.filterListWidget, 1, 1)

        self.sheetLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.sheetLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.sheetLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(181, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.sheetLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.sheetLabel.setScaledContents(False)
        self.sheetLabel.setObjectName("sheetLabel")
        self.sheetLabel.setText("Sheets")
        self.gridLayout_2.addWidget(self.sheetLabel, 0, 0)

        # Set sheet list widget
        self.sheetListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.sheetListWidget.setObjectName("sheetList")
        self.sheetListWidget.itemActivated.connect(self.displayDimensionsMeasurements)
        self.gridLayout_2.addWidget(self.sheetListWidget, 1, 0)

        self.measurementLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.measurementLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.measurementLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(181, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.measurementLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.measurementLabel.setScaledContents(False)
        self.measurementLabel.setObjectName("measurementLabel")
        self.measurementLabel.setText("Measurements")
        self.gridLayout_2.addWidget(self.measurementLabel, 4, 0)

        self.measurementWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.measurementWidget.setObjectName("measurementsList")
        self.measurementWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.measurementWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.gridLayout_2.addWidget(self.measurementWidget, 5, 0)


        self.dimensionLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.dimensionLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.dimensionLabel.setStyleSheet("\n"
                                          "background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(181, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.dimensionLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.dimensionLabel.setScaledContents(False)
        self.dimensionLabel.setObjectName("dimensionLabel")
        self.dimensionLabel.setText("Dimensions")
        self.gridLayout_2.addWidget(self.dimensionLabel, 2, 0)

        self.dimensionWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.dimensionWidget.setObjectName("dimensionList")
        self.dimensionWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.dimensionWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.gridLayout_2.addWidget(self.dimensionWidget, 3, 0)


        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.gridLayout.addWidget(self.scrollArea, 0, 0)

        self.scrollArea_2 = QtWidgets.QScrollArea(self.centralwidget)
        self.scrollArea_2.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea_2.setWidgetResizable(True)
        self.scrollArea_2.setObjectName("scrollArea_2")
        self.scrollAreaWidgetContents_2 = QtWidgets.QWidget()
        self.scrollAreaWidgetContents_2.setGeometry(QtCore.QRect(0, 0, 342, 765))
        self.scrollAreaWidgetContents_2.setObjectName("scrollAreaWidgetContents_2")

        self.gridLayout_3 = QtWidgets.QGridLayout(self.scrollAreaWidgetContents_2)
        self.gridLayout_3.setObjectName("gridLayout_3")

        self.continuousChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.continuousChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.continuousChart.setToolButtonStyle(3)  # Set text position under icon
        self.continuousChart.setText("Continuous")
        continuousChartIcon = QtGui.QIcon()
        continuousChartIcon.addPixmap(QtGui.QPixmap(":/resource/continuousChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.continuousChart.setIcon(continuousChartIcon)
        self.continuousChart.setIconSize(QtCore.QSize(70, 70))
        self.continuousChart.setObjectName("continuousChart")
        self.gridLayout_3.addWidget(self.continuousChart, 4, 0)

        self.blockChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.blockChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.blockChart.setToolButtonStyle(3)  # Set text position under icon
        self.blockChart.setText("Block")
        blockChartIcon = QtGui.QIcon()
        blockChartIcon.addPixmap(QtGui.QPixmap(":/resource/blockChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.blockChart.setIcon(blockChartIcon)
        self.blockChart.setIconSize(QtCore.QSize(70, 70))
        self.blockChart.setObjectName("blockChart")
        self.gridLayout_3.addWidget(self.blockChart, 2, 0)

        self.horizontalBarChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.horizontalBarChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.horizontalBarChart.setToolButtonStyle(3)  # Set text position under icon
        self.horizontalBarChart.setText("Horizontal Bar")
        horizontalBarIcon = QtGui.QIcon()
        horizontalBarIcon.addPixmap(QtGui.QPixmap(":/resource/horizontalBarChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.horizontalBarChart.setIcon(horizontalBarIcon)
        self.horizontalBarChart.setIconSize(QtCore.QSize(70, 70))
        self.horizontalBarChart.setObjectName("horizontalBarChart")
        self.gridLayout_3.addWidget(self.horizontalBarChart, 3, 0)

        self.pieChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.pieChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.pieChart.setToolButtonStyle(3)  # Set text position under icon
        self.pieChart.setText("Pie Chart")
        pieChartIcon = QtGui.QIcon()
        pieChartIcon.addPixmap(QtGui.QPixmap(":/resource/pieChart.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pieChart.setIcon(pieChartIcon)
        self.pieChart.setIconSize(QtCore.QSize(70, 70))
        self.pieChart.setObjectName("pieChart")
        self.gridLayout_3.addWidget(self.pieChart, 0, 0)


        self.circleChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.circleChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.circleChart.setToolButtonStyle(3)  # Set text position under icon
        self.circleChart.setText("Circle")
        circleChartIcon = QtGui.QIcon()
        circleChartIcon.addPixmap(QtGui.QPixmap(":/resource/circleView.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.circleChart.setIcon(circleChartIcon)
        self.circleChart.setIconSize(QtCore.QSize(70, 70))
        self.circleChart.setObjectName("circleChart")
        self.gridLayout_3.addWidget(self.circleChart, 5, 0)

        self.table = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.table.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.table.setToolButtonStyle(3)  # Set text position under icon
        self.table.setText("Table")
        tableIcon = QtGui.QIcon()
        tableIcon.addPixmap(QtGui.QPixmap(":/resource/tableChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.table.setIcon(tableIcon)
        self.table.setIconSize(QtCore.QSize(70, 70))
        self.table.setObjectName("table")
        self.gridLayout_3.addWidget(self.table, 1, 0)

        self.histogram = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.histogram.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.histogram.setToolButtonStyle(3)  # Set text position under icon
        self.histogram.setText("Histogram")
        histrogramIcon = QtGui.QIcon()
        histrogramIcon.addPixmap(QtGui.QPixmap(":/resource/histrogramChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.histogram.setIcon(histrogramIcon)
        self.histogram.setIconSize(QtCore.QSize(70, 70))
        self.histogram.setObjectName("histogram")
        self.gridLayout_3.addWidget(self.histogram, 0, 1)

        self.scatter = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        scatterIcon = QtGui.QIcon()
        self.scatter.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.scatter.setToolButtonStyle(3)  # Set text position under icon
        self.scatter.setText("Scatter")
        scatterIcon.addPixmap(QtGui.QPixmap(":/resource/scatterChart.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.scatter.setIcon(scatterIcon)
        self.scatter.setIconSize(QtCore.QSize(70, 70))
        self.scatter.setObjectName("scatter")
        self.gridLayout_3.addWidget(self.scatter, 1, 1)

        self.stackBarChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.stackBarChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.stackBarChart.setToolButtonStyle(3)  # Set text position under icon
        self.stackBarChart.setText("Stack Bar")
        stackBarIcon = QtGui.QIcon()
        stackBarIcon.addPixmap(QtGui.QPixmap(":/resource/stackBarChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.stackBarChart.setIcon(stackBarIcon)
        self.stackBarChart.setIconSize(QtCore.QSize(70, 70))
        self.stackBarChart.setObjectName("stackBarChart")
        self.gridLayout_3.addWidget(self.stackBarChart, 2, 1)

        self.sidebysideBarChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.sidebysideBarChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.sidebysideBarChart.setToolButtonStyle(3)  # Set text position under icon
        self.sidebysideBarChart.setText("Side By Side")
        sidebysideIcon = QtGui.QIcon()
        sidebysideIcon.addPixmap(QtGui.QPixmap(":/resource/sidebysideBarChart.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.sidebysideBarChart.setIcon(sidebysideIcon)
        self.sidebysideBarChart.setIconSize(QtCore.QSize(70, 70))
        self.sidebysideBarChart.setObjectName("sidebysideBarChart")
        self.gridLayout_3.addWidget(self.sidebysideBarChart, 3, 1)

        self.bubbleChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.bubbleChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.bubbleChart.setToolButtonStyle(3)  # Set text position under icon
        self.bubbleChart.setText("Bubble")
        bubbleChartIcon = QtGui.QIcon()
        bubbleChartIcon.addPixmap(QtGui.QPixmap(":/resource/bubbleChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.bubbleChart.setIcon(bubbleChartIcon)
        self.bubbleChart.setIconSize(QtCore.QSize(70, 70))
        self.bubbleChart.setObjectName("bubbleChart")
        self.gridLayout_3.addWidget(self.bubbleChart, 4, 1)

        self.dualChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.dualChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.dualChart.setToolButtonStyle(3)  # Set text position under icon
        self.dualChart.setText("Dual Chart")
        dualChartIcon = QtGui.QIcon()
        dualChartIcon.addPixmap(QtGui.QPixmap(":/resource/dualChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.dualChart.setIcon(dualChartIcon)
        self.dualChart.setIconSize(QtCore.QSize(70, 70))
        self.dualChart.setObjectName("dualChart")
        self.gridLayout_3.addWidget(self.dualChart, 5, 1)

        self.scrollArea_2.setWidget(self.scrollAreaWidgetContents_2)
        self.gridLayout.addWidget(self.scrollArea_2, 0, 2)
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setBaseSize(QtCore.QSize(2000, 2000))
        self.tabWidget.setObjectName("tabWidget")

        self.chartTab = QtWidgets.QWidget()
        self.chartTab.setObjectName("chartTab")
        self.plotLayout = QtWidgets.QVBoxLayout()
        self.plotWidget = pg.PlotWidget(name='Plot')
        self.plotLayout.addWidget(self.plotWidget)
        self.chartTab.setLayout(self.plotLayout)
        self.tabWidget.addTab(self.chartTab, "")
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.chartTab), "Chart")

        self.tableTab = QtWidgets.QWidget()
        self.tableTab.setObjectName("tableTab")
        self.tableLayout = QtWidgets.QVBoxLayout()
        self.tableWidget = pg.TableWidget()
        self.tableLayout.addWidget(self.tableWidget)
        self.tableTab.setLayout(self.tableLayout)
        self.tabWidget.addTab(self.tableTab, "")

        self.gridLayout.addWidget(self.tabWidget, 0, 1)

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.chartTab), "Chart")


        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tableTab), "Table")



        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1261, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuFile.setTitle("File")

        self.actionOpen = QtWidgets.QAction(MainWindow)
        self.actionOpen.setObjectName("actionImport")
        self.actionOpen.setText("Open")
        self.actionOpen.setShortcut("Ctrl+O")
        self.actionOpen.triggered.connect(self.openFileDialog)

        self.actionImport = QtWidgets.QAction(MainWindow)
        self.actionImport.setObjectName("actionImport")
        self.actionImport.setText("Import")
        self.actionImport.setShortcut("Ctrl+O")
        self.actionImport.triggered.connect(self.openFileDialog)

        self.actionSave = QtWidgets.QAction(MainWindow)
        self.actionSave.setObjectName("actionSave")
        self.actionSave.setText("Save")
        self.actionSave.setShortcut("Ctrl+S")
        self.actionSave.triggered.connect(self.saveFileDialog)

        self.actionExport = QtWidgets.QAction(MainWindow)
        self.actionExport.setObjectName("actionExport")
        self.actionExport.setText("Export")
        self.actionExport.setShortcut("Ctrl+E")
        self.actionExport.triggered.connect(self.exportFileDialog)

        self.menuFile.addAction(self.actionOpen)
        self.menuFile.addAction(self.actionImport)
        self.menuFile.addAction(self.actionSave)
        self.menuFile.addAction(self.actionExport)
        self.menubar.addAction(self.menuFile.menuAction())

        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def addDropDownWidget(self):
        self.dropdownButton = QtWidgets.QToolButton(self)
        self.dropdownButton.setPopupMode(QtWidgets.QToolButton.MenuButtonPopup)
        self.dropdownButton.setMenu(QtWidgets.QMenu(self.dropdownButton))
        self.textBox = QtWidgets.QTextBrowser(self)
        action = QtWidgets.QWidgetAction(self.dropdownButton)
        action.setDefaultWidget(self.textBox)
        self.dropdownButton.menu().addAction(action)
        return  self.dropdownButton

    def addListObject(self, iterableObject, widget):
        for eachObject in iterableObject:
            self.item = QtWidgets.QListWidgetItem(eachObject)
            widget.addItem(self.item)

    def openFileDialog(self):
        directory = QtWidgets.QFileDialog.getOpenFileName()
        directory = str(directory[0])
        if directory != '':
            fileName = directory.split('/')[-1]
            extension = fileName.split('.')[-1]
            self.sheetListWidget.clear()
            self.dimensionWidget.clear()
            self.measurementWidget.clear()
            self.rowListWidget.clear()
            self.columnListWidget.clear()
            self.filterListWidget.clear()

            if extension == 'xlsx':
                self.readExcel(fileName)
                self.getSheet()
                self.addListObject(self.workSheet['sheets'], self.sheetListWidget)

            elif extension == 'pkl':
                self.loadFile(fileName)

                self.addListObject(self.workSheet['sheets'], self.sheetListWidget)

                self.addListObject(self.workSheet['dimensions'], self.dimensionWidget)

                self.addListObject(self.workSheet['measurements'], self.measurementWidget)

                self.addListObject(self.workSheet['selectedColumns'], self.columnListWidget)

                self.addListObject(self.workSheet['selectedRows'], self.rowListWidget)

                if self.workSheet['currentSelectedFilter'] is not None:
                    checkBoxes = self.getCheckBoxes(self.workSheet['columnsValue'][self.workSheet['currentSelectedFilter']])
                    for eachCheckBox in checkBoxes:
                        self.filterListWidget.addItem(eachCheckBox)

    def saveFileDialog(self):
        fileName = QtWidgets.QFileDialog.getSaveFileName()
        fileName = fileName[0]
        self.getState()

        if fileName != '':
            fileName = fileName.split('/')[-1] + '.pkl'
            self.saveFile(fileName, self.workSheet)

    def exportFileDialog(self):
        fileName = QtWidgets.QFileDialog.getSaveFileName()
        fileName = fileName[0].split('/')[-1]
        self.toExcel(fileName, 'sheet', self.workSheet['grouped']['filterGrouped'], self.workSheet['grouped']['graph'][0])

    def displayDimensionsMeasurements(self, sheet):
        self.readSheet(sheet)
        self.dimensionWidget.clear()
        self.measurementWidget.clear()
        self.columnListWidget.clear()
        self.rowListWidget.clear()
        self.filterListWidget.clear()
        self.classifyDimensionMeasurement(self.workSheet["df"])
        self.addListObject(self.workSheet['dimensions'], self.dimensionWidget)
        self.addListObject(self.workSheet['measurements'], self.measurementWidget)
        self.multiThread(self.getColumnValue, list(self.workSheet['df'].columns))

    def getCheckBoxes(self, column):
        items = list()
        column = list(map(str, column))
        for eachFilter in column:
            item = QtWidgets.QListWidgetItem()
            item.setText(eachFilter)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            if item.text() not in self.workSheet['filteredColumns']:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            items.append(item)
        return items

    def getCheckBoxesState(self):
        for eachRow in range(self.filterListWidget.count()):
            item = self.filterListWidget.item(eachRow)
            if item.checkState() != 2:
                self.workSheet['filteredColumns'].add(item.text())

    def deleteFilteredColumns(self, column):
        for eachFilter in list(self.workSheet['filteredColumns']):
            print(eachFilter)
            if eachFilter in self.workSheet['columnsValue'][column]:
                self.workSheet['filteredColumns'].remove(eachFilter)


    def displayColumnFilter(self):
        if self.columnListWidget.currentRow() != -1:
            column = self.columnListWidget.item(self.columnListWidget.currentRow()).text()
            self.workSheet['currentSelectedFilter'] = column
            numColumnList = self.columnListWidget.count()
            self.getState()
            if numColumnList > self.numColumnsList:
                self.numColumnsList = numColumnList

            if numColumnList != 1:
                if self.workSheet['previousCurrentRow'] in self.workSheet['dimensions'] or \
                                self.workSheet['previousCurrentRow'] in self.workSheet['measurements'] or \
                                self.workSheet['previousCurrentRow'] in self.workSheet['selectedRows']:

                    self.deleteFilteredColumns(self.workSheet['previousCurrentRow'])
                else:
                    self.getCheckBoxesState()

            self.filterListWidget.clear()
            self.numColumnsList = numColumnList
            self.workSheet['previousCurrentRow'] = column
            checkBoxes = self.getCheckBoxes(self.workSheet['columnsValue'][column])

            for eachCheckBox in checkBoxes:
                self.filterListWidget.addItem(eachCheckBox)

        else:
            self.filterListWidget.clear()
            self.workSheet['filteredColumns'] = set()
            self.workSheet['currentSelectedFilter'] = None
            self.workSheet['previousCurrentRow'] = None


    def displayRowsFilter(self):
        if self.rowListWidget.currentRow() != -1:
            column = self.rowListWidget.item(self.rowListWidget.currentRow()).text()
            self.workSheet['currentSelectedFilter'] = column
            numRowsList = self.rowListWidget.count()
            self.getState()
            if numRowsList > self.numRowsList:
                self.numRowsList = numRowsList

            if numRowsList != 1:
                if self.workSheet['previousCurrentRow'] in self.workSheet['dimensions'] or \
                                self.workSheet['previousCurrentRow'] in self.workSheet['measurements'] or \
                                self.workSheet['previousCurrentRow'] in self.workSheet['selectedColumns']:

                    self.deleteFilteredColumns(self.workSheet['previousCurrentRow'])
                else:
                    self.getCheckBoxesState()

            self.filterListWidget.clear()
            self.numRowsList = numRowsList
            self.workSheet['previousCurrentRow'] = column
            checkBoxes = self.getCheckBoxes(self.workSheet['columnsValue'][column])

            for eachCheckBox in checkBoxes:
                self.filterListWidget.addItem(eachCheckBox)

        else:
            self.filterListWidget.clear()
            self.workSheet['filteredColumns'] = set()
            self.workSheet['currentSelectedFilter'] = None
            self.workSheet['previousCurrentRow'] = None

    def getState(self):
        self.workSheet['dimensions'] = list()
        for eachRow in range(self.dimensionWidget.count()):
            self.workSheet['dimensions'].append(self.dimensionWidget.item(eachRow).text())

        self.workSheet['measurements'] = list()
        for eachRow in range(self.measurementWidget.count()):
            self.workSheet['measurements'].append(self.measurementWidget.item(eachRow).text())

        self.workSheet['selectedColumns'] = list()
        for eachRow in range(self.columnListWidget.count()):
            self.workSheet['selectedColumns'].append(self.columnListWidget.item(eachRow).text())

        self.workSheet['selectedRows'] = list()
        for eachRow in range(self.rowListWidget.count()):
            self.workSheet['selectedRows'].append(self.rowListWidget.item(eachRow).text())


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

