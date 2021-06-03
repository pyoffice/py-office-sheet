import gc, sys, joblib
from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import *
import pandas as pd
import numpy as np


def spreadsheet(screen_width,screen_height):
    global saved_file
    class MyTableModel(QAbstractTableModel): # numpy array model
        # when tableView is rendered, data of each cell would be called through self.data
        # this method avoid creating widgets for each cell e.g. QtableWiget
        # by storing data in numpy array, fast processing of data could be performed
        def __init__(self, array, headers= None,parent=None):
            super().__init__(parent)
            self.array = array # call current array later through tableWidget.model().array
            self.headers = headers

        def headerData(self, section: int, orientation: Qt.Orientation, role: int):
            if role == Qt.DisplayRole:
                if orientation == Qt.Horizontal:
                    if self.headers != None:
                        try:
                            return self.headers[section]  # column
                        except :
                            return str(section)
                    else:
                        return str(section)
                else:
                    return str(section)  # row

        def columnCount(self, parent=None):
            return len(self.array[0])

        def rowCount(self, parent=None):
            return len(self.array)

        def data(self, index: QModelIndex, role: int): # value called on render
            if role == Qt.DisplayRole or role == Qt.EditRole:
                row = index.row()
                col = index.column()
                return str(self.array[row][col]) # return value of cell

        def setData(self, index, value, role): # set data would be called everytime user edit cell
            if role == Qt.EditRole:
                saved_file = False
                self.array[index.row()][index.column()] = value # asign new data to array
                return True

        def flags(self, index):
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

#####################r read stuff #########################################

    def pick_sys_file(filter="All files (*)"):
        from mimetypes import guess_type
        if filter == False:
            filter = "All files (*)"
        file_name, filter = QFileDialog.getOpenFileName(menuWidget, 'Open File', 'c://', filter=filter)
        type = guess_type(file_name)
        print(type)
        if 'text/csv' in type:
            opencsv(file_name)
        elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in type:
            openexel(file_name)
        elif type[0] == None:
            return None
        else:
            m = QMessageBox()
            m.setText('File type not support\n'+str(type[0]))
            m.exec_()
            return None
        column = tableWidget.model().columnCount()
        columnCount.setRange(column, column+10000)
        columnCount.setValue(column)
        row = tableWidget.model().rowCount()
        rowCount.setRange(row, row +10000)
        rowCount.setValue(row)

        menuSave.setObjectName(file_name)
        return file_name

    
    def opencsv(file):
        #global pandas_data
        data = pd.read_csv(file)
        pandas_data = data
        headers = list(data.keys())
        data = np.array(data)
        tableWidget.setModel(MyTableModel(data,headers=headers))

    def openexel(file):
        #global pandas_data
        data = pd.read_excel(file,sheet_name=None)
        sheets = list(data.keys())
        if len(sheets)>1:
            pass
        pandas_data = data
        data = np.array(data[sheets[0]])
        tableWidget.setModel(MyTableModel(data))
    
    def importJoblib():
        filter = 'Python Object(*.npobj *.pdobj)'
        file_name, filter = QFileDialog.getOpenFileName(menuWidget, 'Open File', 'c://', filter=filter)
        if '.npobj' in file_name:
            data = joblib.load(file_name)
            header = None
        if '.pdobj' in file_name:
            dataframe = joblib.load(file_name)
            header = list(dataframe.keys())
            data = np.array(dataframe)
        
        tableWidget.setModel(MyTableModel(data,headers=header))

#################### save stuff #########################
    def saveFile(directory=None,saveAs = False):
        filter = "Pandas Object(*.pdobj);;Numpy Object(*.npobj);;CSV(*.csv);;JSON(*.json);;PDF( *.pdf)"
        if saveAs:
            directory, filter = QFileDialog.getSaveFileName(menuWidget, 'Save File', 'c://',filter=filter)

        elif directory==None and len(menuSave.objectName())<3:
            directory, filter=QFileDialog.getSaveFileName(menuWidget,'Save File','c://',filter=filter)

        elif directory!=None:
            directory = str(directory)

            if len(directory.split('.'))>1:
                filter = '.'+directory.split('.')[-1]
            else:
                filter = '.pdobj'

        elif len(menuSave.objectName())>3:
            directory= menuSave.objectName()
            filter = directory.split('.')[-1]

        headers = tableWidget.model().headers

        data = pd.DataFrame(tableWidget.model().array,columns=headers)
        print(filter)

        if '.pdobj' in filter:
            directory+='.pdobj'
            joblib.dump(data,directory,9)
        elif '.csv' in filter:
            data.to_csv(directory, index=False)

        elif '.json' in filter:
            directory = directory+'.json' if '.json' not in directory else directory
            out = data.to_json(directory,index=False,orient= 'table')

        elif '.html' in filter:
            data.to_html(directory,index=False)

        elif'.xlsx'in directory:
            directory = directory+'.xlsx' if '.xlsx' not in directory else directory
            data.to_excel(directory,index=False)
            
        elif '.h5'in filter:
            data.to_hdf(directory,key='0')

    def exportJoblib(dtype='np'): # export numpy array by joblib
        filename, filter = QFileDialog.getSaveFileName(menuWidget, 'Save File', 'c://')
        if dtype=='np':
            filename += '.npobj'
            array = tableWidget.model().array
            joblib.dump(array,filename,9)
        elif dtype == 'pd':
            filename +='.pdobj'
            headers = tableWidget.model().headers
            data = pd.DataFrame(tableWidget.model().array,columns=headers)
            joblib.dump(data,filename,9)

        

################# local functions ########################

    def changeEncodeMethod(a):
        for i in [utf_8,utf_7,utf_16,utf_32,ascii,big5,cp,cp037]:
            i.setDisabled(False)
        for i in a:
            i.setDisabled(True)
            encode.setObjectName(i.text())


    def spreadsheetCommand():
        import spreadsheet_command
        spreadsheet_command.main(commandBar,printOutLabel,tableWidget,MyTableModel)

    def commandHandler(event):
        text = event.text()
        if text == '[':
            commandBar.insert(']')
            commandBar.cursorBackward(False, 1)
        elif text == '(':
            commandBar.insert(')')
            commandBar.cursorBackward(False, 1)
        elif text == '{':
            commandBar.insert('}')
            commandBar.cursorBackward(False, 1)
        elif text == '(':
            commandBar.insert(')')
            commandBar.cursorBackward(False, 1)
        elif text == "'":
            commandBar.insert("'")
            commandBar.cursorBackward(False, 1)
        elif text == '"':
            commandBar.insert('"')
            commandBar.cursorBackward(False, 1)


    def manageFunction():
        manage_function_box = QDialog()
        manage_function_box.setWindowTitle('Manage Functions')
        manage_function_box.setWindowIcon(QIcon('pic/icon/main.png'))
        manage_function_box.setGeometry(screen_width/4,screen_height/16
                                        ,screen_width/2,screen_height/1.2)
        layout = QVBoxLayout()
        text = QTextEdit()
        text.setStyleSheet('background-color:white;')
        o = open('spreadsheet_command.py', 'r')

        text.setTextColor(Qt.red)
        text.setText(o.readline())
        text.moveCursor(QTextCursor.End)
        text.setTextColor(Qt.black)
        text.insertPlainText(o.read())
        o.close()
        layout.addWidget(text)
        manage_function_box.setLayout(layout)
        manage_function_box.exec_()
        o = open('spreadsheet_command.py','w')
        o.write(text.toPlainText())
        o.close()
        manage_function_box.deleteLater()
        layout.deleteLater()

    def analyze():
        box = QDialog()
        box.setWindowTitle('Analyze data')
        layout = QFormLayout()
        graphOption = QComboBox()
        graphOption.addItems(['matplotlib','pyqtgraph'])
        layout.addRow(QLabel('Library: '),graphOption)
        box.setLayout(layout)
        def matplot(data):
            box = QDialog()
            layout = QVBoxLayout()
            box.setLayout(layout)
            try:
                from matplotlib.figure import Figure
                from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
                from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT
            except Exception as e:  # it is mostly due to pillow
                from sys import executable
                from subprocess import check_call
                check_call([executable, "-m", "pip", "install", '-U', 'pillow'])
                i = QMessageBox()  # upgrade pillow may resolve the problem
                i.setText(e + '\nnow upgrading pillow')
                i.exec_()
            figure = Figure()
            canvas = FigureCanvasQTAgg(figure)
            toolbar = NavigationToolbar2QT(canvas, parent=canvas)
            layout.addWidget(toolbar)
            layout.addWidget(canvas)
            box.exec_()

        box.exec_()



#################################################
    data = np.array([['']*30]*30)
    model = MyTableModel(data)
    table_tab_box = QVBoxLayout()
    tableWidget = QTableView()
    tableWidget.setModel(model)
    tableWidget.horizontalHeader().stretchLastSection()
    tableWidget.setStyleSheet('background-color:white;color:black;')

############################ Menu section ######################################3

    menuWidget = QWidget()
    menuLayout = QVBoxLayout()
    menuWidget.setLayout(menuLayout)
    menuWidget.setMinimumHeight(screen_height /8)

################## command layout ##########################
    menuLayout_command = QVBoxLayout()
    commandBar = QLineEdit()
    commandBar.keyReleaseEvent = commandHandler
    commandBar.setMaximumWidth(int(screen_width/1.05))
    commandBar.returnPressed.connect(spreadsheetCommand)
    functionCompleter = QCompleter(
        ['replaceColumn()'])
    functionCompleter.setCaseSensitivity(Qt.CaseInsensitive)
    commandBar.setCompleter(functionCompleter)
    printOutLabel = QLabel()
    menuLayout_command.addWidget(printOutLabel)
    menuLayout_command.addWidget(commandBar)

    commandWidget = QWidget()
    commandWidget.setLayout(menuLayout_command)

################### home layout ###############################
    menuLayout_home = QVBoxLayout()
    menuLayout_home.setAlignment(Qt.AlignTop)

    def resizeTableToContent():
        tableWidget.resizeRowsToContents()
        tableWidget.resizeColumnsToContents()

    fontTypeCB = QComboBox()


    fontSizeCB = QSpinBox()
    fontSizeCB.setRange(6,48)
    #fontSizeCB.addItems([str(i) for i in range(6,24)])

    dtypeLabel = QLabel('dtype:')
    dtypeCB = QComboBox()
    dtypeCB.addItems(['float32','float64','int','complex','bytes','str'])

    rowLabel = QLabel('rows:')
    rowCount = QSpinBox()
    rowCount.setRange(30, 10000)
    rowCount.setReadOnly(True)
    rowCount.setFixedSize(100, 30)

    columnLabel = QLabel('columns:')
    columnCount = QSpinBox()
    columnCount.setRange(30,10000)
    columnCount.setReadOnly(True)
    columnCount.setFixedSize(100,30)

    bar1 = QToolBar()
    bar1.addAction(QIcon('pic/icon/save.png'),'Save file').triggered.connect(saveFile)
    bar1.addAction(QIcon('pic/icon/file.jpeg'),'Open file').triggered.connect(pick_sys_file)
    bar1.addAction(QIcon('pic/icon/newFile.png'),'New file')
    bar1.addSeparator()
    bar1.addAction(QIcon('pic/icon/exportPDF.png'),'Export pdf')
    bar1.addAction(QIcon('pic/icon/printer.png'),'Print')
    bar1.addSeparator()
    bar1.addAction(QIcon('pic/icon/cut_icon.png'),'Cut')
    bar1.addAction(QIcon(),'Copy')
    bar1.addAction(QIcon(),'Paste')
    bar1.addSeparator()
    bar1.addAction(QIcon('pic/icon/cellResize.png'),'resize cell to content').triggered.connect(resizeTableToContent)
    menuLayout_home.addWidget(bar1)

    bar2 = QToolBar()

    bar2.addWidget(fontTypeCB)
    bar2.addWidget(fontSizeCB)
    bar2.addWidget(dtypeLabel)
    bar2.addWidget(dtypeCB)
    bar2.addWidget(rowLabel)
    bar2.addWidget(rowCount)
    bar2.addWidget(columnLabel)
    bar2.addWidget(columnCount)
    menuLayout_home.addWidget(bar2)

    homeWidget = QWidget()
    homeWidget.setLayout(menuLayout_home)

    menuLayout.addWidget(homeWidget)

##########################################################################
    table_tab_box.addWidget(menuWidget)
    table_tab_box.addWidget(tableWidget)

########################### Menu bar ###############################################
    bar = QMenuBar()
    menuLayout.setMenuBar(bar)
    bar.setGeometry(0, 0, int(menuWidget.frameGeometry().width() / 1.6), int(screen_height / 10))

    #bar.setStyleSheet("background-color: white;")
    file = bar.addMenu("&File")
    file.addAction('&Open...').triggered.connect(pick_sys_file)
    file.addAction("Save As").triggered.connect(lambda : saveFile(saveAs=True))
    menuSave = file.addAction("Save")
    menuSave.triggered.connect(saveFile)
    menuSave.setShortcut("Ctrl+s")

    menuImport = file.addMenu('&Import')
    menuImport.addAction('numpy object(joblib)').triggered.connect(importJoblib)

    menuExport = file.addMenu('&Export')
    menuExport.addAction('numpy object(joblib)').triggered.connect(exportJoblib)
    menuExport.addAction('pandas object(joblib)').triggered.connect(lambda :exportJoblib('pd'))

    edit = bar.addMenu("Edit")
    edit.addAction("cu&t")#.setShortcut("Ctrl+X")
    edit.addAction("copy")#.setShortcut("Ctrl+C")
    edit.addAction("paste")#.setShortcut("Ctrl+V")
    edit.addSeparator()
    find  = edit.addMenu("Find...")
    find.addAction("Find")
    find.addAction("Replace")
    format = bar.addMenu('Format')
    encode = format.addMenu('Encoding...')
    encode.setObjectName('utf-8')
    utf_8 = encode.addAction('UTF-8')
    utf_8.triggered.connect(lambda :changeEncodeMethod([utf_8]))
    utf_8.setDisabled(True)
    utf_7 = encode.addAction('UTF_7')
    utf_7.triggered.connect(lambda :changeEncodeMethod([utf_7]))
    utf_16 = encode.addAction('UTF_16')
    utf_16.triggered.connect(lambda: changeEncodeMethod([utf_16]))
    utf_32 = encode.addAction('UTF_32')
    utf_32.triggered.connect(lambda: changeEncodeMethod([utf_32]))
    ascii = encode.addAction('Ascii')
    ascii.triggered.connect(lambda: changeEncodeMethod([ascii]))
    big5 = encode.addAction('big5')
    big5.triggered.connect(lambda :changeEncodeMethod([big5]))
    cp = encode.addMenu('CP')
    cp037 = cp.addAction('cp073')
    cp037.triggered.connect(lambda: changeEncodeMethod([cp,cp037]))

    view = bar.addMenu('&View')
    bar.addAction('Functions').triggered.connect(manageFunction)
    bar.addAction('Analyze').triggered.connect(analyze)

    def changeBarLayout(button):
        if button == 'home':
            homeAction.setDisabled(True)
            commandAction.setEnabled(True)
            commandWidget.hide()
            menuLayout.replaceWidget(commandWidget,homeWidget)
            homeWidget.show()
        elif button == 'command':
            commandAction.setDisabled(True)
            homeAction.setEnabled(True)
            homeWidget.hide()
            menuLayout.replaceWidget(homeWidget,commandWidget)
            commandWidget.show()

    homeAction = bar.addAction('Home')
    homeAction.triggered.connect(lambda :changeBarLayout('home'))
    homeAction.setDisabled(True)

    commandAction = bar.addAction('command')
    commandAction.triggered.connect(lambda :changeBarLayout('command'))

    return table_tab_box

if __name__ == '__main__':
    saved_file = True
    def closeEventHandler(event):
        if saved_file == True:
            event.accept()
        else:
            m = QMessageBox()
            m.setWindowTitle('file not save')
            ret = m.question(mainWidget,'', "Exit without saving?", m.Yes | m.No,m.No)
            
            #m.exec_()
            if ret == m.Yes:
                event.accept()
            else:
                event.ignore()

    app = QApplication(sys.argv)
    mainWidget = QWidget()
    mainWidget.setLayout(spreadsheet(1920,1080))
    mainWidget.show()
    mainWidget.closeEvent = closeEventHandler
    #app.aboutToQuit.connect(closeEventHandler)
    app.exec_()
    gc.collect()
    sys.exit()