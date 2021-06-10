"""    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import gc, sys, joblib
from os import close
from typing import Any

from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import *

import pandas as pd
import numpy as np
from string import ascii_uppercase
from webbrowser import open as webbrowser_open

import spreadsheet_command

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
            self.di=dict(zip([str((ord(c)%32)-1) for c in ascii_uppercase],ascii_uppercase))
            if '<U' in str(array.dtype) :
                self.numeric = False
            else:
                self.numeric = True

        def formatNumericHeader(self,section):
            if '10' in section:
                pass
            section = [i for i in section]
            a =''
            for i in section:
                a += self.di[i]
            return a

        def headerData(self, section: int, orientation: Qt.Orientation, role: int):
            if role == Qt.DisplayRole:
                if orientation == Qt.Horizontal:
                    if self.headers != None:
                        try:
                            return self.headers[section]  # column
                        except :
                            return self.formatNumericHeader(str(section))
                    else:
                        return self.formatNumericHeader(str(section)) # column
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
            global saved_file
            if role == Qt.EditRole:
                if value:
                    if value[0] == '=':
                        return True
                    saved_file = False
                    if value.isnumeric() and self.numeric:
                        if 'float' in str(self.array.dtype):
                            value = float(value)
                        elif 'int' in str(self.array.dtype):
                            value = int(value)
                    self.array[index.row()][index.column()] = value # asign new data to array
                    tableWidget.update()
                    return True
                else:
                    return False

        def flags(self, index):
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

#####################r read stuff #########################################

    def pick_sys_file(filter="All files (*)"):
        global current_file_name
        if saved_file == False:
            m = QMessageBox()
            m.setWindowTitle('file not save')
            ret = m.question(mainWidget,'', "open new file without saving?", m.Yes | m.No,m.No)
            
            if ret == m.No:
                return None

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
        elif '.pdobj' in file_name or '.npobj' in file_name:
            importJoblib(pick=False,filename=file_name,filter=filter)
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
        current_file_name = file_name
        return file_name

    
    def opencsv(file):
        global pandas_data
        data = pd.read_csv(file)
        pandas_data = data
        headers = list(data.keys())
        data = np.array(data)
        tableWidget.setModel(MyTableModel(data,headers=headers))

    def openexel(file):
        global pandas_data
        data = pd.read_excel(file,sheet_name=None)
        sheets = list(data.keys())

        num= [0]
        if len(sheets) >1:

            dialog = QDialog()
            dlayout = QVBoxLayout()
            dlayout.addWidget(QLabel(f'there is {len(sheets)} sheets\nonly one sheet can be imported'))
            box = QComboBox()
            box.addItems([f'sheet{i}: {sheets[i]}' for i in range(len(sheets))])
            def change(event,close=False):
                num[0] = event
                if close:
                    dialog.close()
            box.currentIndexChanged.connect(change)
            dlayout.addWidget(box)
            click = QPushButton('OK')
            click.clicked.connect(lambda:change(box.currentIndex(),True))
            dlayout.addWidget(click)
            dialog.setLayout(dlayout)
            dialog.setWindowTitle('select sheet')
            dialog.exec_()
            dialog.deleteLater()

        data = np.array(data[sheets[num[0]]])

        tableWidget.setModel(MyTableModel(data))

        headers = tableWidget.model().headers

        pandas_data = pd.DataFrame(tableWidget.model().array,columns=headers)
    
    def importJoblib(pick=True,filename=None,filter=None): # main importer to load binary file
        global pandas_data
        if pick:
            filter = 'Python Object(*.npobj *.pdobj)'
            filename, filter = QFileDialog.getOpenFileName(menuWidget, 'Open File', 'c://', filter=filter)
        if '.npobj' in filename:
            data = joblib.load(filename)
            header = None
        elif '.pdobj' in filename:
            dataframe = joblib.load(filename)
            header = list(dataframe.keys())
            data = np.array(dataframe)
        
        tableWidget.setModel(MyTableModel(data,headers=header))

        headers = tableWidget.model().headers

        pandas_data = pd.DataFrame(tableWidget.model().array,columns=headers)

#################### save stuff #########################
    def saveFile(directory=None,saveAs = False):
        global saved_file
        filter = "Pandas Object(*.pdobj);;Numpy Object(*.npobj);;CSV(*.csv);;JSON(*.json);;HTML(*.html);;Excel(*.xlsx);;PDF( *.pdf)"
        if saveAs:
            directory, filter = QFileDialog.getSaveFileName(menuWidget, 'Save File', 'c://',filter=filter)

        elif directory==None and current_file_name == None:
            directory, filter=QFileDialog.getSaveFileName(menuWidget,'Save File','c://',filter=filter)

        elif directory!=None:
            directory = str(directory)

            if len(directory.split('.'))>1:
                filter = '.'+directory.split('.')[-1]
            else:
                filter = '.pdobj'

        

        elif current_file_name !=None:
            directory= current_file_name
            filter = '.'+directory.split('.')[-1]

        headers = tableWidget.model().headers

        data = pd.DataFrame(tableWidget.model().array,columns=headers)
        print(filter)

        if '.pdobj' in filter:
            directory = directory+'.pdobj' if '.pdobj' not in directory else directory
            joblib.dump(data,directory,9)

        elif '.csv' in filter:
            directory = directory+'.csv' if '.csv' not in directory else directory
            data.to_csv(directory, index=False,header=False)

        elif '.json' in filter:
            directory = directory+'.json' if '.json' not in directory else directory
            out = data.to_json(directory,index=False,orient= 'table')

        elif '.html' in filter:
            directory = directory+'.html' if '.html' not in directory else directory
            html = data.to_html()
            with open(directory,'w') as f:
                f.write(html)
                f.close()

        elif'.xlsx'in directory:
            directory = directory+'.xlsx' if '.xlsx' not in directory else directory
            data.to_excel(directory,index=False)

        elif '.h5'in filter:
            data.to_hdf(directory,key='0')

        else :
            print('file type not supported yet')
        saved_file = True

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
    def alertbox(err):
        alert = QMessageBox()
        alert.setWindowTitle('ERROR')
        alert.setWindowIcon(QIcon('pic/icon/warning.png'))
        alert.setText(str(err))
        alert.setAttribute(Qt.WA_DeleteOnClose)
        alert.exec()
        alert.deleteLater()

    def changeEncodeMethod(a):
        for i in [utf_8,utf_7,utf_16,utf_32,ascii,big5,cp,cp037]:
            i.setDisabled(False)
        for i in a:
            i.setDisabled(True)
            encode.setObjectName(i.text())


    def spreadsheetCommand(interactive=False,scripting = False):
        global saved_file
        spreadsheet_command.main(commandBar,printOutLabel,tableWidget,scripting=scripting,interact=interactive,screen_width=screen_width,screen_height=screen_height)
        saved_file = False

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
        try:
            from matplotlib import pyplot as plt
        except Exception as e:  # it is mostly due to pillow
                from sys import executable
                from subprocess import check_call
                check_call([executable, "-m", "pip", "install", '-U', 'pillow'])
                i = QMessageBox()  # upgrade pillow may resolve the problem
                i.setText(e + '\nnow upgrading pillow')
                i.exec_()
                from matplotlib import pyplot as plt

        def setting(key,value):
            plt_setting[key] = value

        def plotHandler():
            try:
                x = xDataEdit.text()
                y = yDataEdit.text()
                for i in (x,y):
                    pass

            except Exception as e:
                print(e)
                alertbox(e)
                return

            plt.show()
            box.close()

        setting('set',True)

        scroll = QScrollArea()
        mainWidget = QWidget()

        scroll.setWidget(mainWidget)
        scroll.setWidgetResizable(True)

        plt.plot([4,6,2,8,99,22,73,68])
        
        box = QDialog()
        box.setMinimumSize(screen_width/2,screen_height/2)
        box.setWindowTitle('Analyze data')

        layout = QFormLayout(mainWidget)
        graphOption = QComboBox()
        graphOption.addItems(['matplotlib']) # pyqtgraph for data streaming, not piority, to do

        xDataEdit = QLineEdit()
        yDataEdit = QLineEdit()

        xLabelEdit = QLineEdit()
        xLabelEdit.textChanged.connect(lambda x: plt.xlabel(xLabelEdit.text()) | setting('xlabel',xLabelEdit.text()))
        yLabelEdit = QLineEdit()
        yLabelEdit.textChanged.connect(lambda x: plt.ylabel(yLabelEdit.text()) | setting('ylabel',yLabelEdit.text()))

        
        advenceBt = QPushButton('Advence')
        advenceBt.setFlat(True)
        
        advenceLayout = QFormLayout()
        advenceLayout.addRow(QLabel('hello'),QCheckBox())
        advenceWidget = QWidget()
        advenceWidget.setLayout(advenceLayout)
        advenceWidget.hide()# default as hide
        advenceBt.clicked.connect(lambda: advenceWidget.show()if advenceWidget.isHidden() else advenceWidget.hide()) # if hidden,show. if shown,hide

        plotBt = QPushButton('Plot')

        # plt.show() automatically creates a new window through qt
        plotBt.clicked.connect(plotHandler) 

        layout.setVerticalSpacing(20)
        layout.addRow(QLabel('Library: '),graphOption)
        layout.addRow(QLabel('xData: '),xDataEdit)
        layout.addRow(QLabel('yData: '),yDataEdit)
        layout.addItem(QSpacerItem(0,4))
        layout.addRow(QLabel('xlabel: '),xLabelEdit)
        layout.addRow(QLabel('ylabel: '),yLabelEdit)
        layout.addWidget(advenceBt)
        layout.addWidget(advenceWidget)
        layout.addWidget(plotBt)

        mainlayout = QVBoxLayout()
        mainlayout.addWidget(scroll)

        box.setLayout(mainlayout)

        if plt_setting['set']: # set all values user choosed
            if  'xlabel' in plt_setting:
                xLabelEdit.setText(plt_setting['xlabel'])
            if  'ylabel' in plt_setting:
                yLabelEdit.setText(plt_setting['ylabel'])

        box.setAttribute(Qt.WA_DeleteOnClose)
        box.exec_()
        box.deleteLater()



    def webview(site):
        import webview

        w = webview.WebView(width=int(screen_width/2), height=int(screen_height/2), title=site, url="https://www.desmos.com/calculator", resizable=True, debug=False)
        w.run()

    def managePip():
        pass


#################################################
    data = np.array([['']*60]*60,dtype='<U30')
    model = MyTableModel(data)
    table_tab_box = QVBoxLayout()
    tableWidget = QTableView()
    tableWidget.setModel(model)
    tableWidget.horizontalHeader().stretchLastSection()
    tableWidget.setStyleSheet('background-color:white;color:black;')

    spreadsheet_command.init(tableWidget)

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

    fontTypeLabel = QLabel('style:')
    fontTypeCB = QComboBox()
    
    fontTypeCB.addItems(QFontDatabase().families())

    fontSizeLabel = QLabel('fontsize:')
    fontSizeCB = QComboBox()
    fontSizeCB.addItems([str(i) for i in range(6,24)])

    dtypeLabel = QLabel('dtype:')
    dtypeLabel2 = QLabel('<U30')
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

    bar2.addWidget(fontTypeLabel)
    bar2.addWidget(fontTypeCB)
    bar2.addWidget(fontSizeLabel)
    bar2.addWidget(fontSizeCB)
    bar2.addWidget(dtypeLabel)
    bar2.addWidget(dtypeLabel2)
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

    menuOpen = file.addAction('&Open...')
    menuOpen.triggered.connect(pick_sys_file)
    menuOpen.setShortcut(QKeySequence("Ctrl+O"))
    
    menuSave = file.addAction("Save")
    menuSave.triggered.connect(saveFile)
    menuSave.setShortcut(QKeySequence("Ctrl+S"))

    menuSaveAs = file.addAction("Save As")
    menuSaveAs.triggered.connect(lambda : saveFile(saveAs=True))
    menuSaveAs.setShortcut(QKeySequence("Ctrl+Shift+S"))

    menuImport = file.addMenu('&Import')
    menuImport.addAction('numpy object(joblib)').triggered.connect(importJoblib)

    menuExport = file.addMenu('&Export')
    menuExport.addAction('numpy object(joblib)').triggered.connect(exportJoblib)
    menuExport.addAction('pandas object(joblib)').triggered.connect(lambda :exportJoblib('pd'))

    edit = bar.addMenu("Edit")
    editCut= edit.addAction("cu&t")
    editCut.setShortcut(QKeySequence("Ctrl+X"))
    #editCut.trigger.connect(cutCopyPasteHandler)

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

    viewStyle = view.addMenu('Style')
    viewStyle.addAction('Breeze')
    viewStyle.addAction('Oxygen')
    viewStyle.addAction('QtCurve')
    viewStyle.addAction('Fusion')
    viewStyle.addAction('Windows')

    viewTheme = view.addMenu('Theme')
    viewTheme.addAction('Dark').triggered.connect(lambda : mainWidget.setStyleSheet('background-color:#373737; color:white;'))
    viewTheme.addAction('Lite').triggered.connect(lambda : mainWidget.setStyleSheet('background-color:white; color:black;'))

    tools = bar.addMenu('Tool')

    # show website https://www.desmos.com/calculator using package webview, func at line 335
    tools.addAction('desmos').triggered.connect(lambda :webview('desmos'))

    # user interface to manage moduler functions, allow custom functions, func at line 284
    bar.addAction('Functions').triggered.connect(manageFunction)

    # analyze data using matplotlib, func at line 310
    bar.addAction('Analyze').triggered.connect(analyze)

    # manage pip packages, func at line 348
    bar.addAction('library')

    # call function from imported module, state interactive mode, func at line 251
    bar.addAction('console').triggered.connect(lambda :spreadsheetCommand(interactive=True))

    def changeBarLayout(button): # this switches between the home and command menu
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
    homeAction.triggered.connect(lambda :changeBarLayout('home')) # hardcode the selection
    homeAction.setDisabled(True)

    commandAction = bar.addAction('command')
    commandAction.triggered.connect(lambda :changeBarLayout('command'))

    menuhelp = bar.addAction('help').triggered.connect(lambda : webbrowser_open('https://github.com/YC-Lammy/np_spreadsheet/issues'))


    return table_tab_box # return the main layout

########################### execute ################################

if __name__ == '__main__':

    saved_file = True #state if the file is modified, notice user to save file
    current_file_name = None #current file name is the file user opened using open file function
    plt_setting = {'set':False}

    def closeEventHandler(event): # this function is called when user tries to close app, line 559

        if saved_file == True: # is nothing is modified, quit normally
            event.accept()
        else:
            m = QMessageBox()
            m.setWindowTitle('file not save')
            ret = m.question(mainWidget,'', "Exit without saving?", m.Yes | m.No,m.No) # default as No

            if ret == m.Yes:
                event.accept() # if user choose yes, exit without saving
            else:
                event.ignore() # when user choose no, stop exit event


    app = QApplication(['-style fusion']+sys.argv)

    mainWidget = QWidget()
    mainWidget.setLayout(spreadsheet(1920,1080)) # spreedsheet returns a layout
    mainWidget.show()
    mainWidget.closeEvent = closeEventHandler # reassign the app's close event
    mainWidget.setWindowState(Qt.WindowMaximized)
    mainWidget.setWindowTitle('spreadsheet') # actual title not desided

    app.exec_()
    gc.collect()
    sys.exit()