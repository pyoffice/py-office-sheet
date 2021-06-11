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
#################################################################
##      ___  ___       ___   _   __   _      ___     __    __  ##
##     /   |/   |     /   | | | |  \ | |    |    \  \  \  /  / ##
##    / /|   /| |    / /| | | | |   \| |    |     |  \  \/  /  ##
##   / / |__/ | |   / / | | | | | |\   |    |  __/    \    /   ##
##  / /       | |  / /  | | | | | | \  |    | |       /   /    ##
## /_/        |_| /_/   |_| |_| |_|  \_| ⬤ |_|      /__ /     ##
##                                                             ##
#################################################################

#                               _     _               _   
#                              | |   | |             | |  
#  ___ _ __  _ __ ___  __ _  __| |___| |__   ___  ___| |_ 
# / __| '_ \| '__/ _ \/ _` |/ _` / __| '_ \ / _ \/ _ \ __|
# \__ \ |_) | | |  __/ (_| | (_| \__ \ | | |  __/  __/ |_ 
# |___/ .__/|_|  \___|\__,_|\__,_|___/_| |_|\___|\___|\__|
#     | |                                                 
#     |_|                                                 


import gc, sys, joblib,os
import pyOfficeSheet
from os import close
from typing import Any

from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import *

import pandas as pd
import numpy as np
from string import ascii_uppercase
from webbrowser import open as webbrowser_open
from inspect import getfile
from json import loads as json_loads
from json import dumps as json_dumps

try :
    import spreadsheet_command # direct import if call locally
except:

    from pyOfficeSheet import spreadsheet_command


def spreadsheet(screen_width,screen_height):
    global saved_file

########################################################################################################################################################
############################## QAbstractTableModel #####################################################################################################
########################################################################################################################################################

    class MyTableModel(QAbstractTableModel): # numpy array model
        # when tableView is rendered, data of each cell would be called through self.data()
        # this method avoid creating widgets for each cell e.g. QtableWiget
        # by storing data in numpy array, fast processing of data could be performed
        # numpy opperation is also supported. 

        def __init__(self, array, headers= None,parent=None):
            super().__init__(parent)

            self.array = array # call current array later through tableWidget.model().array
            self.headers = headers

            self.stack = QUndoStack() # support undo redo function

            # create a dict {'A':'1', 'B':'2', 'C':'3' ...}
            self.di=dict(zip([str((ord(c)%32)-1) for c in ascii_uppercase],ascii_uppercase))

            if '<U' in str(array.dtype) : # '<U' is unicode
                self.numeric = False
            else:
                self.numeric = True

        def formatNumericHeader(self,section):

            section = [i for i in section]
            a =''
            for i in section:
                a += self.di[i]
            return a

        def headerData(self, section: int, orientation: Qt.Orientation, role: int): # fetch the header to GUI
            if role == Qt.DisplayRole:
                if orientation == Qt.Horizontal:
                    if self.headers != None:
                        try:
                            return self.headers[section]  # column headers maybe out of range
                        except :
                            return self.formatNumericHeader(str(section)) # return number instead if out of range
                    else:
                        return self.formatNumericHeader(str(section)) # column
                else:
                    return str(section)  # row

        def columnCount(self, parent=None):
            return len(self.array[0])

        def rowCount(self, parent=None):
            return len(self.array)

        def data(self, index: QModelIndex, role: int): # fetch data to GUI

            if role == Qt.DisplayRole or role == Qt.EditRole:
                row = index.row()
                col = index.column()
                return str(self.array[row][col]) # return value of cell

        def setData(self, index, value, role): # set data would be called everytime user edit cell
            global saved_file # if user modify the array, the array is modified

            if role == Qt.EditRole: # check if func called correctly
                if value:

                    self.stack.push(CellEdit(index, value, self)) # push a new command

                    if value[0] == '=': # '=' indicates to perform a function
                        pass

                    saved_file = False # indicate file modifyed

                    if value.isnumeric() and self.numeric:
                        if 'float' in str(self.array.dtype):
                            value = float(value)
                        elif 'int' in str(self.array.dtype):
                            value = int(value)

                    self.array[index.row()][index.column()] = value # asign new data to array
                    tableWidget.update() # update GUI
                    return True
                else:
                    return False # vlue not provided mal function call

        def undo(self):
            self.stack.undo()

        def redo(self):
            self.stack.redo()

        def flags(self, index): # indicate the model's flags
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable 

    # support for undo redo command
    class CellEdit(QUndoCommand): # a new command is pushed to the model stack every time cell is edit

        def __init__(self, index, value, model, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.index = index # save the cell location
            self.value = value # save the new value 
            self.prev = model.array[index.row()][index.column()] # save the previous value
            self.model = model # a pointer to the model

        def undo(self):
            # set the specific cell to the previous value
            self.model.array[self.index.row()][self.index.column()] = self.prev
            tableWidget.update()

        def redo(self):
            # set the specific cell to the new value
            self.model.array[self.index.row()][self.index.column()] = self.value
            tableWidget.update()

############################################################################################################################################################
############################## read stuff ##################################################################################################################
############################################################################################################################################################
#                                                                    dddddddd                                                                                                        
#                                                                    d::::::d                               tttt                               ffffffffffffffff    ffffffffffffffff  
#                                                                    d::::::d                            ttt:::t                              f::::::::::::::::f  f::::::::::::::::f 
#                                                                    d::::::d                            t:::::t                             f::::::::::::::::::ff::::::::::::::::::f
#                                                                    d:::::d                             t:::::t                             f::::::fffffff:::::ff::::::fffffff:::::f
# rrrrr   rrrrrrrrr       eeeeeeeeeeee    aaaaaaaaaaaaa      ddddddddd:::::d          ssssssssss   ttttttt:::::ttttttt    uuuuuu    uuuuuu   f:::::f       fffffff:::::f       ffffff
# r::::rrr:::::::::r    ee::::::::::::ee  a::::::::::::a   dd::::::::::::::d        ss::::::::::s  t:::::::::::::::::t    u::::u    u::::u   f:::::f             f:::::f             
# r:::::::::::::::::r  e::::::eeeee:::::eeaaaaaaaaa:::::a d::::::::::::::::d      ss:::::::::::::s t:::::::::::::::::t    u::::u    u::::u  f:::::::ffffff      f:::::::ffffff       
# rr::::::rrrrr::::::re::::::e     e:::::e         a::::ad:::::::ddddd:::::d      s::::::ssss:::::stttttt:::::::tttttt    u::::u    u::::u  f::::::::::::f      f::::::::::::f       
#  r:::::r     r:::::re:::::::eeeee::::::e  aaaaaaa:::::ad::::::d    d:::::d       s:::::s  ssssss       t:::::t          u::::u    u::::u  f::::::::::::f      f::::::::::::f       
#  r:::::r     rrrrrrre:::::::::::::::::e aa::::::::::::ad:::::d     d:::::d         s::::::s            t:::::t          u::::u    u::::u  f:::::::ffffff      f:::::::ffffff       
#  r:::::r            e::::::eeeeeeeeeee a::::aaaa::::::ad:::::d     d:::::d            s::::::s         t:::::t          u::::u    u::::u   f:::::f             f:::::f             
#  r:::::r            e:::::::e         a::::a    a:::::ad:::::d     d:::::d      ssssss   s:::::s       t:::::t    ttttttu:::::uuuu:::::u   f:::::f             f:::::f             
#  r:::::r            e::::::::e        a::::a    a:::::ad::::::ddddd::::::dd     s:::::ssss::::::s      t::::::tttt:::::tu:::::::::::::::uuf:::::::f           f:::::::f            
#  r:::::r             e::::::::eeeeeeeea:::::aaaa::::::a d:::::::::::::::::d     s::::::::::::::s       tt::::::::::::::t u:::::::::::::::uf:::::::f           f:::::::f            
#  r:::::r              ee:::::::::::::e a::::::::::aa:::a d:::::::::ddd::::d      s:::::::::::ss          tt:::::::::::tt  uu::::::::uu:::uf:::::::f           f:::::::f            
#  rrrrrrr                eeeeeeeeeeeeee  aaaaaaaaaa  aaaa  ddddddddd   ddddd       sssssssssss              ttttttttttt      uuuuuuuu  uuuufffffffff           fffffffff            
                                                                                                                                                                           

    def pick_sys_file(filter="All files (*)"): # this func is called whenever user open or import a file
        global current_file_name, saved_file

        if saved_file == False: # if file not save, opening new file will remove the data

            m = QMessageBox()
            m.setWindowTitle('file not save')
            ret = m.question(mainWidget,'', "open new file without saving?", m.Yes | m.No,m.No)
            
            if ret == m.No:
                return False # stop the func 

        from mimetypes import guess_type # this package is not necessary, can be remove in future release

        if filter == False:
            filter = "All files (*)" # user can choose any file

        file_name, filter = QFileDialog.getOpenFileName(menuWidget, 'Open File', filter=filter)

        type = guess_type(file_name) 

        print(type)

        if 'text/csv' in type:

            opencsv(file_name) # open as csv

        elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in type:

            openexel(file_name) # open as excel

        elif '.pdobj' in file_name or '.npobj' in file_name: # guess_type does not support binary formats

            importJoblib(pick=False,filename=file_name,filter=filter)

        else:
            # hyperlink to github discussions
            alertbox('<p>file type not supported yet\r\n</p><p>request feature on:\n</p><a href="https://github.com/YC-Lammy/np_spreadsheet/discussions">github.com/YC-Lammy/np_spreadsheet/discussions</a>',)
            return None

        updateInfo()

        current_file_name = file_name
        saved_file = True
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
            filename, filter = QFileDialog.getOpenFileName(menuWidget, 'Open File', filter=filter)
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

############################################################################################################################################################
############################# save stuff ###################################################################################################################
############################################################################################################################################################
#                                                                                                                                                                                   
#                                                                                                          tttt                               ffffffffffffffff    ffffffffffffffff  
#                                                                                                       ttt:::t                              f::::::::::::::::f  f::::::::::::::::f 
#                                                                                                       t:::::t                             f::::::::::::::::::ff::::::::::::::::::f
#                                                                                                       t:::::t                             f::::::fffffff:::::ff::::::fffffff:::::f
#     ssssssssss     aaaaaaaaaaaaavvvvvvv           vvvvvvv eeeeeeeeeeee             ssssssssss   ttttttt:::::ttttttt    uuuuuu    uuuuuu   f:::::f       fffffff:::::f       ffffff
#   ss::::::::::s    a::::::::::::av:::::v         v:::::vee::::::::::::ee         ss::::::::::s  t:::::::::::::::::t    u::::u    u::::u   f:::::f             f:::::f             
# ss:::::::::::::s   aaaaaaaaa:::::av:::::v       v:::::ve::::::eeeee:::::ee     ss:::::::::::::s t:::::::::::::::::t    u::::u    u::::u  f:::::::ffffff      f:::::::ffffff       
# s::::::ssss:::::s           a::::a v:::::v     v:::::ve::::::e     e:::::e     s::::::ssss:::::stttttt:::::::tttttt    u::::u    u::::u  f::::::::::::f      f::::::::::::f       
#  s:::::s  ssssss     aaaaaaa:::::a  v:::::v   v:::::v e:::::::eeeee::::::e      s:::::s  ssssss       t:::::t          u::::u    u::::u  f::::::::::::f      f::::::::::::f       
#    s::::::s        aa::::::::::::a   v:::::v v:::::v  e:::::::::::::::::e         s::::::s            t:::::t          u::::u    u::::u  f:::::::ffffff      f:::::::ffffff       
#       s::::::s    a::::aaaa::::::a    v:::::v:::::v   e::::::eeeeeeeeeee             s::::::s         t:::::t          u::::u    u::::u   f:::::f             f:::::f             
# ssssss   s:::::s a::::a    a:::::a     v:::::::::v    e:::::::e                ssssss   s:::::s       t:::::t    ttttttu:::::uuuu:::::u   f:::::f             f:::::f             
# s:::::ssss::::::sa::::a    a:::::a      v:::::::v     e::::::::e               s:::::ssss::::::s      t::::::tttt:::::tu:::::::::::::::uuf:::::::f           f:::::::f            
# s::::::::::::::s a:::::aaaa::::::a       v:::::v       e::::::::eeeeeeee       s::::::::::::::s       tt::::::::::::::t u:::::::::::::::uf:::::::f           f:::::::f            
#  s:::::::::::ss   a::::::::::aa:::a       v:::v         ee:::::::::::::e        s:::::::::::ss          tt:::::::::::tt  uu::::::::uu:::uf:::::::f           f:::::::f            
#   sssssssssss      aaaaaaaaaa  aaaa        vvv            eeeeeeeeeeeeee         sssssssssss              ttttttttttt      uuuuuuuu  uuuufffffffff           fffffffff            
                                                                                                                                                                       

    def saveFile(directory=None,saveAs = False):
        global saved_file
        filter = "Pandas Object(*.pdobj);;Numpy Object(*.npobj);;CSV(*.csv);;JSON(*.json);;HTML(*.html);;Excel(*.xlsx);;PDF( *.pdf)"
        if saveAs:
            directory, filter = QFileDialog.getSaveFileName(menuWidget, 'Save File',filter=filter)

        elif directory==None and current_file_name == None:
            directory, filter=QFileDialog.getSaveFileName(menuWidget,'Save File',filter=filter)

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
            alertbox('file type not supported yet\nrequest feature on:\nhttps://github.com/YC-Lammy/np_spreadsheet/issues')
        saved_file = True

    def exportJoblib(dtype='np'): # export numpy array by joblib

        filename, filter = QFileDialog.getSaveFileName(menuWidget, 'Save File')

        if dtype=='np':
            filename += '.npobj'
            array = tableWidget.model().array
            joblib.dump(array,filename,9)

        elif dtype == 'pd':
            filename +='.pdobj'
            headers = tableWidget.model().headers
            data = pd.DataFrame(tableWidget.model().array,columns=headers)
            joblib.dump(data,filename,9)

        
############################################################################################################################################################
################################# operational functions ##########################################################################################################
############################################################################################################################################################
#                                                                                                                                                                
#     ffffffffffffffff                                                                 tttt            iiii                                                      
#    f::::::::::::::::f                                                             ttt:::t           i::::i                                                     
#   f::::::::::::::::::f                                                            t:::::t            iiii                                                      
#   f::::::fffffff:::::f                                                            t:::::t                                                                      
#   f:::::f       ffffffuuuuuu    uuuuuunnnn  nnnnnnnn        ccccccccccccccccttttttt:::::ttttttt    iiiiiii    ooooooooooo   nnnn  nnnnnnnn        ssssssssss   
#   f:::::f             u::::u    u::::un:::nn::::::::nn    cc:::::::::::::::ct:::::::::::::::::t    i:::::i  oo:::::::::::oo n:::nn::::::::nn    ss::::::::::s  
#  f:::::::ffffff       u::::u    u::::un::::::::::::::nn  c:::::::::::::::::ct:::::::::::::::::t     i::::i o:::::::::::::::on::::::::::::::nn ss:::::::::::::s 
#  f::::::::::::f       u::::u    u::::unn:::::::::::::::nc:::::::cccccc:::::ctttttt:::::::tttttt     i::::i o:::::ooooo:::::onn:::::::::::::::ns::::::ssss:::::s
#  f::::::::::::f       u::::u    u::::u  n:::::nnnn:::::nc::::::c     ccccccc      t:::::t           i::::i o::::o     o::::o  n:::::nnnn:::::n s:::::s  ssssss 
#  f:::::::ffffff       u::::u    u::::u  n::::n    n::::nc:::::c                   t:::::t           i::::i o::::o     o::::o  n::::n    n::::n   s::::::s      
#   f:::::f             u::::u    u::::u  n::::n    n::::nc:::::c                   t:::::t           i::::i o::::o     o::::o  n::::n    n::::n      s::::::s   
#   f:::::f             u:::::uuuu:::::u  n::::n    n::::nc::::::c     ccccccc      t:::::t    tttttt i::::i o::::o     o::::o  n::::n    n::::nssssss   s:::::s 
#  f:::::::f            u:::::::::::::::uun::::n    n::::nc:::::::cccccc:::::c      t::::::tttt:::::ti::::::io:::::ooooo:::::o  n::::n    n::::ns:::::ssss::::::s
#  f:::::::f             u:::::::::::::::un::::n    n::::n c:::::::::::::::::c      tt::::::::::::::ti::::::io:::::::::::::::o  n::::n    n::::ns::::::::::::::s 
#  f:::::::f              uu::::::::uu:::un::::n    n::::n  cc:::::::::::::::c        tt:::::::::::tti::::::i oo:::::::::::oo   n::::n    n::::n s:::::::::::ss  
#  fffffffff                uuuuuuuu  uuuunnnnnn    nnnnnn    cccccccccccccccc          ttttttttttt  iiiiiiii   ooooooooooo     nnnnnn    nnnnnn  sssssssssss    
                                                                                                                                                               
    pic_file_path = os.path.join(getfile(pyOfficeSheet).replace('__init__.py',''),'pic','icon')

    def alertbox(err,arg=False): # a function to alert user when error occourse
        alert = QMessageBox()
        alert.setWindowTitle('ERROR')
        alert.setWindowIcon(QIcon(os.path.join(pic_file_path,'warning.png')))
        alert.setText(str(err))
        alert.setAttribute(Qt.WA_DeleteOnClose) # prevent memory leak
        alert.setTextInteractionFlags(Qt.TextBrowserInteraction)
        alert.exec()
        if arg:
            exec(arg)
        alert.deleteLater()

    def changeEncodeMethod(a):
        for i in [utf_8,utf_7,utf_16,utf_32,ascii,big5,cp,cp037]:
            i.setDisabled(False)
        for i in a:
            i.setDisabled(True)
            encode.setObjectName(i.text())

    def changeSettings(key,value):
        settings[key] = value

    def updateInfo(): # update GUI info
        column = tableWidget.model().columnCount()
        columnCount.setRange(column, column+10000)
        columnCount.setValue(column)
        row = tableWidget.model().rowCount()
        rowCount.setRange(row, row +10000)
        rowCount.setValue(row)
        dtypeLabel2.setText(tableWidget.model().array.dtype.name)

    def spreadsheetCommand(interactive=False,scripting = False): # call function from spreadsheet command.py
        global saved_file
        spreadsheet_command.main(commandBar,printOutLabel,tableWidget,scripting=scripting,interact=interactive,screen_width=screen_width,screen_height=screen_height)
        saved_file = False
        updateInfo()

    def commandHandler(event): # handles the keyboard input of the command QLineEdit

        text = event.text()

        if event.key() == 16777235:
            commandBar.clear()
            commandBar.insert('lastcommand')
            spreadsheetCommand()
            return

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

    ################ manage Functions #######################

    def manageFunction():
        manage_function_box = QDialog()
        manage_function_box.setWindowModality(Qt.WindowModal)
        manage_function_box.setWindowTitle('Manage Functions')
        manage_function_box.setWindowIcon(QIcon(os.path.join(pic_file_path,'pic/icon/main.png')))
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

    ##################### analyze ############################

    def analyze(): # a panal to select matplotlib options and call matplotlib
        try:
            from matplotlib import pyplot as plt

        except Exception as e:  # it is mostly due to pillow, this may happen on linux
                from sys import executable
                from subprocess import check_call

                # upgrade pillow may resolve the problem
                err = check_call([executable, "-m", "pip", "install", '-U', 'pillow'])

                alertbox(e + '\r\nnow upgrading pillow')

                from matplotlib import pyplot as plt # reimport after upgrade

        def setting(key,value):
            plt_setting[key] = value # sets the dictionary

        def plotHandler(): # call when user click the "plot" button
            try:
                x = xDataEdit.text()
                y = yDataEdit.text()

                plt.show()
                box.close()

            except Exception as e:
                print(e)
                alertbox(e)

           

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
        xDataEdit.textChanged.connect(lambda x: setting('xdata',xDataEdit.text()))
        yDataEdit = QLineEdit()
        yDataEdit.textChanged.connect(lambda x: setting('ydata',yDataEdit.text()))

        titleEdit = QLineEdit()
        titleEdit.textChanged.connect(lambda x: plt.title(titleEdit.text()) | setting('title',titleEdit.text()))

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
        layout.addRow(QLabel('title'), titleEdit)
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
            if  'xdata' in plt_setting:
                xDataEdit.setText(plt_setting['xdata'])
            if  'ydata' in plt_setting:
                yDataEdit.setText(plt_setting['ydata'])
            if  'title' in plt_setting:
                titleEdit.setText(plt_setting['title'])

        box.setWindowModality(Qt.WindowModal)
        box.setAttribute(Qt.WA_DeleteOnClose)
        box.exec_()


    ######################## webview ##################################
    def webview(site):
        import webview

        if site == 'desmos':
            w = webview.WebView(width=int(screen_width/2), height=int(screen_height/2), title=site, url="https://www.desmos.com/calculator", resizable=True, debug=False)
        elif site == 'desmos calc':
            w = webview.WebView(width=int(screen_width/2), height=int(screen_height/2), title=site, url="https://www.desmos.com/scientific", resizable=True, debug=False)
        w.run()

    ######################## manage Pip #####################################
    def managePip():
        pass

##############################################################################################################################################
############################ set up ##########################################################################################################
##############################################################################################################################################
#                                               tttt                                               
#                                            ttt:::t                                               
#                                            t:::::t                                               
#                                            t:::::t                                               
#     ssssssssss       eeeeeeeeeeee    ttttttt:::::ttttttt    uuuuuu    uuuuuu ppppp   ppppppppp   
#   ss::::::::::s    ee::::::::::::ee  t:::::::::::::::::t    u::::u    u::::u p::::ppp:::::::::p  
# ss:::::::::::::s  e::::::eeeee:::::eet:::::::::::::::::t    u::::u    u::::u p:::::::::::::::::p 
# s::::::ssss:::::se::::::e     e:::::etttttt:::::::tttttt    u::::u    u::::u pp::::::ppppp::::::p
#  s:::::s  ssssss e:::::::eeeee::::::e      t:::::t          u::::u    u::::u  p:::::p     p:::::p
#    s::::::s      e:::::::::::::::::e       t:::::t          u::::u    u::::u  p:::::p     p:::::p
#       s::::::s   e::::::eeeeeeeeeee        t:::::t          u::::u    u::::u  p:::::p     p:::::p
# ssssss   s:::::s e:::::::e                 t:::::t    ttttttu:::::uuuu:::::u  p:::::p    p::::::p
# s:::::ssss::::::se::::::::e                t::::::tttt:::::tu:::::::::::::::uup:::::ppppp:::::::p
# s::::::::::::::s  e::::::::eeeeeeee        tt::::::::::::::t u:::::::::::::::up::::::::::::::::p 
#  s:::::::::::ss    ee:::::::::::::e          tt:::::::::::tt  uu::::::::uu:::up::::::::::::::pp  
#   sssssssssss        eeeeeeeeeeeeee            ttttttttttt      uuuuuuuu  uuuup::::::pppppppp    
#                                                                               p:::::p            
#                                                                               p:::::p            
#                                                                              p:::::::p           
#                                                                              p:::::::p           
#                                                                              p:::::::p           
#                                                                              ppppppppp  


    data = np.array([['']*60]*60,dtype='object')# empty array

    table_tab_box = QVBoxLayout()
    tableWidget = QTableView()
    tableWidget.setModel(MyTableModel(data))
    tableWidget.horizontalHeader().stretchLastSection()
    tableWidget.setStyleSheet('background-color:white;color:black;')

    spreadsheet_command.init(tableWidget)

############################ Menu section #########################################################################################################

    menuWidget = QWidget()
    menuLayout = QVBoxLayout()
    menuWidget.setLayout(menuLayout)
    menuWidget.setMinimumHeight(screen_height /8)

################## command layout #################################################################################################################

    menuLayout_command = QFormLayout()
    commandBar = QLineEdit()
    commandBar.keyReleaseEvent = commandHandler
    commandBar.setMaximumWidth(int(screen_width/1.05))
    commandBar.returnPressed.connect(spreadsheetCommand)
    functionCompleter = QCompleter(
        ['replaceColumn()'])
    functionCompleter.setCaseSensitivity(Qt.CaseInsensitive)
    commandBar.setCompleter(functionCompleter)
    printOutLabel = QLabel()
    menuLayout_command.addRow(QLabel('>>>'),printOutLabel)
    menuLayout_command.addWidget(commandBar)

    commandWidget = QWidget()
    commandWidget.setLayout(menuLayout_command)

################### home layout ###################################################################################################################

    menuLayout_home = QVBoxLayout()
    menuLayout_home.setAlignment(Qt.AlignTop)

    def resizeTableToContent():
        tableWidget.resizeRowsToContents()
        tableWidget.resizeColumnsToContents()

    fontTypeCB = QComboBox()
    
    fontTypeCB.addItems(QFontDatabase().families())
    fontTypeCB.setFixedWidth(int(screen_height/8))
    fontTypeCB.setCurrentText('Noto Sans New Tai Lue')

    def setFont():
        tableWidget.setFont(QFont(fontTypeCB.currentText()))
        tableWidget.update()
        changeSettings('font',fontTypeCB.currentText())

    fontTypeCB.currentIndexChanged.connect(setFont)
    tableWidget.setFont(QFont(fontTypeCB.currentText()))

    fontSizeCB = QComboBox()
    fontSizeCB.addItems([str(i) for i in range(6,24)])
    fontSizeCB.setCurrentText(str(tableWidget.font().pointSize()))

    def setFontSize():
        font = tableWidget.font()
        font.setPointSize(int(fontSizeCB.currentText()))
        tableWidget.setFont(font)
        tableWidget.update()
        changeSettings('fontsize',fontSizeCB.currentText())

    fontSizeCB.currentIndexChanged.connect(setFontSize)

    dtypeLabel = QLabel('dtype:')
    dtypeLabel2 = QLabel('object')
    #dtypeCB = QComboBox()
    #dtypeCB.addItems(['float32','float64','int','complex','bytes','str'])

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
    bar1.setIconSize(QSize(int(screen_height/30),int(screen_height/30)))
    bar1.addAction(QIcon(os.path.join(pic_file_path,'save.png')),'Save file').triggered.connect(saveFile)
    bar1.addAction(QIcon(os.path.join(pic_file_path,'file.png')),'Open file').triggered.connect(pick_sys_file)
    bar1.addAction(QIcon(os.path.join(pic_file_path,'newFile.png')),'New file')
    bar1.addSeparator()
    bar1.addAction(QIcon(os.path.join(pic_file_path,'exportCSV.png')),'Export csv')
    bar1.addAction(QIcon(os.path.join(pic_file_path,'printer.png')),'Print')
    bar1.addSeparator()
    bar1.addAction(QIcon(os.path.join(pic_file_path,'cut_icon.png')),'Cut')
    bar1.addAction(QIcon(),'Copy')
    bar1.addAction(QIcon(),'Paste')
    bar1.addSeparator()
    bar1.addAction(QIcon(os.path.join(pic_file_path,'cellResize.png')),'resize cell to content').triggered.connect(resizeTableToContent)
    menuLayout_home.addWidget(bar1)

    bar2 = QToolBar()
    bar2.setIconSize(QSize(int(screen_height/30),int(screen_height/30)))
    bar2.addWidget(fontTypeCB)
    bar2.addWidget(fontSizeCB)
    bar2.addWidget(dtypeLabel)
    bar2.addWidget(dtypeLabel2)
    bar2.addSeparator()
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

    editUndo = edit.addAction('Undo')
    editUndo.setShortcut(QKeySequence("Ctrl+Z"))
    editUndo.triggered.connect(lambda:tableWidget.model().undo())

    editRedo = edit.addAction('Redo')
    editRedo.setShortcut(QKeySequence("Ctrl+Shift+Z"))
    editRedo.triggered.connect(lambda:tableWidget.model().redo())

    editCut= edit.addAction("cu&t")
    editCut.setShortcut(QKeySequence("Ctrl+X"))
    #editCut.triggered.connect(cutCopyPasteHandler)

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

    def changeTheme(theme):
        if theme == 'dark':
            mainWidget.setStyleSheet('background-color:#2f2f35; color:white;')
        elif theme == 'lite':
            mainWidget.setStyleSheet('background-color:white; color:black;')
        elif theme == 'system':
            mainWidget.setStyleSheet('')
        changeSettings('theme',theme)

    viewTheme = view.addMenu('Theme')
    viewTheme.addAction('Dark').triggered.connect(lambda : changeTheme('dark'))
    viewTheme.addAction('Lite').triggered.connect(lambda : changeTheme('lite'))
    viewTheme.addAction('System').triggered.connect(lambda: changeTheme('system'))


    


    # manage pip packages, func at line 348
    menuLibrary = bar.addMenu('library')

    # user interface to manage moduler functions, allow custom functions, func at line 284
    menuLibrary.addAction('Functions').triggered.connect(manageFunction)
    menuLibrary.addAction('Pip packages').triggered.connect(managePip)

    # call function from imported module, state interactive mode, func at line 251
    bar.addAction('console').triggered.connect(lambda :spreadsheetCommand(interactive=True))

    # analyze data using matplotlib, func at line 310
    bar.addAction('Analyze').triggered.connect(analyze)

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


    tools = bar.addMenu('Tool')

    # show website https://www.desmos.com/calculator using package webview, func at line 335
    tools.addAction('desmos').triggered.connect(lambda :webview('desmos'))
    tools.addAction('desmos calc').triggered.connect(lambda: webview('desmos calc'))

    homeAction = bar.addAction('Home')
    homeAction.triggered.connect(lambda :changeBarLayout('home')) # hardcode the selection
    homeAction.setDisabled(True)

    commandAction = bar.addAction('command')
    commandAction.triggered.connect(lambda :changeBarLayout('command'))

    menuhelp = bar.addAction('help').triggered.connect(lambda : webbrowser_open('https://github.com/YC-Lammy/np_spreadsheet/issues'))

    # font, fontsize, theme
    # get the module path and read config file
    jsonpath = os.path.join(getfile(pyOfficeSheet).replace('__init__.py',''),'config.json')

    with open(jsonpath,'r') as f:
        json = f.read()
    json = json_loads(json) # convert json to dict

    if 'font' in json:
        fontTypeCB.setCurrentText(json['font'])
    if 'fontsize' in json:
        fontSizeCB.setCurrentText(json['fontsize'])
    if 'theme' in json:
        changeTheme(json['theme'])


    return table_tab_box # return the main layout


# EEEEEEEEEEEEEEEEEEEEEE                                                                                  tttt                              
# E::::::::::::::::::::E                                                                               ttt:::t                              
# E::::::::::::::::::::E                                                                               t:::::t                              
# EE::::::EEEEEEEEE::::E                                                                               t:::::t                              
#   E:::::E       EEEEEExxxxxxx      xxxxxxx eeeeeeeeeeee        ccccccccccccccccuuuuuu    uuuuuuttttttt:::::ttttttt        eeeeeeeeeeee    
#   E:::::E              x:::::x    x:::::xee::::::::::::ee    cc:::::::::::::::cu::::u    u::::ut:::::::::::::::::t      ee::::::::::::ee  
#   E::::::EEEEEEEEEE     x:::::x  x:::::xe::::::eeeee:::::ee c:::::::::::::::::cu::::u    u::::ut:::::::::::::::::t     e::::::eeeee:::::ee
#   E:::::::::::::::E      x:::::xx:::::xe::::::e     e:::::ec:::::::cccccc:::::cu::::u    u::::utttttt:::::::tttttt    e::::::e     e:::::e
#   E:::::::::::::::E       x::::::::::x e:::::::eeeee::::::ec::::::c     cccccccu::::u    u::::u      t:::::t          e:::::::eeeee::::::e
#   E::::::EEEEEEEEEE        x::::::::x  e:::::::::::::::::e c:::::c             u::::u    u::::u      t:::::t          e:::::::::::::::::e 
#   E:::::E                  x::::::::x  e::::::eeeeeeeeeee  c:::::c             u::::u    u::::u      t:::::t          e::::::eeeeeeeeeee  
#   E:::::E       EEEEEE    x::::::::::x e:::::::e           c::::::c     cccccccu:::::uuuu:::::u      t:::::t    tttttte:::::::e           
# EE::::::EEEEEEEE:::::E   x:::::xx:::::xe::::::::e          c:::::::cccccc:::::cu:::::::::::::::uu    t::::::tttt:::::te::::::::e          
# E::::::::::::::::::::E  x:::::x  x:::::xe::::::::eeeeeeee   c:::::::::::::::::c u:::::::::::::::u    tt::::::::::::::t e::::::::eeeeeeee  
# E::::::::::::::::::::E x:::::x    x:::::xee:::::::::::::e    cc:::::::::::::::c  uu::::::::uu:::u      tt:::::::::::tt  ee:::::::::::::e  
# EEEEEEEEEEEEEEEEEEEEEExxxxxxx      xxxxxxx eeeeeeeeeeeeee      cccccccccccccccc    uuuuuuuu  uuuu        ttttttttttt      eeeeeeeeeeeeee  
                                                                                                                                          

def main():
    global plt_setting, saved_file, current_file_name, settings, mainWidget

    saved_file = True #state if the file is modified, notice user to save file
    current_file_name = None #current file name is the file user opened using open file function
    plt_setting = {'set':False}
    settings = {}


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
    jsonpath = os.path.join(getfile(pyOfficeSheet).replace('__init__.py',''),'config.json')

    with open(jsonpath,'w') as f:
        f.write(json_dumps(settings))
        f.close()
    
    gc.collect()
    sys.exit()

if __name__ == '__main__':
    main()