""" Console
Interactive console widget.  Use to add an interactive python interpreter
in a GUI application.
"""

from PySide2.QtWidgets import *
from PySide2.QtCore import Qt, QEvent
import PySide2.QtGui as QtGui



def console(screen_width, screen_height):
    def execute(line):
        text = lineEdit['now'].text()
        label = QLabel('>>')
        label.setStyleSheet('color:white;')
        label1 = QLabel(text)
        label1.setStyleSheet('color:white;')
        formlayout.insertRow(formlayout.rowCount()-1,label,label1)
        lineEdit['now'].clear()
        spacing = ''

        def print(text, sep='',end='\n'):
            label = QLabel(str(text)+end)
            label.setStyleSheet('color:white;')
            label.setMaximumHeight(15)
            formlayout.insertRow(formlayout.rowCount()-1,QLabel(),label)

        if text[-1:] == ':':
            spacing = '    '
            mainlabel.setText('..')
            lineEdit['space'] = True
            lineEdit['buffer'] = text +'\n'

        elif lineEdit['space']== True:
            if text.isspace():
                mainlabel.setText('>>')
                label.setText('/')
                lineEdit['space'] = False
                text = lineEdit['buffer']
                try:
                    loca = loc
                    text = compile('locals().update(loca)\n' + text,'user','exec')
                    exec(text, locals(), loc)
                except Exception as e:
                    print(e)
            else :
                spacing='    '
                mainlabel.setText('..')
                lineEdit['buffer'] += text + '\n'

        else:
            try:
                loca = loc
                text = compile('locals().update(loca)\n' + text, 'user', 'exec')
                exec(text, locals(), loc)
            except Exception as e:
                print(e)
        lineEdit['now'].insert(spacing)

    loc = {}
    scroll = QScrollArea()
    scroll.verticalScrollBar().rangeChanged\
        .connect(lambda :scroll.verticalScrollBar().setValue(scroll.verticalScrollBar().maximum()))

    mainwidget = QWidget()

    scroll.setWidget(mainwidget)
    mainwidget.setStyleSheet('background-color:black;')

    layout = QVBoxLayout()
    layout.addWidget(scroll)
    scroll.setWidgetResizable(True)

    formlayout = QFormLayout(mainwidget)
    formlayout.setMargin(0)
    formlayout.setSpacing(0)
    formlayout.setVerticalSpacing(0)

    from sys import version,platform
    label = QLabel("Python "+str(version)+' on '+platform)
    label.setStyleSheet('color:white;')
    formlayout.addRow(QLabel(),label)

    line = QLineEdit()
    lineEdit = {'now':line,'space':False}
    line.setFrame(False)
    line.returnPressed.connect(lambda :execute(0))
    line.setStyleSheet('background-color:black; color:white;')
    #line.setFont(QtGui.QFont('Lucida Sans Typewriter', 10))

    mainlabel = QLabel('>>')
    mainlabel.setStyleSheet('color:white;')
    formlayout.addRow(mainlabel,line)

    return layout