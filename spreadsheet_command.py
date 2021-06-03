# Do NOT Edit This File If You Do NOT Know What You Doing :)

# this is an extension for editing table data
# table data are stored in an numpy ndarray
# the np array can be access by calling tableWidget.model().array
# to modify the data simply do array[row][column] = new_data
# if modify the shape of array, tableWidget.setModel(tableModel(ndarray, headers= headers))
# good luck !

from typing_extensions import ParamSpecArgs


def main(commandBar, printOutLabel, tableWidget, tableModel,scripting = False, interact=False):
    import numpy as np
    import sys

    array = tableWidget.model().array

    def print(*args,**kw):
        colsep= kw.get('sep',' ')
        flush= kw.get('flush',True)
        linesep = kw.get('end','\n')
        if flush:
            printOutLabel.setText(colsep.join( map(str,args)))
        else:
            printOutLabel.setText(printOutLabel.text()+linesep+colsep.join( map(str,args)))
# write your functions here

#    def someFunc():
#        print('done some function')

# End of all functions
    if scripting == False and interact==False:
        try:
            command = commandBar.text()
            command = compile(command, 'user', 'exec')
            exec(command)
        except Exception as e:
            print(e)
    elif scripting:
        pass
    elif interact:
        pass