import pyautogui
import time
import keyboard
import xlwings as xw

#Environment settings
file = r'test.xlsb'
mainSH = "BG"

def caller():

    #Getting Source and destination
    xw.Book(file).set_mock_caller()
    wb = xw.Book.caller()

    #destination
    celldes = wb.app.selection
    rowNum = celldes.row
    colNum = celldes.column

    #source
    cellsrc = wb.sheets[mainSH].range(rowNum, colNum).formula

    typeit(cellsrc)


def typeit(s):

    # Getting Source and destination
    xw.Book(file).set_mock_caller()
    wb = xw.Book.caller()

    # destination
    celldes = wb.app.selection
    rowNum = celldes.row
    colNum = celldes.column

    for element in s:
        pyautogui.write(element)
    #wb.sheets["Sheet1"].range(rowNum, colNum).formula = s


if __name__ == '__main__':

    keyboard.add_hotkey('insert', caller)
    print('ppp')
    keyboard.wait('insert+esc')



