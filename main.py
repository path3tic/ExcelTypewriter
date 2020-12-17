
import xlwings as xw

def print_hi(name):
    xw.Book(r'test.xlsb').set_mock_caller()
    wb = xw.Book.caller()

    celldes = wb.app.selection
    rowNum = celldes.row
    colNum = celldes.column

    cellsrc = wb.sheets["BG"].range(rowNum, colNum).formula

    wb.sheets["Sheet1"].range(rowNum, colNum).formula = cellsrc


if __name__ == '__main__':
    print_hi('PyCharm')

