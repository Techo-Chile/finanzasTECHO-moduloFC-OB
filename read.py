"""
Actualmente genera un archivo sample.xlsx con el formato de las nominas de Office Banking.
Todo:
  Hacer más filtros.
  Filtrar conmutativamente.
  Completar información de nominas
"""
from __future__ import print_function
from xlrd.sheet import ctype_text
import xlrd
import write
from datetime import date, time
from datetime import datetime  
from datetime import timedelta 

class Reader:
    """Clase lectora de archivos Excel"""

    def __init__(self, fileName, index=0):
        """Intancia un objeto de clase Reader.
    Args:
      filename (string) : Full Filepath del archivo
      index (int)       : indice de la hoja a leer (default = 0)"""
        self.fileName = fileName
        self.xlrdWorkbook = xlrd.open_workbook(fileName)
        self.xlrdSheet = self.xlrdWorkbook.sheet_by_index(index)

    def setSheetByIndex(self, index):
        self.xlrdSheet = self.xlrdWorkbook.sheet_by_index(index)

    def setSheetByName(self, name):
        self.xlrdSheet = self.xlrdWorkbook.sheet_by_name(name)

    def getRowList(self):
        """Genera una lista de filas.
    Returns:
      object[][] : lista de filas (cada fila es una lista de objetos)"""
        ret = []
        for rowInd in range(0, self.xlrdSheet.nrows):
            ret.append(self.xlrdSheet.row(rowInd))
        return ret

    def filterDate(self, rowList, dt, columnIndex):
        """Retorna una lista de filas cuyas fechas de pago coinciden con el argumento.
    Args:
      rowList (obj[][])       : Lista de filas
      dt (datetime.datetime)  : Instancia de datetime.datetime
      columnIndex (int)       : Indice de la columna que contiene la fecha
    Returns:
      object[][] : Lista de filas (cada fila es una fila de objetos)"""
        ret = []
        for row in rowList:
            try:
                py_date = toDatetime(row[columnIndex], self.xlrdWorkbook)
            except ValueError:  #greedy: empty row
                continue
            if (dt == py_date) | (dt - timedelta(days=7) < py_date):
                ret.append(row)
        return ret

    def filterType(self, rowList, orderType, columnIndex):
        """Retorna una lista de filas cuyos tipos de orden (OP, TDV, etc) coinciden
    con el argumento.
    Args:
      rowList (obj[][])       : Lista de filas
      orderType (string)      : Tipo de orden de compra
      columnIndex (int)       : Indice de la columna que contiene el tipo"""
        ret = []
        for row in rowList:
            if orderType in row[columnIndex].value:
                ret.append(row)
        return ret
    
    def filterReemb(self, rowList):
        ret = []
        for row in rowList:
            try:
              if ('X' in row[0].value) | ('x' in row[0].value):
                  ret.append(row)
            except Exception:
              continue
        return ret

    def printSheet(self):
        """Imprime fila por fila la hoja activa.
    Args:
      sheetIndex (int) : indice cero-indexado de la hoja"""
        print('Sheet name: %s' % self.xlrdSheet.name)
        num_cols = self.xlrdSheet.ncols
        for rowInd in range(0, self.xlrdSheet.nrows):
            print('-' * 40)
            print('Row: %s' % rowInd)
            for colInd in range(0, num_cols):  # Iterate through columns
                cell_obj = self.xlrdSheet.cell(rowInd, colInd)
                print('Column: [%s] cell_obj: [%s]' % (colInd, cell_obj))


def toDatetime(xldateCell, book):
    return xlrd.xldate.xldate_as_datetime(int(xldateCell.value), book.datemode)
