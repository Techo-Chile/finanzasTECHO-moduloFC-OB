# -*- coding: utf-8 -*-
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

class Reader:
  """Clase lectora de archivos Excel"""
  def __init__(self, fileName, index = 0):
    """Intancia un objeto de clase Reader.

    Args:
      filename (string) : Full Filepath del archivo
      index (int)       : indice de la hoja a leer (default = 0)"""
    self.fileName = fileName;
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
    ret = [];
    for rowInd in range(0, self.xlrdSheet.nrows):
      ret.append(self.xlrdSheet.row(rowInd))
    return ret

  def filterDate(self, rowList, dt, columnIndex):
    """Retorna una lista de filas cuyas fechas de pago coinciden con el argumento.

    Args:
      rowList (obj[][])      : Lista de filas
      dt (datetime.datetime)  : Instancia de datetime.datetime
      columnIndex (int)       : Indice de la columna que contiene la fecha

    Returns:
      object[][] : Lista de filas (cada fila es una fila de objetos)"""
    ret = []
    for row in rowList:
      try:
        py_date = toDatetime(row[columnIndex], self.xlrdWorkbook)
      except ValueError: #greedy: empty row
        continue;
      if dt == py_date:
        ret.append(row);     
    return ret

  def printSheet(self):
    """Imprime fila por fila la hoja activa.

    Args:
      sheetIndex (int) : indice cero-indexado de la hoja"""
    print ('Sheet name: %s' % self.xlrdSheet.name)
    num_cols = self.xlrdSheet.ncols
    for rowInd in range(0, self.xlrdSheet.nrows):
      print ('-'*40)
      print ('Row: %s' % rowInd)
      for colInd in range(0, num_cols):  # Iterate through columns
        cell_obj = self.xlrdSheet.cell(rowInd, colInd)
        print ('Column: [%s] cell_obj: [%s]' % (colInd, cell_obj))

def toDatetime(xldateCell, book):
  return xlrd.xldate.xldate_as_datetime(xldateCell.value, book.datemode)

'''
Legacy

print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))

def filterDate(filen, dt):
  """Retorna un arreglo de filas cuyas fechas de pago coinciden con el argumento.

  Args:
    dt (datetime.datetime) : Instancia de datetime.datetime

  Returns:
    object[][] : Arreglo de filas (cada fila es un arreglo de objetos)"""
  ret = []
  fname = filen # join(dirname(dirname(abspath(__file__))), 'Flujo caja 2018.xlsx') 
  xlrdWorkbook = xlrd.open_workbook(fname)
  xlrdSheet = xlrdWorkbook.sheet_by_name('Proveedores') 
  for rowInd in range(0, xlrdSheet.nrows):    # Iterate through rows
    a = xlrdSheet.row(rowInd)
    # print ('%s' % a)
    try:      
      py_date = toDatetime(a[5], xlrdWorkbook)
    except ValueError: #greedy: empty row
      continue
    if dt == py_date:
      ret.append(a);
      # print ('%s, %s, %s' % (a[0], a[1], py_date))
  return ret

def printSheet(sheetIndex):
  """Imprime fila por fila la hoja pedida del libro
  Args:
    sheetIndex (int) : indice cero-indexado de la hoja"""
  fname = join(dirname(abspath(__file__)), 'Flujo caja 2018.xlsx')
  xlrdWorkbook = xlrd.open_workbook(fname)
  #sheet_names = xlrdWorkbook.sheet_names()
  #print('Sheet Names', sheet_names)
  xlrdSheet = xlrdWorkbook.sheet_by_index(sheetIndex)
  print ('Sheet name: %s' % xlrdSheet.name)
  num_cols = xlrdSheet.ncols
  for rowInd in range(0, xlrdSheet.nrows):
    print ('-'*40)
    print ('Row: %s' % rowInd)
    for colInd in range(0, num_cols):  # Iterate through columns
      cell_obj = xlrdSheet.cell(rowInd, colInd)
      print ('Column: [%s] cell_obj: [%s]' % (colInd, cell_obj))

    '''