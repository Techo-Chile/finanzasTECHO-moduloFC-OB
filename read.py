# -*- coding: utf-8 -*-
"""
Actualmente genera un archivo sample.xlsx con el formato de las nominas de Office Banking.
Todo:
  Hacer más filtros.

  Filtrar conmutativamente.

  Completar información de nominas

  Mover main a su propia clase, refactoring y más modulación
"""
from __future__ import print_function
from os.path import join, dirname, abspath
from xlrd.sheet import ctype_text
import xlrd
import write
import datetime
from datetime import date, time

def toDatetime(xldateCell, book):
  return xlrd.xldate.xldate_as_datetime(xldateCell.value, book.datemode)


def filterDate(dt):
  """Retorna un arreglo de filas cuyas fechas de pago coinciden con el argumento.

  Args:
    dt (datetime.datetime) : Instancia de datetime.datetime

  Returns:
    object[][] : Arreglo de filas (cada fila es un arreglo de objetos)"""
  ret = []
  fname = join(dirname(abspath(__file__)), 'Flujo caja 2018.xlsx') 
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


def printSheet(sheetIndex = 0):
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

def main():
  d = date(2018, 2, 23)
  t = time(0, 0)
  print ('%s' % filterDate(datetime.datetime.combine(d, t)))
  write.write("sample.xlsx")

if __name__ == "__main__":
  main()

'''
print('(Column #) type:value')
for idx, cell_obj in enumerate(row):
cell_type_str = ctype_text.get(cell_obj.ctype, 'unknown type')
print('(%s) %s %s' % (idx, cell_type_str, cell_obj.value))
'''