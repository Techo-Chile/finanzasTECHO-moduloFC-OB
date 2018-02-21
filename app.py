# -*- coding: utf-8 -*-
import read, write
import datetime
from datetime import date, time
from os.path import join, dirname, abspath

FORMAT_PROVEEDOR = [0, 1, None, None, None, 2, 3];

def printList(list):
  for e in list:
    print (e)
    print('-'*40)

def main():
  #YYYY MM YY
  fecha = date(2018, 2, 16)
  provider('sample.xlsx', 'OP', 8, fecha)
  provider('sample2.xlsx', 'TDV', 8, fecha)

def reembolsosMain():
  print 2

def provider(outputName, opType, column, fecha):
  """Genera la nomina para proveedores.

  Args:
    outputName (string) : full filepath del archivo de destino
    opType (string)     : tipo de operación ('OP', 'TDV', etc)
    column ("""
  
  hora = time(0, 0)
  
  inputFile = join(dirname(dirname(abspath(__file__))), 'Flujo caja 2018.xlsx')
  inputFileReader = read.Reader(inputFile);
  inputFileReader.setSheetByName('Proveedores');
  lst = inputFileReader.filterDate(inputFileReader.getRowList(), datetime.datetime.combine(fecha, hora), 5)
  lst = inputFileReader.filterType(lst, opType, column)

  printList(lst)

  dbFile = join(dirname(dirname(abspath(__file__))), 'BD Transferencia Proveedores..xlsx')
  dbFileReader = read.Reader(dbFile)
  dbRows = dbFileReader.getRowList();
  # printList(dbRows)

  outputFile = join(dirname(dirname(abspath(__file__))), outputName)   
  outputFileWriter = write.Writer(outputFile);
  outputFileWriter.setWriteList(lst)
  outputFileWriter.setFormatList(FORMAT_PROVEEDOR)
  outputFileWriter.write(dbRows)

if __name__ == "__main__":
  main()