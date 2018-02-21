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
  d = date(2018, 2, 23)
  t = time(0, 0)
  
  inputFile = join(dirname(dirname(abspath(__file__))), 'Flujo caja 2018.xlsx')
  inputFileReader = read.Reader(inputFile);
  inputFileReader.setSheetByName('Proveedores');
  lst = inputFileReader.filterDate(inputFileReader.getRowList(), datetime.datetime.combine(d, t), 5)

  # printList(lst)

  dbFile = join(dirname(dirname(abspath(__file__))), 'BD Transferencia Proveedores..xlsx')
  dbFileReader = read.Reader(dbFile)
  dbRows = dbFileReader.getRowList();
  # printList(dbRows)

  outputFile = join(dirname(dirname(abspath(__file__))), 'sample.xlsx')   
  outputFileWriter = write.Writer(outputFile);
  outputFileWriter.setWriteList(lst)
  outputFileWriter.setFormatList(FORMAT_PROVEEDOR)
  outputFileWriter.write(dbRows)

if __name__ == "__main__":
  main()