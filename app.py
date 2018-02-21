# -*- coding: utf-8 -*-
import read, write
import datetime
from datetime import date, time
from os.path import join, dirname, abspath

def printList(list):
  for e in list:
    print (e)
    print('-'*40)

def main():
  d = date(2018, 2, 23)
  t = time(0, 0)
  inputFile = join(dirname(dirname(abspath(__file__))), 'Flujo caja 2018.xlsx') 
  outputFile = join(dirname(dirname(abspath(__file__))), 'sample.xlsx') 

  rd = read.Reader(inputFile);
  wr = write.Writer(outputFile);

  rd.setSheetByName('Proveedores');

  printList(rd.filterDate(rd.getRowList(), datetime.datetime.combine(d, t), 5))

  wr.write()

if __name__ == "__main__":
  main()