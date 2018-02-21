# -*- coding: utf-8 -*-
import read, write
import datetime
from datetime import date, time
from os.path import join, dirname, abspath

def printArray(array):
  for e in array:
    print (e)
    print('-'*40)

def main():
  d = date(2018, 2, 23)
  t = time(0, 0)
  inputFile = join(dirname(dirname(abspath(__file__))), 'Flujo caja 2018.xlsx') 
  outputFile = join(dirname(dirname(abspath(__file__))), 'sample.xlsx') 

  rd = read.Reader(inputFile);

  rd.setSheetByName('Proveedores');

  # write.write(outputFile)

  printArray(rd.filterDate(rd.getRowArray(), datetime.datetime.combine(d, t), 5))

if __name__ == "__main__":
  main()