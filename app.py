from openpyxl import load_workbook
from coordinatesTools import *
from datetime import timedelta, datetime, tzinfo
from constants import *

def numberfy(string):
  return sum([(ord(string[i]) - ord('0')) * 10 ** (len(string) - 1 - i) for i in range(len(string) - 1, -1, -1)]);

def getDMY(value):
  return [numberfy(n) for n in value.split('/')];

def filter(date):
  wb = load_workbook('Flujo de Caja.xlsx');
  sheet = wb['Proveedores']
  tr = False;
  for row in tuple(sheet.rows):
    val = row[P_FdP].value
    if not isinstance(val,  datetime):
      tr = False
      continue;
    if not tr and date == val:
      tr = True;
    if tr:
      print row[P_Pro].value, val;

dt1 = datetime(2018, 1, 19, 0, 0)

print ""

filter(dt1)

print ""