# -*- coding: utf-8 -*-
"""Todo:
Agregar control de flujo para ajustar los datos al formato de la nomina"""
from openpyxl import Workbook
from openpyxl.styles import Font, Color, PatternFill, Border, Side
from os.path import join, dirname, abspath

class Writer:
  def __init__(self, filename, writeList = [], writeFormat = []):
    """Instancia un objeto de clase Writer.

    Args:
      filename (string)       : Fullpath del archivo
      writeList (object[][])  : Lista de filas (cada fila es una lista de objetos)
      writeFormat (int[])     : Lista de enteros en donde cada entero Y en la posición X 
                                indica que el elemento en posición Y de la lista writeList
                                debe ir en la columna X (0-indexada) de la nomina de pago.
                                Ciertos valores deben ser djeados en Nene (Cod Banco), Cuenta"""
    self.filename = filename
    self.format = writeFormat
    self.list = writeList

  def setWriteList(self, list):
    self.list = list

  def setFormatList(self, list):
    self.format = list

  def write(self, dbFileRows = [None]*7):
    """Escribe en la nomina los datos de la lista."""
    if not self.list:
      print ("No ha definido una lista para escribir")
      return
    if not self.format:
      print ("No ha definido una lista de formato")
      return
    wb = Workbook()
    ws = wb.active

    # Column Header
    self.writeHeader(ws);

    #TODO: FORMAT LOGIC GOES HERE    
    for row in self.list:
      data = [None] * 7
      for i in range(len(self.format)):
        try:
          data[i] = row[self.format[i]];
        except TypeError: #Greedy, None index
          continue
      data = self.completeRow(data, dbFileRows)
      if data[4]:
        ws.append(data)

    wb.save(self.filename);

  def writeHeader(self, ws):
    ws.append(['Rut Beneficiario', 'Nombre Beneficiario', 'Cod. Modalidad', 'Cod Banco',
      'Cta Abono', 'N Factura 1', 'Monto 1', 'N Factura 2', 'Monto 2', 'N Factura 3',
      'Monto 3', 'N Factura 3', 'Monto 4', 'N Factura 4', 'Monto 5', 'N Factura 5',
      'Monto 6', 'N Factura 6', 'Monto 7', 'N Factura 7', 'Monto 8', 'N Factura 8',
      'Monto 9', 'N Factura 9', 'Monto 10', 'N Factura 10', 'Monto 11', 'N Factura 11',
      'Monto Total']) 
    sd = Side(border_style = 'thin', color = 'FF000000')
    for cell in ws["1:1"]:
      cell.font = Font(name='Arial', size = 9, color = 'FF0000FF')
      cell.fill = PatternFill(fill_type = 'solid', start_color = "ffcccccc")
      cell.border = Border(left = sd, right = sd, top = sd, bottom= sd)
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["F"].width = 15
    ws.column_dimensions["AC"].width = 15

  def completeRow(self, data, dbFileRows):
    '''data[0] : RUT
    data[1] : Nombre
    data[2] : Codigo modalidad ---
    data[3] : Cod Banco --
    data[4] : Cuenta Abono --
    data[5] : Factura
    data[6] : Monto'''
    ret = [None]*29
    ret[0] = data[0].value
    ret[1] = data[1].value
    ret[2] = 3
    for row in dbFileRows:
      try:
        if row[0].value[:-1] == data[0].value[:-1]:
          ret[3] = row[3].value
          try:
            ret[4] = int(row[4].value)
          except:
            ret[4] = row[4].value
      except:
        continue
    ret[5] = data[5].value
    ret[6] = data[6].value
    ret[28] = ret[6]
    print '-',
    return ret
  

