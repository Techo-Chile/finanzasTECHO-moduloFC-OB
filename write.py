from openpyxl import Workbook
from openpyxl.styles import Font, Color, PatternFill, Border, Side

def write(filename):
	wb = Workbook()

	# grab the active worksheet
	ws = wb.active

	# Data can be assigned directly to cells
	# ws['A1'] = 42

	ws.append(['Rut Beneficiario', 'Nombre Beneficiario', 'Cod. Modalidad', 'Cod Banco', 'Cta Abono', 
	  'N Factura 1', 'Monto 1', 'N Factura 2', 'Monto 2', 'N Factura 3', 'Monto 3', 'N Factura 3',
	  'Monto 4', 'N Factura 4', 'Monto 5', 'N Factura 5', 'Monto 6', 'N Factura 6', 'Monto 7',
	  'N Factura 7', 'Monto 8', 'N Factura 8', 'Monto 9', 'N Factura 9', 'Monto 10', 'N Factura 10',
	  'Monto 11', 'N Factura 11', 'Monto Total'])

	sd = Side(border_style = 'thin', color = 'FF000000')

	for cell in ws["1:1"]:
	  cell.font = Font(name='Arial', size = 9, color = 'FF0000FF')
	  cell.fill = PatternFill(fill_type = 'solid', start_color = "ffcccccc")
	  cell.border = Border(left = sd, right = sd, top = sd, bottom= sd)

	#import datetime
	#ws['A2'] = datetime.datetime.now()
	ws.column_dimensions["A"].width = 15
	ws.column_dimensions["B"].width = 40
	ws.column_dimensions["C"].width = 15
	ws.column_dimensions["D"].width = 15
	ws.column_dimensions["E"].width = 15
	ws.column_dimensions["G"].width = 15
	ws.column_dimensions["F"].width = 15
	ws.column_dimensions["AC"].width = 15

	# Save the file
	wb.save(filename)