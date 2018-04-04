0  # -*- coding: utf-8 -*-
"""TODO:
  write.write() requiere reformatting para permitir el uso de una base de datos de RUTS y hacer string matching
  con los nombres.
"""
import filters, write
import datetime
from datetime import date, time
from os.path import join, dirname, abspath
import configparser
from shutil import copyfile
import os
import calendar
from datetime import timedelta
import sys, traceback
import pygsheets, oauth2client
from oauth2client.service_account import ServiceAccountCredentials
import gspread

def main():

    #Carga propiedades
    configSectionMap = load_properties()
    #Validación de fecha
    ref_date = configSectionMap['init']['reference_date']
    if (ref_date == 'today'):
        fecha = datetime.datetime.now().date()
    else:
        fecha = date(
            int(ref_date.split('/')[0]), int(ref_date.split('/')[1]),
            int(ref_date.split('/')[2]))

    output_folder = join(configSectionMap['init']['output_folder'],
                         datetime.datetime.now().strftime("%Y-%m-%d %H%M%S"))

    if (not os.path.exists(output_folder)):
        os.makedirs(output_folder)

    '''
      Ejecución de archivos de nómina:
        1. Nomina Dev. socios semana 09 - 26 al 02 marzo
        2. Nomina FPR Operacionales semana 09 - 26 al 02 marzo
        
        #--2/4/18: Se cambia el manejo de proveedores por la solución de 2WIN
        3. Nomina Proveedores Operacionales semana  09 - 26 al 02 marzo
        4. Nomina Proveedores TDV semana  09 - 26 al 02 marzo
        
        #--2/4/18: Se cambia el manejo de proveedores por la solución de 2WIN
        5. Nomina Reembolsos Operacionales semana 09 - 26 al 02 marzo
        6. Nomina Reembolsos TDE semana 09 - 26 al 02 marzo
        7. Nomina Reembolsos TDV semana 09 - 26 al 02 marzo
    '''

    #Data para nombre de archivos
    week_of_year = fecha.isocalendar()[1]

    #Inicio de Semana
    week_monday = fecha + datetime.timedelta(days=-fecha.weekday())
    mont_init = calendar.month_name[week_monday.month]

    #Fin de la semana
    week_friday = fecha + datetime.timedelta(
        days=-fecha.weekday()) + timedelta(days=5)
    mont_end = calendar.month_name[week_friday.month]

    #Obtenemos las tablas de GSheets - Leemos una sola vez
    reembolsos_table = get_worksheet(configSectionMap['init']['flujo_caja_url'], configSectionMap['init']['Reembolsos_worksheet_name'])
    personas_table = get_worksheet(configSectionMap['init']['employees_url'], configSectionMap['init']['employees_worksheet_name'])
    
    #Recorremos por los types de reembolsos configurados
    reemb_types = configSectionMap['init']['reemb_types'].split(',') 
    for val in reemb_types:
        reembolsos(reembolsos_table, personas_table, 'Nómina Reembolsos {} Semana {} - {} {} to {} {}.xlsx'.format(val,
        week_of_year, mont_init, week_monday, mont_end, week_friday), val,
               fecha, output_folder, configSectionMap)

def reembolsos(reembolsos_table, personas_table, outputName, opType, fecha, output_folder, configSectionMap):
    try:
        #Filtramos por fecha
        lst = filters.filterDate(reembolsos_table,
                                         datetime.datetime.combine(
                                             fecha, time(0, 0)), 5)
        lst = filters.filterType(lst, opType, 6)
        #Identificamos cuáles son reembolsos
        lst = filters.filterReemb(lst)
        #Obtenemos nombre de archivo de salida
        outputFile = join(output_folder, outputName)
        #Escribimos el archivo con el formato
        outputFileWriter = write.Writer()
        outputFileWriter.write_reembolso(lst, outputFile, personas_table)
        print('Archivo generado: {}'.format(outputFile))
    except Exception as inst:
        print('Error generando archivo {} - Causa {}'.format(outputFile, inst))

def load_properties():
    config = configparser.ConfigParser()
    config = configparser.ConfigParser()
    config.sections()
    config.read(join(dirname(abspath(__file__)), 'finanzas.ini'))
    return config

def get_worksheet(source_url, worksheetName):
    try:
        client = pygsheets.authorize(outh_file=join(dirname(abspath(__file__)),'client_secret.json'), outh_nonlocal=True)
        spreadsheet = client.open_by_url(source_url)
        worksheet = spreadsheet.worksheet_by_title(worksheetName)
        rows = worksheet.get_all_values(returnas='matrix', majdim='ROWS', include_empty=True)
        return rows
    except Exception as e: 
        print(e)
        return None


if __name__ == "__main__":
    main()
