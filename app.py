0  # -*- coding: utf-8 -*-
"""TODO:
  write.write() requiere reformatting para permitir el uso de una base de datos de RUTS y hacer string matching
  con los nombres.
"""
import read, write
import datetime
from datetime import date, time
from os.path import join, dirname, abspath
import configparser
from shutil import copyfile
import os
import calendar
from datetime import timedelta
import sys, traceback

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
    #Copia archivos fuente desde servidor BODEGA a ruta de output
    output_folder = copy_base_files(configSectionMap)
    #print('Working directory {%0}', output_folder)
    '''
      Ejecución de archivos de nómina:
        1. Nomina Dev. socios semana 09 - 26 al 02 marzo
        2. Nomina FPR Operacionales semana 09 - 26 al 02 marzo
        3. Nomina Proveedores Operacionales semana  09 - 26 al 02 marzo
        4. Nomina Proveedores TDV semana  09 - 26 al 02 marzo
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

    #3. Nomina Proveedores Operacionales semana  09 - 26 al 02 marzo
    provider('Nómina Proveedores OP Semana {} - {} {} to {} {}.xlsx'.format(
        week_of_year, mont_init, week_monday, mont_end, week_friday), 'OP',
             fecha, output_folder)

    #4. Nomina Proveedores TDV semana  09 - 26 al 02 marzo
    provider('Nómina Proveedores TDV Semana {} - {} {} to {} {}.xlsx'.format(
        week_of_year, mont_init, week_monday, mont_end, week_friday), 'TDV',
             fecha, output_folder)

    #5. Nomina Reembolsos Operacionales semana 09 - 26 al 02 marzo
    reembolsos('Nómina Reembolsos OP Semana {} - {} {} to {} {}.xlsx'.format(
        week_of_year, mont_init, week_monday, mont_end, week_friday), 'OP',
               fecha, output_folder)

    #6. Nomina Reembolsos TDE semana 09 - 26 al 02 marzo
    reembolsos('Nómina Reembolsos TDE Semana {} - {} {} to {} {}.xlsx'.format(
        week_of_year, mont_init, week_monday, mont_end, week_friday), 'TDE',
               fecha, output_folder)

    #7. Nomina Reembolsos TDV semana 09 - 26 al 02 marzo
    reembolsos('Nómina Reembolsos TDV Semana {} - {} {} to {} {}.xlsx'.format(
        week_of_year, mont_init, week_monday, mont_end, week_friday), 'TDV',
               fecha, output_folder)

    #7. Nomina Reembolsos FNDR semana 09 - 26 al 02 marzo
    reembolsos('Nómina Reembolsos FNDR Semana {} - {} {} to {} {}.xlsx'.format(
        week_of_year, mont_init, week_monday, mont_end, week_friday), 'FNDR',
               fecha, output_folder)


def reembolsos(outputName, opType, fecha, output_folder):
    try:
        inputFile = join(output_folder + '/WorkingFiles/',
                         'Flujo caja 2018.xlsx')
        inputFileReader = read.Reader(inputFile)
        inputFileReader.setSheetByName('Reembolsos')
        #Filtramos por fecha
        lst = inputFileReader.filterDate(inputFileReader.getRowList(),
                                         datetime.datetime.combine(
                                             fecha, time(0, 0)), 4)
        lst = inputFileReader.filterType(lst, opType, 5)
        #Identificamos cuáles son reembolsos
        lst = inputFileReader.filterReemb(lst)
        #Cargamos BD Proveedores
        dbFile = join(output_folder + '/WorkingFiles/', 'Personas.xlsx')
        dbFileReader = read.Reader(dbFile)
        dbRows = dbFileReader.getRowList()
        #Obtenemos nombre de archivo de salida
        outputFile = join(output_folder, outputName)
        #Escribimos el archivo con el formato
        outputFileWriter = write.Writer()
        outputFileWriter.write_reembolso(lst, outputFile, dbRows)
        print('Archivo generado: {}'.format(outputFile))
    except Exception as inst:
        print('Error generando archivo {} - Causa {}'.format(outputFile, inst))


def provider(outputName, opType, fecha, output_folder):
    """Genera la nomina para proveedores.
  Args:
    outputName (string) : full filepath del archivo de destino
    opType (string)     : tipo de operación ('OP', 'TDV', etc)
    fecha (date.date)   : fecha a filtrar"""
    try:
        inputFile = join(output_folder + '/WorkingFiles/',
                         'Flujo caja 2018.xlsx')
        inputFileReader = read.Reader(inputFile)
        inputFileReader.setSheetByName('Proveedores')
        #Realiza filtro de fecha
        lst = inputFileReader.filterDate(inputFileReader.getRowList(),
                                         datetime.datetime.combine(
                                             fecha, time(0, 0)), 5)
        lst = inputFileReader.filterType(lst, opType, 8)
        #Leemos la lista de proveedores
        dbFile = join(output_folder + '/WorkingFiles/',
                      'BD Transferencia Proveedores..xlsx')
        dbFileReader = read.Reader(dbFile)
        dbRows = dbFileReader.getRowList()
        #Obtenemos nombre del archivo a imprimir
        outputFile = join(output_folder, outputName)
        outputFileWriter = write.Writer()
        outputFileWriter.write_provider(lst, outputFile, dbRows)
        print('Archivo generado: {}'.format(outputFile))
    except Exception as inst:
        print('Error generando archivo {} - Causa {}'.format(outputFile, inst))


def load_properties():
    config = configparser.ConfigParser()
    config.sections()
    print (os.name)
    config.read(join(dirname(abspath(__file__)),  'finanzas.linux.ini' if os.name == 'nt' else 'finanzas.win.ini'))
    return config


def copy_base_files(propertiesMap):

    output_folder = propertiesMap['init']['output_folder']
    proveed_file = propertiesMap['init']['db_proveedores_file']
    flujo_caja_file = propertiesMap['init']['flujo_caja_file']
    employees_file = propertiesMap['init']['employees_file']

    #Validación de conectividad con server
    if(os.path.exists(proveed_file)):
        print("No se pueden recuperar los archivos de BODEGA (Abrí primero la carpeta manualmente e intentá nuevamente)")
        sys.exit(0)

    output_folder = join(output_folder,
                         datetime.datetime.now().strftime("%Y-%m-%d %H%M%S"))
    if (not os.path.exists(output_folder)):
        os.makedirs(output_folder)
        os.makedirs(output_folder + '/WorkingFiles/')

    copyfile(proveed_file,
             join(output_folder + '/WorkingFiles/',
                  os.path.basename(proveed_file)))
    copyfile(flujo_caja_file,
             join(output_folder + '/WorkingFiles/',
                  os.path.basename(flujo_caja_file)))
    copyfile(employees_file,
             join(output_folder + '/WorkingFiles/',
                  os.path.basename(employees_file)))
    return output_folder


if __name__ == "__main__":
    main()
