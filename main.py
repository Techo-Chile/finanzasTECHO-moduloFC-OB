# -*- coding: utf-8 -*-
import utils, write
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
from oauth2client.service_account import ServiceAccountCredentials

def main():

    #Carga propiedades
    configSectionMap = load_properties()
    
    #reemb_typeidación de semana del año
    ref_week = configSectionMap['init']['week_reference']
    if (ref_week == 'current_week'):
        year = int(datetime.datetime.now().isocalendar()[0])
        week = int(datetime.datetime.now().isocalendar()[1])
        weekday = 1 #Indicamos el inicio de semana
    else:
        year = int(datetime.datetime.now().isocalendar()[0])
        week = int(ref_week)
        weekday = 1

    output_folder_name = join(configSectionMap['init']['output_folder'],
                         datetime.datetime.now().strftime("%Y-%m-%d %H%M%S"))

    if (not os.path.exists(output_folder_name)):
        os.makedirs(output_folder_name)

    #Inicio de Semana
    week_monday = utils.ywd_to_date(year, week, weekday)
    mont_init = calendar.month_name[week_monday.month]

    #Fin de la semana
    week_friday = week_monday + timedelta(days=4)
    month_end = calendar.month_name[week_friday.month]

    #Obtenemos las tablas de GSheets - Leemos una sola vez
    reembolsos_table = get_worksheet(configSectionMap['init']['flujo_caja_url'], configSectionMap['init']['Reembolsos_worksheet_name'])
    personas_table = get_worksheet(configSectionMap['init']['employees_url'], configSectionMap['init']['employees_worksheet_name'])
    
    #Recorremos por los types de reembolsos configurados
    reemb_types = configSectionMap['init']['reemb_types'].split(',') 
    for reemb_type in reemb_types:
        output_file_name = 'Nómina Reembolsos {} Semana {} - {} {} to {} {}.xls'.format(reemb_type, week, mont_init, week_monday, month_end, week_friday)
        print('Generando archivo {}'.format(output_file_name))
        filtered_reemb = filter_reemb(reembolsos_table, reemb_type, week_monday, week_friday)
        if(len(filtered_reemb) > 0):
            writeReembfile(filtered_reemb, personas_table, output_file_name, output_folder_name)
        else:
            print('No se encuentran registros para type {}, archivo no generado'.format(reemb_type))
    
    #Fin de la ejecución:
    text = input("Enter para finalizar") 

def filter_reemb(reembolsos_table, opType, date_ini, date_end):
    lst = []
    try:
        #Filtramos por fecha
        lst = utils.filterDate(reembolsos_table, date_ini, date_end, 5)
        #Filtramos por tipo de reembolso
        lst = utils.filterType(lst, opType, 6)
        #Filtramos por fecha
        lst = utils.filterReemb(lst)
    except Exception as inst:
        print('Error filtrando reembolsos - Causa {}'.format(inst))
    return lst

def writeReembfile(filtered_list, personas_table, output_file_name, output_folder_name):
    try:
        outputFile = join(output_folder_name, output_file_name)
        outputFileWriter = write.Writer()
        outputFileWriter.write_reembolso(filtered_list, outputFile, personas_table)
        print('Archivo generado')
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
        scope = ['https://spreadsheets.google.com/feeds']
        creds = ServiceAccountCredentials.from_json_keyfile_name(join(dirname(abspath(__file__)),'client_secret.json'), scope)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(source_url)
        worksheet = spreadsheet.worksheet(worksheetName)
        rows = worksheet.get_all_values()
        return rows
    except Exception as e: 
        print(e)
        return None

if __name__ == "__main__":
    main()
