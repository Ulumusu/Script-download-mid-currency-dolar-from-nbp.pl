#
#This script download mid currency (dolar) from nbp.pl site.
#If you want check, you must put this file to:
# .....\LibreOffice\share\Scripts\python\pythonSamples
#


#import module from LibreOffice
import uno, unohelper
from com.sun.star.util import XModifyListener
import os
from com.sun.star.table.CellContentType import TEXT, EMPTY, VALUE, FORMULA
import json

#other library from Python, default
import sys
import urllib.request
from datetime import datetime


def main_function(cell, cell2):
        cell.getFormula()
        date = cell.getValue()
        lt = date
        dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(lt) - 2)
        ls = str(dt)
        g = ls[0:4] +"-"+ ls[5:7] +"-"+ ls[8:10]

        #Link to the nbp.pl, download currency, delete not important words.

        try:
                link = 'http://api.nbp.pl/api/exchangerates/rates/A/USD/'+ str(g) +'/?format=json&fbclid=IwAR1ESI9JytcV7S6lisaIcYzChj-uoY6XW0soUB1jznG2CaiJHrT1RVyOQ94'
                response = urllib.request.urlopen(link)
                html = response.read()
                string_link = str(html)
                a = string_link.find("\"mid\":")
                del_not_important = string_link[a:]
                b = del_not_important.find(':')
                c = del_not_important.find('}')
                clean_number = del_not_important[b+1:c]
        except:
                clean_number = "Nie znaleziono"

        #Put value to the row
        e = "" + str(clean_number)
        cell2.setString(e)
        

#Check all date function
def check_many_values(*args):
        oSheet = XSCRIPTCONTEXT.getDocument().getSheets().getByIndex(0)
        #Start index, first row
        cell_index = 2

        while True:

                #Cell where you have date
                name_of_cell = "B" + str(cell_index)
                cell = oSheet.getCellRangeByName(name_of_cell)
                cell2 = oSheet.getCellRangeByName("C" + str(cell_index))
                cell.getType()

                if cell.getType() == EMPTY:
                        break
                else:
                        main_function(cell, cell2)
                        cell_index += 1



#Check only last date function
def check_last_value(*args):
        oSheet = XSCRIPTCONTEXT.getDocument().getSheets().getByIndex(0)
        #start index, first row
        cell_index = 2
        empty_cell_index = 0

        #This loop find empty index
        while empty_cell_index == 0:
                name_of_cell = "B" + str(cell_index)
                cell = oSheet.getCellRangeByName(name_of_cell)
                cell.getType()

                if cell.getType() == EMPTY:
                    empty_cell_index += 1
                else:
                    cell_index += 1

        #last index
        last_cell_index = cell_index - 1

        #Cell where you have date
        cell = oSheet.getCellRangeByName("B" + str(last_cell_index))
        cell2 = oSheet.getCellRangeByName("C" + str(last_cell_index)) 
        main_function(cell, cell2)


def check_few_values(*args):
        oSheet = XSCRIPTCONTEXT.getDocument().getSheets().getByIndex(0)
        #Start index, first row
        cell_index = 2

        while True:

                #Cell where you have date
                name_of_cell = "B" + str(cell_index)
                second_cell = "C" + str(cell_index)
                cell = oSheet.getCellRangeByName(name_of_cell)
                cell2= oSheet.getCellRangeByName(second_cell)
                cell.getType()
                cell2.getType()

                   
                        
                if cell.getType() == EMPTY:
                        break
                elif cell2.getType() == EMPTY:
                        main_function(cell, cell2)
                        cell_index += 1    
                else:
                        cell_index += 1


#Load function
g_exportedScripts = check_many_values, check_few_values, check_last_value,
