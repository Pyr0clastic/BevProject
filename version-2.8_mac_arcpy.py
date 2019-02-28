# -*- coding: utf-8 -*-
# import sys
# import arcpy
import numpy as np
import pandas as pd
# import xlsxwriter
# import xlrd
import os
import re

encoding = "utf-8"

# Displays all Rows in a DataFrame; default is limited to around hundred
pd.set_option('display.max_rows', None)
# pd.set_option('display.max_columns', None)

cwd = os.getcwd()
print(cwd)

x = '/DATEN201617_VGL.xlsx'  # TODO delete lines 20 until 30
y = '/DATEN201516_VGL.xlsx'

inp = input("Which Dataset?: ")
if inp == "1":
    file = y
elif inp == "2":
    file = x
else:
    file = y

# TODO uncomment filePath = arcpy.GetParameterAsText(0).replace('\\','\u005C')
filePath = os.getcwd()
# TODO: uncomment arcpy.AddMessage ("Imported from : " + filePath)
fileFolder, fileName = os.path.split(filePath)
print(filePath)
print(fileFolder, fileName)
os.chdir(fileFolder)
xl = pd.ExcelFile(filePath + file)  # TODO: remove line
# Load spreadsheet
# TODO: xl = pd.ExcelFile(file)

# Print the sheet names
print(xl.sheet_names)


def find_sheet(excelFile):
    """
    This function searches for the exact sheetnames containing
    "Bevoelkerungsstand" and "Bevoelkerungsentwicklung" in the ExcelFile
    It also checks the survey year of sheet Bevoelkerungsstand
    Parameters:
    -----------
    excelFile : the loaded excel spreasdsheet
    """
    wsheetName_namen = xl.sheet_names
    cache = []
    for wsheetName in wsheetName_namen:
        if "stand" in wsheetName:
            cache.append(wsheetName)
            surveyYear = re.findall(r'\d+', wsheetName)
            for number in surveyYear:
                if len(number) == 4:
                    cache.append(number)
                    # yield wsheetName, number
                    # return wsheetName, number
        elif "entwicklung" in wsheetName:
            cache.append(wsheetName)
        else:
            continue
    return cache


def subset_dpr(xl, wsheetName, year):
    """
    Function erwerbData's purpose is to select just the data needed for the
    calculation of the parameter dependency ratio (Abhaengigenquote)

    Parameters:
    -----------
    wsheetName:     The exact name of the needed excel worksheet that
                    contains Bevoelkerungsstand
    year:           Survey year
    xl:             loaded excel spreadsheet
    """
    df_raw = xl.parse(wsheetName)
    all = [
        'Wbv2016', 'Wbv2016v00b04', 'Wbv2016v05b09', 'Wbv2016v10b14',
        'Wbv2016v15b19', 'Wbv2016v20b24', 'Wbv2016v25b29', 'Wbv2016v30b34',
        'Wbv2016v35b39', 'Wbv2016v40b44', 'Wbv2016v45b49', 'Wbv2016v50b54',
        'Wbv2016v55b59', 'Wbv2016v60b64', 'Wbv2016v65b69', 'Wbv2016v70b74',
        'Wbv2016v75b79', 'Wbv2016v80b84', 'Wbv2016v85+'
    ]
    ##########################
    new = []
    dict = {}

    if year != "2016":
        for i in all:
            new.append(i.replace("2016", year))
            # all.append(i.replace("2016", year))
            q = i.replace("2016", year)
            dict[i] = q
        all.clear()
        all = new
        df_raw.rename(columns=dict, inplace=True)
    ########################

    # BUG If not this but another function is executed other function needs
    # this columns too for a shp join. Possible Solution => Put whole code in
    # class that has these columns as class attributes
    df_erw = df_raw[["Kennz", "Name", all[0]]]
    erwerbslos = set(all[1:4] + all[14:19])
    erwerbsf = set(all[4:14])
    # slices based on column index; creates new slick DataFrame with values
    # Kennz, Name, and Wbv
    df_erw = df_raw.iloc[:, 2:5].copy()
    # Create sum based on data in set erwerbsf; axis = 1, build sum of each
    # row in the data provided in erwerbsf
    df_erw["erwerbsfaehig"] = df_raw.drop(erwerbsf, axis=1).sum(axis=1)
    # Create sum based on data in set erwerbslos; axis = 1, build sum of each
    # row in the data provided in erwerbslos
    df_erw["nicht_erwerbsfaehig"] = df_raw.drop(erwerbslos, axis=1).sum(axis=1)
    return df_erw, df_raw  # dict, all


def dpr_calc(erwerbsfaehig, nicht_erwerbsfaehig):
    """
    dpr_calc function calculats dependency ratio (Abhaengigenquote)
    Parameters:
    -----------
    erwerbsfaehig:          Column in Dataframe containing Erwerbsfaehige
    nicht_erwerbsfaehig:    Column in Dataframe containing nicht Erwerbsfaehige
    """
    x = (nicht_erwerbsfaehig / erwerbsfaehig) * 100
    return round(x, 2)


def subset_aagr(xl, bevEntwSheet):
    """
    this function selects and aggregates Data for the calculation of the
    average annual growth rate (durchschn. jaehrl. Bevoelkerungsveraenderung)
    Parameters:
    -----------
    xl: loaded excel spreadsheet
    bevEntwSheet:   excel sheet with Bevoelkerungsentwicklung information
    startYear:      startYear     # TODO Input has to be selectable over toolbox
    endYear:                      # TODO Has to be selectable over toolbox
    """

    startYear = input("Ausgangszeitpunkt: ")
    endYear = input("Endzeitpunkt: ")
    timeDiff = int(endYear) - int(startYear)
    print("Ausgangszeitpunkt: {} \nEndzeitpunkt: {} \nZeitdifferenz: {}".format(
        startYear, endYear, timeDiff))

    df_bevEntw = xl.parse(bevEntwSheet)
    startY = "Wbv" + startYear
    endY = "Wbv" + endYear
    df_aagr = df_bevEntw[["Kennz", "Name", startY, endY]]

    print(df_aagr)

    return df_aagr, timeDiff


def aagr_calc(startColumn, endColumn, timeDelta):
    """Calculates average annual growth rate.

    Args:
        startColumn (DataFrame column): population at interval start.
        endColumn (DataFrame column): population at interval end.
        timeDelta (int): end year - start year.

    Returns:
        DataFrame column: Returns average annual growth rate.

    """

    result = ((endColumn / startColumn)**(1 / timeDelta) - 1) * 100
    return result


def exportCSV(
        DataFrame1, DataFrame2=None
):  # FIXME Possibly more data should be exported (Michael's function)
    """Exports calculated data as csv file.
    Args:
        DataFrame1 (DataFrame): pandas DataFrame to export.
        DataFrame2 (DataFrame): pandas DataFrame to export.

    Returns:
        .csv: Returns calculated DataFrame(s) as csv file.

    """
    if DataFrame2 is not None:
        df_joined = pd.merge(DataFrame1, DataFrame2, on=['Kennz', 'Name'])
        df_joined.to_csv("~/test_csv.csv", sep=';', encoding='utf-8')
    else:
        DataFrame1.to_csv("~/test_csv.csv", sep=';', encoding='utf-8')


print("______________________________________________________________")
print("|                                                            |")
print("|  This software is purely experimental!                     |")
print("|  By using this software the user acknowledges that         |")
print("|  he/she takes full responsibility for any consequences     |")
print("|  which may arise from it's application.                    |")
print("|                                                            |")
print("|                                                            |")
print("|                                   © Marko Csenar  2019     |")
print("|                                   © Michael Möstl 2019     |")
print("|                                   © Robert Brand  2019     |")
print("|                                   © Kuanlun Chiem 2019     |")
print("|                                                            |")
print("|                                                            |")
print("--------------------------------------------------------------")
print(
    "_________________________________________________________________________")
# wsheetName, number = sheet_bevstand(xl)
sheetInfo = find_sheet(xl)
bevEntwSheet, wsheetName, number = sheetInfo
# number = sheetInfo[2]

print(bevEntwSheet)
print(wsheetName)
print(number)

df_dpr, df_raw = subset_dpr(xl, wsheetName, number)
# usage of vectorize function for performance improvement when handling
# big datasets
command = np.vectorize(dpr_calc)
df_dpr['dependency_ratio'] = command(df_dpr.erwerbsfaehig,
                                     df_dpr.nicht_erwerbsfaehig)

# TODO uncomment arcpy.AddMessage(df_dpr)
# # ====> uncomment
# print(df_dpr.info())

#########################################
print(df_dpr.head())
# print(df_raw.head()

# aagr calculation
df_aagr, timeDiff = subset_aagr(xl, bevEntwSheet)
aagr_calc_vect = np.vectorize(aagr_calc)
df_aagr['average_annual_growth_rate'] = aagr_calc_vect(
    df_aagr.iloc[:, 2], df_aagr.iloc[:, 3], timeDiff)
print(df_aagr)
print(type(df_aagr.iloc[:, 2]))

###########################################################################
# NOTE Export Data as table and csv
# BUG if dpr and aagr are both calculated: both DataFrames contain
# columns Kennz and Name.
# Solution: performe join

# df_dpr.to_csv(
#     "~/test_csv.csv", sep=';',
#     encoding='utf-8')  # TODO: adjust so file is exported to gdb folder

exportCSV(df_aagr, df_dpr)

###########################################################################
# NOTE data preparation for table creation & further addition to .gdb
# data Join .shp with table
v = df_dpr.to_numpy()

# create numpy array from pandas dataframe
# important to be able to save the data as table in a gdb

# creates numpy array from pandas DataFrame (df_dpr)
dpr_array = np.array(np.rec.fromrecords(df_dpr.values))
dpr_names = df_dpr.dtypes.index.tolist()
# print(x.dtype)

# saves column names in a tuple
dpr_array.dtype.names = tuple(dpr_names)

print(dpr_array.dtype)
print(dpr_array)
print(type(dpr_array))
# TODO uncomment arcpy.da.NumPyArrayToTable(dpr_array, r'C:\Temp\BevProject\BevProject.gdb\testTable')

aagr_array = np.array(np.rec.fromrecords(df_aagr.values))
aagr_names = df_aagr.dtypes.index.tolist()

aagr_array.dtype.names = tuple(aagr_names)

print(aagr_array)
# This is some random stuff just so no error shows up ...
