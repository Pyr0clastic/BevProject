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

x = '/DATEN201617_VGL.xlsx'     # TODO delete lines 20 until 30
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
    erwerbsfaehig:              Column in Dataframe containing Erwerbsfaehige
    nicht_erwerbsfaehig:        Column in Dataframe containing nicht Erwerbsfaehige
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
    startYear:      startYear    # TODO Input has to be selectable over toolbox
    endYear:                # TODO Has to be selectable over toolbox
    """
    startYear = int(input("Ausgangszeitpunkt: "))
    endYear = int(input("Endzeitpunkt: "))
    timeDiff = endYear - startYear
    print(
        "Ausgangszeitpunkt: {} \nEndzeitpunkt: {} \nZeitdifferenz: {}".format(
            startYear,
            endYear,
            timeDiff))

    df_bevEntw = xl.parse(bevEntwSheet)


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
    "_________________________________________________________________________"
)
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
df_dpr['dependency_ratio'] = command(
    df_dpr.erwerbsfaehig,
    df_dpr.nicht_erwerbsfaehig)

# TODO unckommen arcpy.AddMessage(df_dpr)
# # ====> uncomment
# print(df_dpr.info())

df_dpr.to_csv(
    "~/test_csv.csv", sep=';',
    encoding='utf-8')  # TODO: adjust so file is exported to gdb folder

v = df_dpr.to_numpy()

# create numpy array from pandas dataframe
# important to be able to save the data as table in a gdb

# creates numpy array from pandas DataFrame (df_dpr)
x = np.array(np.rec.fromrecords(df_dpr.values))
names = df_dpr.dtypes.index.tolist()
# print(x.dtype)

# saves column names in a tuple
x.dtype.names = tuple(names)

# print(x.dtype)
# TODO: uncomment! arcpy.da.NumPyArrayToTable(x, r'C:\Temp\BevProject\BevProject.gdb\testTable')
#########################################
print(df_dpr.head())
# print(df_raw.head()
subset_aagr(xl, bevEntwSheet)
