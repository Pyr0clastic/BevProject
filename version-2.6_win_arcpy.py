import arcpy
#encoding="utf-8"
import numpy as np
import pandas as pd
#import xlsxwriter

import sys
import os
import re

# Displays all Rows in a DataFrame; default is limited to around hundred
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

cwd = os.getcwd()
print(cwd)

# Assign spreadsheet filename to `file`
#file = 'DATEN201617_VGL.xlsx'

filePath = arcpy.GetParameterAsText(0).replace('\\', '\u005C')
arcpy.AddMessage("Imported from : " + filePath)
fileFolder, fileName = os.path.split(filePath)
os.chdir(fileFolder)
xl = pd.ExcelFile(fileName)
# Load spreadsheet
#xl = pd.ExcelFile(file)

# Print the sheet names
print(xl.sheet_names)


def sheet_bevstand(excelFile):
    """
    This function searches for the Sheet "Bevoelkerungsstand" in the provided ExcelFile
    Additionally it checks the survey year

    Parameters:
    -----------
    excelFile : the loaded excel spreasdsheet
    """
    wsheetName_namen = xl.sheet_names
    for wsheetName in wsheetName_namen:
        if "stand" in wsheetName:
            surveyYear = re.findall('\d+', wsheetName)
            for number in surveyYear:
                if len(number) == 4:
                    return wsheetName, number
        else:
            continue


def subset_dpr(xl, wsheetName, year):
    """
    Function erwerbData's purpose is to select just the data needed for the calculation of the parameter dependency ratio (Abhaengigenquote)

    Parameters:
    -----------
    wsheetName:  The exact name of the needed excel worksheet that contains Bevoelkerungsstand
    year:   Survey year
    xl:     loaded excel spreadsheet
    """
    df_raw = xl.parse(wsheetName)
    all = [
        'Wbv2016', 'Wbv2016v00b04', 'Wbv2016v05b09', 'Wbv2016v10b14', 'Wbv2016v15b19',
        'Wbv2016v20b24', 'Wbv2016v25b29', 'Wbv2016v30b34', 'Wbv2016v35b39', 'Wbv2016v40b44',
        'Wbv2016v45b49', 'Wbv2016v50b54', 'Wbv2016v55b59', 'Wbv2016v60b64', 'Wbv2016v65b69',
        'Wbv2016v70b74', 'Wbv2016v75b79', 'Wbv2016v80b84', 'Wbv2016v85+'
    ]
    ##########################
    new = []
    dict = {}
    if year != "2016":
        for i in all:
            new.append(i.replace("2016", year))
            #all.append(i.replace("2016", year))
            q = i.replace("2016", year)
            dict[i] = q
        all.clear()
        all = new
    ########################
    df_erw = df_raw[["Kennz", "Name", all[0]]]
    df_raw.rename(columns=dict, inplace=True)
    jugend = set(all[1:4])
    pensionisten = set(all[14:19])
    erwerbslos = set(all[1:4] + all[14:19])
    erwerbsf = set(all[4:14])
    df_erw = df_raw.iloc[:, 2:5].copy(
    )  # slices based on column index; creates new slick DataFrame with values Kennz, Name, and Wbv
    df_erw["erwerbsfaehig"] = df_raw.drop(
        erwerbsf, axis=1
    ).sum(
        axis=1
    )  # Create sum based on data in set erwerbsf; axis = 1, build sum of each row in the data provided in erwerbsf
    df_erw["nicht_erwerbsfaehig"] = df_raw.drop(
        erwerbslos, axis=1
    ).sum(
        axis=1
    )  # Create sum based on data in set erwerbslos; axis = 1, build sum of each row in the data provided in erwerbslos
    return df_erw  # dict, all


def dpr_calc(erwerbsfaehig, nicht_erwerbsfaehig):
    """
    dpr_calc function calculats dependency ratio (Abhaengigenquote)
    Parameters:
    -----------
    erwerbsfaehig:              Column in Dataframe containing population Erwerbsfaehige
    nicht_erwerbsfaehig:        Column in Dataframe containing population nicht Erwerbsfaehige
    """
    x = (nicht_erwerbsfaehig / erwerbsfaehig) * 100
    return round(x, 2)


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
    "________________________________________________________________________________________________________"
)
wsheetName, number = sheet_bevstand(xl)
df_dpr = subset_dpr(xl, wsheetName, number)
command = np.vectorize(
    dpr_calc)  # usage of vectorize function for performance improvement when handling big datasets
df_dpr['dependency_ratio'] = command(df_dpr.erwerbsfaehig, df_dpr.nicht_erwerbsfaehig)
print(df_dpr)
arcpy.AddMessage(df_dpr)
