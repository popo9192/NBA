import openpyxl
from openpyxl import load_workbook
import os
from datetime import date, time
from datetime import datetime
import pandas as pd
from pandas import ExcelWriter
from functools import reduce



def saveToExcel(df,filename,tab):
    writer = ExcelWriter(filename)
    df.to_excel(writer,tab)
    writer.save()

def getDataSet(dataset):
    df = pd.read_excel(dataset)
    # print('Dataset Loaded')
    return df

def getDataSetcsv(dataset):
    df = pd.read_csv(dataset)
    print('Dataset Loaded')
    return df

def getDateCutoff(year):
    # year = int(year)
    if year == '2017':
        return '2017-10-17'
    if year == '2016':
        return '2016-10-25'
    if year == '2015':
        return '2015-10-27'
    else:
        return('Invalid Year')

def getHomeIndex(df):
    if df['HomeTeam'] == df['TEAM_ABBREVIATION']:
        return (0)
    else:
        return (1)

def getHomeORTG(df):
    if df['HomeIndex'] == 0:
        return df['OFF_RATING']


def getAwayORTG(df):
    if df['HomeIndex'] == 1:
        return df['OFF_RATING']

def resetRest(df):
    x = df['DaysRest']
    x = x.days
    if x > 3:
        x = 3
    return x

def VSresetRest(df):
    x = df['vs_DaysRest']
    x = x.days
    if x > 3:
        x = 3
    return x


def getHomeDRTG(df):
    if df['HomeIndex'] == 0:
        return df['DEF_RATING']


def getAwayDRTG(df):
    if df['HomeIndex'] == 1:
        return df['DEF_RATING']

def getLocationORTG(df):
    if df['HomeIndex'] == 0:
        return df['HomeORTG']
    else:
        return df['AwayORTG']

def getLocationDRTG(df):
    if df['HomeIndex'] == 0:
        return df['HomeDRTG']
    else:
        return df['AwayDRTG']
