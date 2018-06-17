import nba_py
import openpyxl
from openpyxl import load_workbook
import os
from datetime import date, timedelta
import json
import urllib.request
import requests
import pandas as pd
from pandas import ExcelWriter
import codecs
from datetime import datetime, timedelta
from nba_py.constants import CURRENT_SEASON
from nba_py import constants,game
from Format_data import getDataSet, saveToExcel
os.chdir('C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel')

# source = '/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/'
# os.chdir('/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/')


def getToday():
    today = datetime.strftime(datetime.now(), '%Y-%m-%d')
    return today

def getYesterday():
   yesterday = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')
   return yesterday



def getADVStats(gameList):
    df1 = pd.DataFrame()
    for a in gameList:
        # print(a)
        boxscore_summary = game.BoxscoreSummary(a)
        sql_team_basic = boxscore_summary.game_summary()
        sql_team_basic = sql_team_basic[['GAME_DATE_EST','GAMECODE']]

        boxscore_advanced = game.BoxscoreAdvanced(a)
        sql_team_advanced = boxscore_advanced.sql_team_advanced()

        team_four_factors = game.BoxscoreFourFactors(a)
        sql_team_four_factors = team_four_factors.sql_team_four_factors()

        boxscore = game.Boxscore(a)
        sql_team_scoring = boxscore.team_stats()

        boxscore_scoring = game.BoxscoreScoring(a)
        sql_boxscore_scoring = boxscore_scoring.sql_team_scoring()

        boxscore_misc = game.BoxscoreMisc(a)
        sql_team_misc = boxscore_misc.sql_team_misc()

        df = pd.concat([sql_team_basic, sql_team_advanced,sql_team_four_factors,sql_team_scoring,sql_boxscore_scoring,sql_team_misc], axis=1)
        df1 = pd.concat([df1,df],axis=0)
    df1.fillna(method='ffill',inplace=True)
    # print(df1.head())
    print('Stats Compiled')
    return df1

def StatFinder(gameList):
    df1 = pd.DataFrame()
    for a in gameList:
        # print(a)

        boxscore_scoring = game.BoxscoreScoring(a)
        sql_boxscore_scoring = boxscore_scoring.sql_team_scoring()

        boxscore_misc = game.BoxscoreMisc(a)
        sql_team_misc = boxscore_misc.sql_team_misc()

        df = pd.concat([sql_boxscore_scoring, sql_team_misc], axis=1)
        df1 = pd.concat([df1,df],axis=0)
    df1.fillna(method='ffill',inplace=True)
    # print(df1.head())
    print('Stats Compiled')
    return df1

def getAllGames():
    #------------2016-----------
    # url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2016/league/00_full_schedule.json'
    #------------2015-----------
    url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2015/league/00_full_schedule.json'

    response = urllib.request.urlopen(url)
    reader = codecs.getreader("utf-8")
    data = json.load(reader(response))
    gameIDs = []
    months = [0,1,2,3,4,5,6,7,8]
    for x in months:
        print(x)
        games = (data['lscd'][x]['mscd']['g'])
        for i in range(len(games)):
            gameIDs.append(games[i]['gid'])
    print('Games Aquired')
    return gameIDs



def getGames(date):
    #------------2017-----------
    url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2017/league/00_full_schedule.json'

    response = urllib.request.urlopen(url)
    reader = codecs.getreader("utf-8")
    data = json.load(reader(response))
    gameIDs = []
    months = [0,1,2,3,4,5,6,7]
    for x in months:
        games = (data['lscd'][x]['mscd']['g'])
        for i in range(len(games)):
            if games[i]['gdte'] == date:
                gameIDs.append(games[i]['gid'])
    print('Games Aquired')
    return gameIDs

def getTodaysGames(date):
    #------------2017-----------
    url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2017/league/00_full_schedule.json'

    response = urllib.request.urlopen(url)
    reader = codecs.getreader("utf-8")
    data = json.load(reader(response))
    gameIDs = []
    months = [0,1,2,3,4,5,6,7]
    for x in months:
        games = (data['lscd'][x]['mscd']['g'])
        for i in range(len(games)):
            if games[i]['gdte'] == date:
                gameIDs.append(games[i]['gcode'])
    print('Games Aquired')
    return gameIDs

def getGamesTilNow(last,date):
    #------------2017-----------
    url = 'http://data.nba.com/data/10s/v2015/json/mobile_teams/nba/2017/league/00_full_schedule.json'

    response = urllib.request.urlopen(url)
    reader = codecs.getreader("utf-8")
    data = json.load(reader(response))
    gameIDs = []
    months = [0,1,2,3,4,5,6,7]
    for x in months:
        games = (data['lscd'][x]['mscd']['g'])
        for i in range(len(games)):
            if games[i]['gdte'] < date and games[i]['gdte'] > last and games[i]['gid'] != '0021700924':
                gameIDs.append(games[i]['gid'])
    print('Games Aquired')
    return gameIDs

# today = getToday()
# yesterday = getYesterday()
# print(today)
# #
# todaysGames = getGames(today)
# yesterdaysGames = getGames(yesterday)
def getMissingGames():
    year = '2017'
    today = getToday()
    df = getDataSet('AllStats_' + year + '.xlsx')
    df1 = df.tail(1)
    dfh = df.head(1)
    last = df1['GAME_DATE_EST'].str[:10].values
    first = dfh['GAME_DATE_EST'].str[:10].values
    # print(last[0])
    GamesSoFar = getGamesTilNow(last,today)
    df2 = getADVStats(GamesSoFar)
    saveToExcel(df2,'MissingGames_' + year + '.xlsx',year)
    frames = [df,df2]
    df3= pd.concat(frames)
    saveToExcel(df3,'AllStats_' + year + '.xlsx',year)

def squish():
    df2 = getDataSet('MissingGames_' + year + '_1' +'.xlsx')
    df1 = getDataSet('MissingGames_' + year + '.xlsx')
    df3 = df1.append(df2)
    saveToExcel(df3,'AllStats_' + year + '.xlsx',year)

# print(todaysGames,yesterdaysGames)
# # print(GamesSoFar)
#
# allGames = getAllGames()
# getADVStats(GamesSoFar)
# getMissingGames()
year = '2017'
# df = StatFinder(['0021700923'])
# saveToExcel(df,'Endpoint_test.xlsx',year)
# squish()
