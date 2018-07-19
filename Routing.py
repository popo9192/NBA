import openpyxl
from openpyxl import load_workbook
import os
from datetime import date, time
from datetime import datetime
import pandas as pd
from pandas import ExcelWriter,DataFrame
import urllib.request
from requests import get
from Stat_Finder import getAllGames


source = '/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/'
os.chdir('/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/')

BASE_URL = 'http://stats.nba.com/stats/{endpoint}'
HEADERS = {
    'user-agent': ('Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36'),  # noqa: E501
    'Dnt': ('1'),
    'Accept-Encoding': ('gzip, deflate, sdch'),
    'Accept-Language': ('en'),
    'origin': ('http://stats.nba.com')
    }

endpoints = ['boxscoreadvancedv2','boxscorefourfactorsv2','boxscorescoringv2']
# params = "gameid=0041700401&startPeriod=01&endPeriod=01&startrange=01&endrange=01&rangetype=0"
game_id="0041700401"
params={'GameID': game_id,
    'Season': "0",
    'SeasonType': "0",
    'RangeType': "0",
    'StartPeriod': "0",
    'EndPeriod': "0",
    'StartRange': "0",
    'EndRange': "0"}
# ndx is the index in the json. this is for the team stats in boxscoreadvancedv2
ndx=1


def getStat(endpoint, params):
    h = dict(HEADERS)
    _get = get(BASE_URL.format(endpoint=endpoint), params=params,
               headers=h)
    # print (_get.url)
    _get.raise_for_status()
    data = _get.json()
    return data

def _api_scrape(json_inp, ndx):
    """
    Internal method to streamline the getting of data from the json
    Args:
        json_inp (json): json input from our caller
        ndx (int): index where the data is located in the api
    Returns:
        If pandas is present:
            DataFrame (pandas.DataFrame): data set from ndx within the
            API's json
        else:
            A dictionary of both headers and values from the page
    """

    try:
        headers = json_inp['resultSets'][ndx]['headers']
        values = json_inp['resultSets'][ndx]['rowSet']
    except KeyError:
        # This is so ugly but this is what you get when your data comes out
        # in not a standard format
        try:
            headers = json_inp['resultSet'][ndx]['headers']
            values = json_inp['resultSet'][ndx]['rowSet']
        except KeyError:
            # Added for results that only include one set (ex. LeagueLeaders)
            headers = json_inp['resultSet']['headers']
            values = json_inp['resultSet']['rowSet']
    return DataFrame(values, columns=headers)

# json_inp = getStat(endpoint, params)
# df = _api_scrape(json_inp,ndx)

# print(df[["GAME_ID","TEAM_NAME"]])

def hitEndpoints(endpoints,params,ndx):
    df1 = pd.DataFrame()
    for e in endpoints:
        json_inp = getStat(e, params)
        df = _api_scrape(json_inp,ndx)
        df1 = pd.concat([df1,df],axis=1)
    print(df1)

def getADVStats(gameList,endpoints,params,ndx):
    df1 = pd.DataFrame()
    for a in gameList:
        print(a)
        params[game_id] = a
        df = hitEndpoints(endpoints,params,ndx)
        df1 = pd.concat([df1,df],axis=0)
    # df1.fillna(method='ffill',inplace=True)
    # print(df1.head())
    print('Stats Compiled')
    return df1

gameList = getAllGames()
gameList = gameList[:5]
getADVStats(gameList,endpoints,params,ndx)





def getDataSet(dataset):
    df = pd.read_excel(dataset)
    # print('Dataset Loaded')
    return df
