import openpyxl
from openpyxl import load_workbook
import os
from datetime import date
import pandas as pd
from pandas import ExcelWriter
import numpy as np
from base_functions import (saveToExcel,getDataSet,getDataSetcsv,getDateCutoff,LoadModel)
# from Stat_Finder import getGames, getToday,getYesterday,getADVStats,getTodaysGames
from Odds_Scraper import scrapeOdds, main
# os.chdir('C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel')

source = '/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/'
os.chdir('/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/')



# today = getToday()
# yesterday = getYesterday()
# print(today)

# -----------Get Yesterdays Games and Add to existing game Data--------------
def GetYesterdaysData():
    yesterdaysGames = getGames(yesterday)
    df = getADVStats(yesterdaysGames)
    saveToExcel(df,'yesterdaysGames.xlsx','Master')
    df1 = getDataSet('AllStats_' + year + '.xlsx')
    df2 = getDataSet('yesterdaysGames.xlsx')
    df4 =df1.tail(1)
    x = df4['GAMECODE'].str[:9].values
    df5 =df2.head(1)
    y = df5['GAMECODE'].str[:9].values
    df3= df1.append(df2)
    if x != y:
        saveToExcel(df3,'AllStats_' + year + '.xlsx','Master')
    return df3

def getResults(ActualDF,ProjectionsDF,Period):
    dfr = ActualDF[['GAMECODE','TEAM_ABBREVIATION','PTS']]
    dfr['GAMECODE_x'] = dfr['GAMECODE']
    dfr['TEAM_ABBREVIATION_x_x'] = dfr['TEAM_ABBREVIATION']
    # dfr['TEAM_ABBREVIATION_x'] = dfr['TEAM_ABBREVIATION']
    # --------------MAKE TEAMABBR_XX for daily----------------
    dfr = dfr[['GAMECODE_x','TEAM_ABBREVIATION_x_x','PTS']]
    # print(dfr.head(),ProjectionsDF.head())
    df4 = pd.merge(ProjectionsDF, dfr, on=['GAMECODE_x','TEAM_ABBREVIATION_x_x'],how='outer')
    # print(df4.head())
    df4 = df4.dropna()
    df4['ActualSpread'] = df4['PTS'].shift(1) - df4['PTS']
    df4['ActualOU'] = df4['PTS'].shift(-1) + df4['PTS']

    df4['ActualLines'] = df4.apply(getActualLines,axis=1)
    df4['BetType'] = df4.apply(getBetType,axis=1)
    df4['Correct'] = df4.apply(isBetCorrect,axis=1)
    df4 = df4.loc[df4['Correct'] != 'Push']
    print('Results Found')
    # dfs = getDataSet('Season_Results.xlsx')
    # dfs = dfs.append(df4)
    saveToExcel(df4,Period + '_Results.xlsx','Master')
    return df4
    # df5 =dfs.tail(1)
    # x = df5['GAMECODE_x'].str[:9].values
    # df6 =df4.head(1)
    # y = df5['GAMECODE_x'].str[:9].values
    # if x != y:
    # saveToExcel(dfs,'Season_Results.xlsx','Master')



def GetTodaysData():
    todaysGames = getTodaysGames(today)
    game = []
    df = getDataSet('DataForModel_'+ year +'.xlsx')
    # df = df.dropna()
    for i in todaysGames:
        hometeam = i[12:]
        awayteam = i[9:12]
        game.append([hometeam,awayteam])
    print(game)

    dfb = pd.DataFrame(columns=['Match','GAMECODE_x','GAME_DATE_x','TEAM_ABBREVIATION_x_x','HomeIndex_x_x','DaysRest_x','AvgPace_x_x','AvgORTG_x_x','AvgDRTG_x_x',
    'AvgORTG_L5_x_x','AvgDRTG_L5_x_x','std_AvgORTG_x_x', 'std_AvgDRTG_x_x','std_AvgORTG_L5_x_x','std_AvgDRTG_L5_x_x','HomeORTG_x_x',
    'HomeDRTG_x_x','AwayORTG_x_x','AwayDRTG_x_x','Location_Avg_ORTG_x_x','DaysRest_y','AvgPace_x_y','AvgORTG_x_y','AvgDRTG_x_y',
    'AvgORTG_L5_x_y','AvgDRTG_L5_x_y','std_AvgORTG_x_y', 'std_AvgDRTG_x_y','std_AvgORTG_L5_x_y','std_AvgDRTG_L5_x_y','HomeORTG_x_y',
    'HomeDRTG_x_y','AwayORTG_x_y','AwayDRTG_x_y','Location_Avg_ORTG_x_y'])
    for x in game:
        match = x[1] + x[0]

        df1 = df.loc[df['TEAM_ABBREVIATION_x'] == x[0]]
        df2 = df.loc[df['TEAM_ABBREVIATION_x'] == x[1]]
        df1 = df1.tail(1)
        df2 = df2.tail(1)
        df1['Match'] = match
        df2['Match'] = match
        df1['GAME_DATE'] = pd.to_datetime(df1['GAMECODE'].str[:9])
        df2['GAME_DATE'] = pd.to_datetime(df2['GAMECODE'].str[:9])
        df1['today'] = pd.to_datetime(today)
        df2['today'] = pd.to_datetime(today)
        df1['DaysRest'] = (df1['today'] - df1['GAME_DATE']).astype('timedelta64[D]')
        df2['DaysRest'] = (df2['today'] - df2['GAME_DATE']).astype('timedelta64[D]')
        df3 = pd.merge(df1, df2, on='Match',how='outer')
        df4 = pd.merge(df2, df1, on='Match',how='outer')
        date = today.replace("-","")
        df3['date'] = date
        df4['date'] = date
        df3['GAMECODE_x'] = df3.apply(getGameCodeToday,axis=1)
        df4['GAMECODE_x'] = df4.apply(getGameCodeToday,axis=1)
        # print(df4.head())
        df3 = df3[['Match','GAMECODE_x','GAME_DATE_x','TEAM_ABBREVIATION_x_x','HomeIndex_x_x','DaysRest_x','AvgPace_x_x','AvgORTG_x_x','AvgDRTG_x_x',
        'AvgORTG_L5_x_x','AvgDRTG_L5_x_x','std_AvgORTG_x_x', 'std_AvgDRTG_x_x','std_AvgORTG_L5_x_x','std_AvgDRTG_L5_x_x','HomeORTG_x_x',
        'HomeDRTG_x_x','AwayORTG_x_x','AwayDRTG_x_x','Location_Avg_ORTG_x_x','DaysRest_y','AvgPace_x_y','AvgORTG_x_y','AvgDRTG_x_y',
        'AvgORTG_L5_x_y','AvgDRTG_L5_x_y','std_AvgORTG_x_y', 'std_AvgDRTG_x_y','std_AvgORTG_L5_x_y','std_AvgDRTG_L5_x_y','HomeORTG_x_y',
        'HomeDRTG_x_y','AwayORTG_x_y','AwayDRTG_x_y','Location_Avg_ORTG_x_y']]
        df4 = df4[['Match','GAMECODE_x','GAME_DATE_x','TEAM_ABBREVIATION_x_x','HomeIndex_x_x','DaysRest_x','AvgPace_x_x','AvgORTG_x_x','AvgDRTG_x_x',
        'AvgORTG_L5_x_x','AvgDRTG_L5_x_x','std_AvgORTG_x_x', 'std_AvgDRTG_x_x','std_AvgORTG_L5_x_x','std_AvgDRTG_L5_x_x','HomeORTG_x_x',
        'HomeDRTG_x_x','AwayORTG_x_x','AwayDRTG_x_x','Location_Avg_ORTG_x_x','DaysRest_y','AvgPace_x_y','AvgORTG_x_y','AvgDRTG_x_y',
        'AvgORTG_L5_x_y','AvgDRTG_L5_x_y','std_AvgORTG_x_y', 'std_AvgDRTG_x_y','std_AvgORTG_L5_x_y','std_AvgDRTG_L5_x_y','HomeORTG_x_y',
        'HomeDRTG_x_y','AwayORTG_x_y','AwayDRTG_x_y','Location_Avg_ORTG_x_y']]
        dfb = dfb.append(df3)
        dfb = dfb.append(df4)

    dfb = dfb.dropna()

    # dfb = dfb[['TEAM_ABBREVIATION_x_x','HomeIndex_x_x','ProjectedPace','DaysRest_x','DaysRest_y','std_AvgORTG_x_x','HomeORTG_x_x','AwayORTG_x_x','std_AvgORTG_L5_x_x','AvgDRTG_x_y','HomeDRTG_x_y','AwayDRTG_x_y','std_AvgDRTG_x_y']]
    return dfb

def RunModelsOnToday(df,odds):
    home_model = LoadHomeModel()
    away_model = LoadAwayModel()
    # pace_model = LoadPaceModel()
    df['ProjectedPace'] = (df['AvgPace_x_x'] + df['AvgPace_x_y'])/2
    df['TEAM_ABBREVIATION_x'] = df['TEAM_ABBREVIATION_x_x']

    dfh = df.loc[df['HomeIndex_x_x'] == 0]
    dfa = df.loc[df['HomeIndex_x_x'] == 1]
    # ------------------REPLACED HOME/AWAY WITH STD BECAUSE TOO MUCH VARIANCE---------------
    x = dfh[['DaysRest_x','DaysRest_y','std_AvgORTG_x_x','std_AvgORTG_x_x','std_AvgORTG_L5_x_x','AwayDRTG_x_y','std_AvgDRTG_x_y','AvgDRTG_x_y']].values
    y = dfa[['DaysRest_x','DaysRest_y','std_AvgORTG_x_x','std_AvgORTG_x_x','std_AvgORTG_L5_x_x','HomeDRTG_x_y','std_AvgDRTG_x_y','AvgDRTG_x_y']].values
    # x[0] = x[0].total_days
    # print(x)
    homepred = home_model.predict(x)
    awaypred = away_model.predict(y)
    # print(pred)
    dfh1 = pd.DataFrame({'GAMECODE_x':dfh['GAMECODE_x'],'TEAM_ABBREVIATION_x':dfh['TEAM_ABBREVIATION_x'],'Predicted':homepred})
    dfa1 = pd.DataFrame({'GAMECODE_x':dfa['GAMECODE_x'],'TEAM_ABBREVIATION_x':dfa['TEAM_ABBREVIATION_x'],'Predicted':awaypred})
    df1 = dfh1.append(dfa1)

    df2 = pd.merge(df, df1, on=['GAMECODE_x','TEAM_ABBREVIATION_x'],how='outer')
    df2['ProjectedScore'] = (df2['ProjectedPace'] * df2['Predicted'])/100
    df3 = pd.merge(df2, odds, on=['GAMECODE_x','TEAM_ABBREVIATION_x_x'],how='left')
    df3['CalculatedSpread'] = df3['ProjectedScore'].shift(1) - df3['ProjectedScore']
    df3['CalculatedOU'] = df3['ProjectedScore'].shift(-1) + df3['ProjectedScore']
    df3['CalculatedLines'] = df3.apply(getCalcedLines,axis=1)
    df3['TEAM_ABBREVIATION_x'] = df3['TEAM_ABBREVIATION_x_x']
    df3 = df3[['GAMECODE_x','TEAM_ABBREVIATION_x','ProjectedScore','CalculatedLines','VegasLines']]
    df3['Difference'] = df3['CalculatedLines'] - df3['VegasLines']
    df3['BetGrade'] = df3.apply(gradeBet,axis=1)
    # saveToExcel(df3,'Projections.xlsx','Master')
    return df3




def getCalcedLines(df):
    if df['VegasLines'] > 100:
        return df['CalculatedOU']
    else:
        return df['CalculatedSpread']

def getActualLines(df):
    if df['VegasLines'] > 100:
        return df['ActualOU']
    else:
        return df['ActualSpread']

def getBetType(df):
    if df['VegasLines'] > 100:
        return 'OU'
    else:
        return 'Spread'

def gradeBet(df):
    if df['Difference'] >= 10 or df['Difference'] <= -10:
        return 'A'
    elif (df['Difference'] >= 5 and df['Difference'] < 10) or (df['Difference'] <= -5 and df['Difference'] >= -10)  :
        return 'B'
    elif (df['Difference'] > 2 and df['Difference'] < 5) or (df['Difference'] < -2 and df['Difference'] > -5)  :
        return 'C'
    else:
        return 'D'

def isBetCorrect(df):
    if (df['ActualLines'] > df['VegasLines'] and df['CalculatedLines'] > df['VegasLines']) or (df['ActualLines'] < df['VegasLines'] and df['CalculatedLines'] < df['VegasLines']):
        return 'Win'
    if df['ActualLines'] == df['VegasLines']:
        return 'Push'
    else:
        return 'Loss'


def GetOdds():
    df1 = getDataSet('Historical_Odds_'+ '2017' +'.xlsx')
    if fetchOdds:
        scrapeOdds()
        df = getDataSet('Todays_Odds.xlsx')
        df['AwayTeam'] = df.apply(getTeams,axis=1)
        df['HomeTeam'] = df.apply(getOppTeams,axis=1)
        df['key'] = df['key'].astype(str)
        df['GAMECODE_x'] = df.apply(getGameCodeODDS,axis=1)
        df['VegasLines'] = df.apply(filterOdds,axis=1)
        df['TEAM_ABBREVIATION_x_x'] = df['AwayTeam']
        df = df[['GAMECODE_x','TEAM_ABBREVIATION_x_x','VegasLines']]
        saveToExcel(df,'Todays_Odds.xlsx','Master')
        df1 = df1.append(df)
        saveToExcel(df1,'Season_Odds.xlsx','Master')

    # df3 = df1.tail(1)
    # x = df3['GAMECODE_x'].str[:9].values
    # df4 =df.head(1)
    # y = df4['GAMECODE_x'].str[:9].values
    # if x != y:
    #     df2 = df1.append(df)
    #     saveToExcel(df2,'Season_Odds.xlsx','Master')
    return df1

def getAllOdds(year):
    df = getDataSet('DataForModel_'+year+'.xlsx')
    dates = df.GAME_DATE_EST_x_x.unique()
    dfall = getDataSet('Historical_Odds_'+year+'.xlsx')
    done = dfall.GAME_DATE_EST_x.unique()
    for i in dates:
        if i > done[-1]:
            x = i
            i = i[:10]
            i = i.replace("-","")
            print(i)
            df = main(i)
            # df = getDataSet('Todays_Odds.xlsx')
            df['AwayTeam'] = df.apply(getTeams,axis=1)
            df['HomeTeam'] = df.apply(getOppTeams,axis=1)
            df['key'] = df['key'].astype(str)
            df['GAMECODE_x'] = df.apply(getGameCodeODDS,axis=1)
            df['VegasLines'] = df.apply(filterOdds,axis=1)
            df['TEAM_ABBREVIATION_x_x'] = df['AwayTeam']
            df = pd.DataFrame({'GAMECODE_x':df['GAMECODE_x'],'TEAM_ABBREVIATION_x_x':df['TEAM_ABBREVIATION_x_x'],'VegasLines':df['VegasLines'],
            'GAME_DATE_EST_x':x})
            # saveToExcel(df,'Todays_Odds.xlsx','Master')
            # dfh = getDataSet('Historical_Odds.xlsx')
            dfall = dfall.append(df)
        saveToExcel(dfall,'Historical_Odds_'+year+'.xlsx','Master')

def getTeams(df):
    return EditTeamList[df['team']]

def getOppTeams(df):
    return EditTeamList[df['opp_team']]

def filterOdds(df):
    if df['rl_time'] == 'home':
        return (df['tot_PIN_line'])
    else:
        return (df['rl_PIN_line'])

def getGameCodeODDS(df):
    if df['rl_time'] == 'away':
        return(df['key'] + '/'+ df['AwayTeam'] + df['HomeTeam'])
    else:
        return(df['key'] + '/'+ df['HomeTeam'] + df['AwayTeam'])

def getGameCodeToday(df):
    return(df['date'] + '/'+ df['Match'])


def backtest():
    # ----Add in ability so select certain models--------------
    odds = GetOdds()
    df = getDataSet('DataForModel_'+ year +'.xlsx')
    startDate = '2017-10-31T00:00:00'
    df = df.loc[df['GAME_DATE_EST_x'] >= startDate]
    actual = getDataSet('AllStats_' + year + '.xlsx')
    proj = RunModels(df,odds)
    getResults(actual,proj,'ALL')

def liveResults():
    todaysGames = getGames(today)
    df = getADVStats(todaysGames)
    df['GAMECODE_x'] = df['GAMECODE']
    df['TEAM_ABBREVIATION_x'] = df['TEAM_ABBREVIATION']
    actual = df[['GAMECODE_x','TEAM_ABBREVIATION_x','PTS']]
    projected = getDataSet('Projections.xlsx')
    getResults(actual,projected,'Todays')

def getMissingGames(df):
    go = True
    today = getToday()
    today = int(float(today))
    df1 = df.tail(1)
    x = df1['GAMECODE'].str[:9].values
    while go:
        x = int(float(x))
        x = x+1
        if x == today:
            go = False
        x = str(x)
        GetYesterdaysData(x)




EditTeamList = {'Atlan':'ATL','Bost':'BOS','Brookl':'BKN','Charlot':'CHA','Chica':'CHI','Clevela':'CLE','Dall':'DAL','Denv':'DEN','Detro':'DET',
        'Golden Sta':'GSW','Houst':'HOU','India':'IND','L.A. Clippe':'LAC','L.A. Lake':'LAL','Memph':'MEM','Mia':'MIA','Milwauk':'MIL',
        'Minneso':'MIN','New Orlea':'NOP','New Yo':'NYK','Oklahoma Ci':'OKC','Orlan':'ORL','Philadelph':'PHI','Phoen':'PHX','Portla':'POR',
        'Sacramen':'SAC','San Anton':'SAS','Toron':'TOR','Ut':'UTA','Washingt':'WAS'}

# --------------Get Yesterdays Data------------------
# GetYesterdaysData()
# year = '2017'
# CalcStats(year)
# # --------------Get Yesterdays Results------------------
# actual = getDataSet('AllStats_' + year + '.xlsx')
# proj = getDataSet('Projections.xlsx')
# getResults(actual,proj,'Yesterdays')
# --------------Get Todays Games and Run Models------------------
# odds = GetOdds()
# df = GetTodaysData()
#
# proj = RunModelsOnToday(df,odds)
# saveToExcel(proj,'Projections.xlsx','Master')

# -------------------Run Model on Specific Day. Start of backtesting---------------
year = '2017'
fetchOdds = False
# # df = getDataSet('DataForModel_'+ year +'.xlsx')
getAllOdds(year)
# backtest()
