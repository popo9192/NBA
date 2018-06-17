import openpyxl
from openpyxl import load_workbook
import os
from datetime import date, time
from datetime import datetime
import pandas as pd
from pandas import ExcelWriter
from functools import reduce
from base_functions import (saveToExcel,getDataSet,getDataSetcsv,getDateCutoff,getHomeIndex,getHomeORTG,getAwayORTG,resetRest,VSresetRest,getHomeDRTG,
getAwayDRTG,getLocationORTG,getLocationDRTG)
os.chdir('C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel')
#
# source = '/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/'
# os.chdir('/Users/peterhaley/Desktop/NBA_Wizard/NBA/Excel/')

teamList = ['ATL','BOS','BKN','CHA','CHI','CLE','DAL','DEN','DET','GSW','HOU',
         'IND','LAC','LAL','MEM','MIA','MIL','MIN','NOP','NYK','OKC','ORL',
         'PHI','PHX','POR','SAC','SAS','TOR','UTA','WAS']

def SplitTeams(df,year):
    writer = ExcelWriter('AllStats_' + year + '_Split.xlsx')
    cutoff = getDateCutoff(year)

    df['GAME_DATE'] = pd.to_datetime(df['GAME_DATE_EST'])
    df = df.loc[df['GAME_DATE'] >= cutoff]

    df['HomeTeam'] = df['GAMECODE'].str[12:]
    df['AwayTeam'] = df['GAMECODE'].str[9:12]
    df['HomeIndex'] = df.apply(getHomeIndex,axis=1)

    df['HomeORTG'] = df.apply(getHomeORTG,axis=1)
    df['AwayORTG'] = df.apply(getAwayORTG,axis=1)
    df['HomeDRTG'] = df.apply(getHomeDRTG,axis=1)
    df['AwayDRTG'] = df.apply(getAwayDRTG,axis=1)

    df['Location_Avg_ORTG'] = df.apply(getLocationORTG,axis=1)
    df['Location_Avg_DRTG'] = df.apply(getLocationDRTG,axis=1)
    df['FT_RATE'] = df['FTM']/df['FGA']

    dfH = df[(df['HomeIndex'] == 0)]
    dfA = df[(df['HomeIndex'] == 1)]
    df3 = pd.merge(dfH, dfA, on='GAMECODE',how='outer')
    df4 = pd.merge(dfA, dfH, on='GAMECODE',how='outer')
    dfList = [df3,df4]
    df= pd.concat(dfList)
    # print(df3.head())
    df = df.sort_values(by=['GAMECODE','HomeIndex_x'],ascending=[True,True])

    for i in teamList:
        df1 = df.loc[df['TEAM_ABBREVIATION_x'] == i]
        #Format Date and calc rest

        df1['DaysRest'] = df1['GAME_DATE_x'] - df1['GAME_DATE_x'].shift(1)
        df1['DaysRest'] = df1.apply(resetRest,axis=1)

        df1['AvgORTG'] = df1['OFF_RATING_x'].expanding().mean()
        df1['AvgDRTG'] = df1['DEF_RATING_x'].expanding().mean()
        df1['AvgNET'] = df1['NET_RATING_x'].expanding().mean()


        df1['HomeORTG'] = df1['HomeORTG_x'].expanding().mean()
        df1['AwayORTG'] = df1['AwayORTG_x'].expanding().mean()
        df1['HomeDRTG'] = df1['HomeDRTG_x'].expanding().mean()
        df1['AwayDRTG'] = df1['AwayDRTG_x'].expanding().mean()

        df1['AvgPace']= df1['PACE_x'].expanding().mean()
        df1['Possessions'] = (df1['FGA_x'] +(.44 * df1['FTA_x'])-df1['OREB_x'] + df1['TO_x'])
        df1['Actual_Possessions'] = (df1['PTS_x'] / df1['OFF_RATING_x']) * 100
        df1['Avg_Possessions'] = df1['Actual_Possessions'].expanding().mean()
        df1['est_avg_Poss'] = df1['Possessions'].expanding().mean()

        df1['avg_AST'] = df1['AST_x'].expanding().mean()
        df1['avg_AST_PCT']= df1['AST_PCT_x'].expanding().mean()
        df1['avg_AST_TOV'] = df1['AST_TOV_x'].expanding().mean()
        df1['avg_AST_RATIO'] = df1['AST_RATIO_x'].expanding().mean()

        df1['avg_EFG%'] = df1['EFG_PCT_x'].expanding().mean()
        df1['avg_TS'] = df1['TS_PCT_x'].expanding().mean()
        df1['avg_FGA'] = df1['FGA_x'].expanding().mean()
        df1['avg_FG%'] = df1['FG_PCT_x'].expanding().mean()
        df1['avg_FG3A'] = df1['FG3A_x'].expanding().mean()
        df1['avg_FG3%'] = df1['FG3_PCT_x'].expanding().mean()
        df1['avg_FTA'] = df1['FTA_x'].expanding().mean()
        df1['avg_FT%'] = df1['FT_PCT_x'].expanding().mean()
        df1['avg_FTA_RATE'] = df1['FTA_RATE_x'].expanding().mean()
        df1['avg_FT_RATE'] = df1['FT_RATE_x'].expanding().mean()


        df1['avg_PIE'] = df1['PIE_x'].expanding().mean()

        df1['avg_Fouls'] = df1['PF_x'].expanding().mean()
        df1['avg_Fouls_Drawn'] = df1['PFD_x'].expanding().mean()

        df1['avg_Steals'] = df1['STL_x'].expanding().mean()
        df1['avg_BLK'] = df1['BLK_x'].expanding().mean()
        df1['avg_times_blocked'] = df1['BLKA_x'].expanding().mean()

        df1['avg_TO'] = df1['TO_x'].expanding().mean()
        df1['avg_TO%'] = df1['TM_TOV_PCT_x'].expanding().mean()
        df1['avg_OREB'] = df1['OREB_x'].expanding().mean()
        df1['avg_OREB%'] = df1['OREB_PCT_x'].expanding().mean()
        df1['avg_DREB'] = df1['DREB_x'].expanding().mean()
        df1['avg_DREB%'] = df1['DREB_PCT_x'].expanding().mean()
        df1['avg_REB'] = df1['REB_x'].expanding().mean()
        df1['avg_REB%'] = df1['REB_PCT_x'].expanding().mean()

        df1['avg_OPP_EFG%'] = df1['OPP_EFG_PCT_x'].expanding().mean()
        df1['avg_OPP_FTA_RATE'] = df1['OPP_FTA_RATE_x'].expanding().mean()
        df1['avg_OPP_TOV%'] = df1['OPP_TOV_PCT_x'].expanding().mean()
        df1['avg_OPP_OREB%'] = df1['OPP_OREB_PCT_x'].expanding().mean()

        df1['avg_PCT_FGA_2PT'] = df1['PCT_FGA_2PT_x'].expanding().mean()
        df1['avg_PCT_FGA_3PT'] = df1['PCT_FGA_3PT_x'].expanding().mean()
        df1['avg_PCT_PTS_2PT'] = df1['PCT_PTS_2PT_x'].expanding().mean()
        df1['avg_PCT_PTS_3PT'] = df1['PCT_PTS_3PT_x'].expanding().mean()
        df1['avg_PCT_PTS_MR'] = df1['PCT_PTS_2PT_MR_x'].expanding().mean()
        df1['avg_PCT_PTS_FB'] = df1['PCT_PTS_FB_x'].expanding().mean()
        df1['avg_PCT_PTS_FT'] = df1['PCT_PTS_FT_x'].expanding().mean()
        df1['avg_PCT_PTS_OFF_TOV'] = df1['PCT_PTS_OFF_TOV_x'].expanding().mean()
        df1['avg_PCT_PTS_PAINT'] = df1['PCT_PTS_PAINT_x'].expanding().mean()
        df1['avg_PCT_AST_2PM'] = df1['PCT_AST_2PM_x'].expanding().mean()
        df1['avg_PCT_UAST_2PM'] = df1['PCT_UAST_2PM_x'].expanding().mean()
        df1['avg_PCT_AST_3PM'] = df1['PCT_AST_3PM_x'].expanding().mean()
        df1['avg_PCT_UAST_3PM'] = df1['PCT_UAST_3PM_x'].expanding().mean()
        df1['avg_PCT_AST_FGM'] = df1['PCT_AST_FGM_x'].expanding().mean()
        df1['avg_PCT_UAST_FGM'] = df1['PCT_UAST_FGM_x'].expanding().mean()
        df1['avg_PTS_OFF_TOV'] = df1['PTS_OFF_TOV_x'].expanding().mean()
        df1['avg_PTS_2ND_CHANCE'] = df1['PTS_2ND_CHANCE_x'].expanding().mean()
        df1['avg_PTS_FB'] = df1['PTS_FB_x'].expanding().mean()
        df1['avg_PTS_PAINT'] = df1['PTS_PAINT_x'].expanding().mean()
        df1['avg_OPP_PTS_OFF_TOV'] = df1['OPP_PTS_OFF_TOV_x'].expanding().mean()
        df1['avg_OPP_PTS_2ND_CHANCE'] = df1['OPP_PTS_2ND_CHANCE_x'].expanding().mean()
        df1['avg_OPP_PTS_FB'] = df1['OPP_PTS_FB_x'].expanding().mean()
        df1['avg_OPP_PTS_PAINT'] = df1['OPP_PTS_PAINT_x'].expanding().mean()

        df1['AvgORTG'] = df1['AvgORTG'].shift(1)
        df1['AvgDRTG'] = df1['AvgDRTG'].shift(1)
        df1['AvgNET'] = df1['AvgNET'].shift(1)

        df1['HomeORTG'] = df1['HomeORTG'].shift(1)
        df1['AwayORTG'] = df1['AwayORTG'].shift(1)
        df1['HomeDRTG'] = df1['HomeDRTG'].shift(1)
        df1['AwayDRTG'] = df1['AwayDRTG'].shift(1)
        df1['Location_Avg_ORTG'] = df1['Location_Avg_ORTG_x'].shift(1)
        df1['Location_Avg_DRTG'] = df1['Location_Avg_DRTG_x'].shift(1)

        df1['AvgPace']= df1['AvgPace'].shift(1)
        df1['Avg_Possessions'] = df1['Avg_Possessions'].shift(1)
        df1['est_avg_Poss'] = df1['est_avg_Poss'].shift(1)

        df1['avg_AST'] = df1['avg_AST'].shift(1)
        df1['avg_AST_PCT']= df1['avg_AST_PCT'].shift(1)
        df1['avg_AST_TOV'] = df1['avg_AST_TOV'].shift(1)
        df1['avg_AST_RATIO'] = df1['avg_AST_RATIO'].shift(1)

        df1['avg_EFG%'] = df1['avg_EFG%'].shift(1)
        df1['avg_TS'] = df1['avg_TS'].shift(1)
        df1['avg_FGA'] = df1['avg_FGA'].shift(1)
        df1['avg_FG%'] = df1['avg_FG%'].shift(1)
        df1['avg_FG3A'] = df1['avg_FG3A'].shift(1)
        df1['avg_FG3%'] = df1['avg_FG3%'].shift(1)
        df1['avg_FTA'] = df1['avg_FTA'].shift(1)
        df1['avg_FT%'] = df1['avg_FT%'].shift(1)
        df1['avg_FTA_RATE'] = df1['avg_FTA_RATE'].shift(1)
        df1['avg_FT_RATE'] = df1['FT_RATE_x'].shift(1)

        df1['avg_PIE'] = df1['avg_PIE'].shift(1)

        df1['avg_Fouls'] = df1['avg_Fouls'].shift(1)
        df1['avg_Fouls_Drawn'] = df1['avg_Fouls_Drawn'].shift(1)

        df1['avg_Steals'] = df1['avg_Steals'].shift(1)
        df1['avg_BLK'] = df1['avg_BLK'].shift(1)
        df1['avg_times_blocked'] = df1['avg_times_blocked'].shift(1)

        df1['avg_TO'] = df1['avg_TO'].shift(1)
        df1['avg_TO%'] = df1['avg_TO%'].shift(1)
        df1['avg_OREB'] = df1['avg_OREB'].shift(1)
        df1['avg_OREB%'] = df1['avg_OREB%'].shift(1)
        df1['avg_DREB'] = df1['avg_DREB'].shift(1)
        df1['avg_DREB%'] = df1['avg_DREB%'].shift(1)
        df1['avg_REB'] = df1['avg_REB'].shift(1)
        df1['avg_REB%'] = df1['avg_REB%'].shift(1)

        df1['avg_OPP_EFG%'] = df1['avg_OPP_EFG%'].shift(1)
        df1['avg_OPP_FTA_RATE'] = df1['avg_OPP_FTA_RATE'].shift(1)
        df1['avg_OPP_TOV%'] = df1['avg_OPP_TOV%'].shift(1)
        df1['avg_OPP_OREB%'] = df1['avg_OPP_OREB%'].shift(1)

        df1['avg_PCT_FGA_2PT'] = df1['avg_PCT_FGA_2PT'].shift(1)
        df1['avg_PCT_FGA_3PT'] = df1['avg_PCT_FGA_3PT'].shift(1)
        df1['avg_PCT_PTS_2PT'] = df1['avg_PCT_PTS_2PT'].shift(1)
        df1['avg_PCT_PTS_3PT'] = df1['avg_PCT_PTS_3PT'].shift(1)
        df1['avg_PCT_PTS_MR'] = df1['avg_PCT_PTS_MR'].shift(1)
        df1['avg_PCT_PTS_FB'] = df1['avg_PCT_PTS_FB'].shift(1)
        df1['avg_PCT_PTS_FT'] = df1['avg_PCT_PTS_FT'].shift(1)
        df1['avg_PCT_PTS_OFF_TOV'] = df1['avg_PCT_PTS_OFF_TOV'].shift(1)
        df1['avg_PCT_PTS_PAINT'] = df1['avg_PCT_PTS_PAINT'].shift(1)
        df1['avg_PCT_AST_2PM'] = df1['avg_PCT_AST_2PM'].shift(1)
        df1['avg_PCT_UAST_2PM'] = df1['avg_PCT_UAST_2PM'].shift(1)
        df1['avg_PCT_AST_3PM'] = df1['avg_PCT_AST_3PM'].shift(1)
        df1['avg_PCT_UAST_3PM'] = df1['avg_PCT_UAST_3PM'].shift(1)
        df1['avg_PCT_AST_FGM'] = df1['avg_PCT_AST_FGM'].shift(1)
        df1['avg_PCT_UAST_FGM'] = df1['avg_PCT_UAST_FGM'].shift(1)
        df1['avg_PTS_OFF_TOV'] = df1['avg_PTS_OFF_TOV'].shift(1)
        df1['avg_PTS_2ND_CHANCE'] = df1['avg_PTS_2ND_CHANCE'].shift(1)
        df1['avg_PTS_FB'] = df1['avg_PTS_FB'].shift(1)
        df1['avg_PTS_PAINT'] = df1['avg_PTS_PAINT'].shift(1)
        df1['avg_OPP_PTS_OFF_TOV'] = df1['avg_OPP_PTS_OFF_TOV'].shift(1)
        df1['avg_OPP_PTS_2ND_CHANCE'] = df1['avg_OPP_PTS_2ND_CHANCE'].shift(1)
        df1['avg_OPP_PTS_FB'] = df1['avg_OPP_PTS_FB'].shift(1)
        df1['avg_OPP_PTS_PAINT'] = df1['avg_OPP_PTS_PAINT'].shift(1)

        # THIS IS FOR THE Opponent

        df1['vs_DaysRest'] = df1['GAME_DATE_y'] - df1['GAME_DATE_y'].shift(1)
        df1['vs_DaysRest'] = df1.apply(VSresetRest,axis=1)

        df1['vs_AvgORTG'] = df1['OFF_RATING_y'].expanding().mean()
        df1['vs_AvgDRTG'] = df1['DEF_RATING_y'].expanding().mean()
        df1['vs_AvgNET'] = df1['NET_RATING_y'].expanding().mean()


        df1['vs_HomeORTG'] = df1['HomeORTG_y'].expanding().mean()
        df1['vs_AwayORTG'] = df1['AwayORTG_y'].expanding().mean()
        df1['vs_HomeDRTG'] = df1['HomeDRTG_y'].expanding().mean()
        df1['vs_AwayDRTG'] = df1['AwayDRTG_y'].expanding().mean()

        df1['vs_AvgPace']= df1['PACE_y'].expanding().mean()
        df1['vs_Possessions'] = (df1['FGA_y'] +(.44 * df1['FTA_y'])-df1['OREB_y'] + df1['TO_y'])
        df1['vs_Actual_Possessions'] = (df1['PTS_y'] / df1['OFF_RATING_y']) * 100
        df1['vs_Avg_Possessions'] = df1['vs_Actual_Possessions'].expanding().mean()
        df1['vs_est_avg_Poss'] = df1['vs_Possessions'].expanding().mean()

        df1['vs_avg_AST'] = df1['AST_y'].expanding().mean()
        df1['vs_avg_AST_PCT']= df1['AST_PCT_y'].expanding().mean()
        df1['vs_avg_AST_TOV'] = df1['AST_TOV_y'].expanding().mean()
        df1['vs_avg_AST_RATIO'] = df1['AST_RATIO_y'].expanding().mean()

        df1['vs_avg_EFG%'] = df1['EFG_PCT_y'].expanding().mean()
        df1['vs_avg_TS'] = df1['TS_PCT_y'].expanding().mean()
        df1['vs_avg_FGA'] = df1['FGA_y'].expanding().mean()
        df1['vs_avg_FG%'] = df1['FG_PCT_y'].expanding().mean()
        df1['vs_avg_FG3A'] = df1['FG3A_y'].expanding().mean()
        df1['vs_avg_FG3%'] = df1['FG3_PCT_y'].expanding().mean()
        df1['vs_avg_FTA'] = df1['FTA_y'].expanding().mean()
        df1['vs_avg_FT%'] = df1['FT_PCT_y'].expanding().mean()
        df1['vs_avg_FTA_RATE'] = df1['FTA_RATE_y'].expanding().mean()
        df1['vs_avg_FT_RATE'] = df1['FT_RATE_y'].expanding().mean()

        df1['vs_avg_PIE'] = df1['PIE_y'].expanding().mean()

        df1['vs_avg_Fouls'] = df1['PF_y'].expanding().mean()
        df1['vs_avg_Fouls_Drawn'] = df1['PFD_y'].expanding().mean()

        df1['vs_avg_Steals'] = df1['STL_y'].expanding().mean()
        df1['vs_avg_BLK'] = df1['BLK_y'].expanding().mean()
        df1['vs_avg_times_blocked'] = df1['BLKA_y'].expanding().mean()

        df1['vs_avg_TO'] = df1['TO_y'].expanding().mean()
        df1['vs_avg_TO%'] = df1['TM_TOV_PCT_y'].expanding().mean()
        df1['vs_avg_OREB'] = df1['OREB_y'].expanding().mean()
        df1['vs_avg_OREB%'] = df1['OREB_PCT_y'].expanding().mean()
        df1['vs_avg_DREB'] = df1['DREB_y'].expanding().mean()
        df1['vs_avg_DREB%'] = df1['DREB_PCT_y'].expanding().mean()
        df1['vs_avg_REB'] = df1['REB_y'].expanding().mean()
        df1['vs_avg_REB%'] = df1['REB_PCT_y'].expanding().mean()

        df1['vs_avg_OPP_EFG%'] = df1['OPP_EFG_PCT_y'].expanding().mean()
        df1['vs_avg_OPP_FTA_RATE'] = df1['OPP_FTA_RATE_y'].expanding().mean()
        df1['vs_avg_OPP_TOV%'] = df1['OPP_TOV_PCT_y'].expanding().mean()
        df1['vs_avg_OPP_OREB%'] = df1['OPP_OREB_PCT_y'].expanding().mean()

        df1['vs_avg_PCT_FGA_2PT'] = df1['PCT_FGA_2PT_y'].expanding().mean()
        df1['vs_avg_PCT_FGA_3PT'] = df1['PCT_FGA_3PT_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_2PT'] = df1['PCT_PTS_2PT_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_3PT'] = df1['PCT_PTS_3PT_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_MR'] = df1['PCT_PTS_2PT_MR_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_FB'] = df1['PCT_PTS_FB_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_FT'] = df1['PCT_PTS_FT_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_OFF_TOV'] = df1['PCT_PTS_OFF_TOV_y'].expanding().mean()
        df1['vs_avg_PCT_PTS_PAINT'] = df1['PCT_PTS_PAINT_y'].expanding().mean()
        df1['vs_avg_PCT_AST_2PM'] = df1['PCT_AST_2PM_y'].expanding().mean()
        df1['vs_avg_PCT_UAST_2PM'] = df1['PCT_UAST_2PM_y'].expanding().mean()
        df1['vs_avg_PCT_AST_3PM'] = df1['PCT_AST_3PM_y'].expanding().mean()
        df1['vs_avg_PCT_UAST_3PM'] = df1['PCT_UAST_3PM_y'].expanding().mean()
        df1['vs_avg_PCT_AST_FGM'] = df1['PCT_AST_FGM_y'].expanding().mean()
        df1['vs_avg_PCT_UAST_FGM'] = df1['PCT_UAST_FGM_y'].expanding().mean()
        df1['vs_avg_PTS_OFF_TOV'] = df1['PTS_OFF_TOV_y'].expanding().mean()
        df1['vs_avg_PTS_2ND_CHANCE'] = df1['PTS_2ND_CHANCE_y'].expanding().mean()
        df1['vs_avg_PTS_FB'] = df1['PTS_FB_y'].expanding().mean()
        df1['vs_avg_PTS_PAINT'] = df1['PTS_PAINT_y'].expanding().mean()
        df1['vs_avg_OPP_PTS_OFF_TOV'] = df1['OPP_PTS_OFF_TOV_y'].expanding().mean()
        df1['vs_avg_OPP_PTS_2ND_CHANCE'] = df1['OPP_PTS_2ND_CHANCE_y'].expanding().mean()
        df1['vs_avg_OPP_PTS_FB'] = df1['OPP_PTS_FB_y'].expanding().mean()
        df1['vs_avg_OPP_PTS_PAINT'] = df1['OPP_PTS_PAINT_y'].expanding().mean()

        df1['vs_AvgORTG'] = df1['vs_AvgORTG'].shift(1)
        df1['vs_AvgDRTG'] = df1['vs_AvgDRTG'].shift(1)
        df1['vs_AvgNET'] = df1['vs_AvgNET'].shift(1)

        df1['vs_HomeORTG'] = df1['vs_HomeORTG'].shift(1)
        df1['vs_AwayORTG'] = df1['vs_AwayORTG'].shift(1)
        df1['vs_HomeDRTG'] = df1['vs_HomeDRTG'].shift(1)
        df1['vs_AwayDRTG'] = df1['vs_AwayDRTG'].shift(1)
        df1['vs_Location_Avg_ORTG'] = df1['Location_Avg_ORTG_y'].shift(1)
        df1['vs_Location_Avg_DRTG'] = df1['Location_Avg_DRTG_y'].shift(1)

        df1['vs_AvgPace']= df1['vs_AvgPace'].shift(1)
        df1['vs_Avg_Possessions'] = df1['vs_Avg_Possessions'].shift(1)
        df1['vs_est_avg_Poss'] = df1['vs_est_avg_Poss'].shift(1)

        df1['vs_avg_AST'] = df1['vs_avg_AST'].shift(1)
        df1['vs_avg_AST_PCT']= df1['vs_avg_AST_PCT'].shift(1)
        df1['vs_avg_AST_TOV'] = df1['vs_avg_AST_TOV'].shift(1)
        df1['vs_avg_AST_RATIO'] = df1['vs_avg_AST_RATIO'].shift(1)

        df1['vs_avg_EFG%'] = df1['vs_avg_EFG%'].shift(1)
        df1['vs_avg_TS'] = df1['vs_avg_TS'].shift(1)
        df1['vs_avg_FGA'] = df1['vs_avg_FGA'].shift(1)
        df1['vs_avg_FG%'] = df1['vs_avg_FG%'].shift(1)
        df1['vs_avg_FG3A'] = df1['vs_avg_FG3A'].shift(1)
        df1['vs_avg_FG3%'] = df1['vs_avg_FG3%'].shift(1)
        df1['vs_avg_FTA'] = df1['vs_avg_FTA'].shift(1)
        df1['vs_avg_FT%'] = df1['vs_avg_FT%'].shift(1)
        df1['vs_avg_FTA_RATE'] = df1['vs_avg_FTA_RATE'].shift(1)
        df1['vs_avg_FT_RATE'] = df1['FT_RATE_y'].shift(1)

        df1['vs_avg_PIE'] = df1['vs_avg_PIE'].shift(1)

        df1['vs_avg_Fouls'] = df1['vs_avg_Fouls'].shift(1)
        df1['vs_avg_Fouls_Drawn'] = df1['vs_avg_Fouls_Drawn'].shift(1)

        df1['vs_avg_Steals'] = df1['vs_avg_Steals'].shift(1)
        df1['vs_avg_BLK'] = df1['vs_avg_BLK'].shift(1)
        df1['vs_avg_times_blocked'] = df1['vs_avg_times_blocked'].shift(1)

        df1['vs_avg_TO'] = df1['vs_avg_TO'].shift(1)
        df1['vs_avg_TO%'] = df1['vs_avg_TO%'].shift(1)
        df1['vs_avg_OREB'] = df1['vs_avg_OREB'].shift(1)
        df1['vs_avg_OREB%'] = df1['vs_avg_OREB%'].shift(1)
        df1['vs_avg_DREB'] = df1['vs_avg_DREB'].shift(1)
        df1['vs_avg_DREB%'] = df1['vs_avg_DREB%'].shift(1)
        df1['vs_avg_REB'] = df1['vs_avg_REB'].shift(1)
        df1['vs_avg_REB%'] = df1['vs_avg_REB%'].shift(1)

        df1['vs_avg_OPP_EFG%'] = df1['vs_avg_OPP_EFG%'].shift(1)
        df1['vs_avg_OPP_FTA_RATE'] = df1['vs_avg_OPP_FTA_RATE'].shift(1)
        df1['vs_avg_OPP_TOV%'] = df1['vs_avg_OPP_TOV%'].shift(1)
        df1['vs_avg_OPP_OREB%'] = df1['vs_avg_OPP_OREB%'].shift(1)

        df1['vs_avg_PCT_FGA_2PT'] = df1['vs_avg_PCT_FGA_2PT'].shift(1)
        df1['vs_avg_PCT_FGA_3PT'] = df1['vs_avg_PCT_FGA_3PT'].shift(1)
        df1['vs_avg_PCT_PTS_2PT'] = df1['vs_avg_PCT_PTS_2PT'].shift(1)
        df1['vs_avg_PCT_PTS_3PT'] = df1['vs_avg_PCT_PTS_3PT'].shift(1)
        df1['vs_avg_PCT_PTS_MR'] = df1['vs_avg_PCT_PTS_MR'].shift(1)
        df1['vs_avg_PCT_PTS_FB'] = df1['vs_avg_PCT_PTS_FB'].shift(1)
        df1['vs_avg_PCT_PTS_FT'] = df1['vs_avg_PCT_PTS_FT'].shift(1)
        df1['vs_avg_PCT_PTS_OFF_TOV'] = df1['vs_avg_PCT_PTS_OFF_TOV'].shift(1)
        df1['vs_avg_PCT_PTS_PAINT'] = df1['vs_avg_PCT_PTS_PAINT'].shift(1)
        df1['vs_avg_PCT_AST_2PM'] = df1['vs_avg_PCT_AST_2PM'].shift(1)
        df1['vs_avg_PCT_UAST_2PM'] = df1['vs_avg_PCT_UAST_2PM'].shift(1)
        df1['vs_avg_PCT_AST_3PM'] = df1['vs_avg_PCT_AST_3PM'].shift(1)
        df1['vs_avg_PCT_UAST_3PM'] = df1['vs_avg_PCT_UAST_3PM'].shift(1)
        df1['vs_avg_PCT_AST_FGM'] = df1['vs_avg_PCT_AST_FGM'].shift(1)
        df1['vs_avg_PCT_UAST_FGM'] = df1['vs_avg_PCT_UAST_FGM'].shift(1)
        df1['vs_avg_PTS_OFF_TOV'] = df1['vs_avg_PTS_OFF_TOV'].shift(1)
        df1['vs_avg_PTS_2ND_CHANCE'] = df1['vs_avg_PTS_2ND_CHANCE'].shift(1)
        df1['vs_avg_PTS_FB'] = df1['vs_avg_PTS_FB'].shift(1)
        df1['vs_avg_PTS_PAINT'] = df1['vs_avg_PTS_PAINT'].shift(1)
        df1['vs_avg_OPP_PTS_OFF_TOV'] = df1['vs_avg_OPP_PTS_OFF_TOV'].shift(1)
        df1['vs_avg_OPP_PTS_2ND_CHANCE'] = df1['vs_avg_OPP_PTS_2ND_CHANCE'].shift(1)
        df1['vs_avg_OPP_PTS_FB'] = df1['vs_avg_OPP_PTS_FB'].shift(1)
        df1['vs_avg_OPP_PTS_PAINT'] = df1['vs_avg_OPP_PTS_PAINT'].shift(1)

        df1.to_excel(writer,i)
    print('Data Split')


def getFirstSplit(year):
    fileName = ('AllStats_'+ year +'_Split.xlsx')
    wb = load_workbook(fileName)
    df = pd.read_excel(fileName, sheetname='ATL')
    tabs = wb.get_sheet_names()
    print("Extracting Team Data")
    for j in tabs:
        if j != 'ATL':
            df4 = pd.read_excel(fileName, sheetname=j)
            frames = [df,df4]
            df= pd.concat(frames)
    dfH = df[(df['HomeIndex_x'] == 0)]
    dfA = df[(df['HomeIndex_x'] == 1)]
    df1 = pd.merge(dfH, dfA, on='GAMECODE',how='outer')
    df2 = pd.merge(dfA, dfH, on='GAMECODE',how='outer')
    dfList = [df1,df2]
    df3= pd.concat(dfList)
    # print(df3.head())
    df3 = df3.sort_values(by=['GAMECODE','HomeIndex_x_x'],ascending=[True,True])
    # saveToExcel(df3,'Formatted_Data' +year+'.xlsx','Master')
    # split(df3)
    print("Complete")
    return df3

# def split(df):
#     writer = ExcelWriter('Split_Teams_2nd_Time_'+ year +'.xlsx')
#     print('Splitting Data')
#     for i in teamList:
#         df1 = df.loc[df['TEAM_ABBREVIATION_x_x'] == i]
#         df1.to_excel(writer,i)
#
#     writer.save()
#     print('Data Split')


def SplitJointTeams(df,year):
    writer = ExcelWriter('Split_Teams_2nd_Time_'+ year +'.xlsx')
    print('Splitting Data')
    for i in teamList:
        df1 = df.loc[df['TEAM_ABBREVIATION_x'] == i]
        df1['AvgPace_OPP']= df1['vs_AvgPace'].expanding().mean()
        df1['AvgORTG_OPP'] = df1['vs_AvgORTG'].expanding().mean()
        df1['AvgDRTG_OPP'] = df1['vs_AvgORTG'].expanding().mean()
        df1['AvgORTG_L5_OPP'] = df1['AvgORTG_y'].rolling(window=5).mean()
        df1['AvgDRTG_L5_OPP'] = df1['AvgDRTG_y'].rolling(window=5).mean()
        df1['Opp_ORTG_vs_Avg'] = df1['LeagueAvgORTG'] / df1['AvgORTG_OPP']
        df1['Opp_DRTG_vs_Avg'] = df1['LeagueAvgDRTG']/ df1['AvgDRTG_OPP']
        df1['Opp_Pace_vs_Avg'] = df1['AvgPace_OPP'] / df1['LeagueAvgPace']
        df1['Opp_ORTG_vs_Avg_L5'] = df1['LeagueAvgORTG_L5'] / df1['AvgORTG_L5_OPP']
        df1['Opp_DRTG_vs_Avg_L5'] = df1['LeagueAvgDRTG_L5']/ df1['AvgDRTG_L5_OPP']

        df1['std_AvgORTG'] = df1['AvgORTG_x'] * df1['Opp_DRTG_vs_Avg']
        df1['std_AvgDRTG'] = df1['AvgDRTG_x'] / df1['Opp_ORTG_vs_Avg']
        df1['std_AvgORTG_L5'] = df1['AvgORTG_L5_x'] * df1['Opp_DRTG_vs_Avg_L5']
        df1['std_AvgDRTG_L5'] = df1['AvgDRTG_L5_x'] / df1['Opp_ORTG_vs_Avg_L5']


        df1.to_excel(writer,i)

    writer.save()
    return(df1)
    print('Data Split')

def trimDF(df):

    dfbase = df[['GAMECODE','GAME_DATE_EST_x_x','TEAM_ABBREVIATION_x_x','TEAM_ABBREVIATION_y_x','HomeIndex_x_x','DaysRest_x','DaysRest_y','PTS_x_x']]

    df3PT = df[['GAMECODE','TEAM_ABBREVIATION_x_x','avg_FG3A_x','avg_FG3%_x','avg_PCT_FGA_3PT_x','avg_PCT_PTS_3PT_x','avg_PCT_AST_3PM_x','avg_PCT_UAST_3PM_x','vs_avg_FG3A_y',
    'vs_avg_FG3%_y','vs_avg_PCT_FGA_3PT_y','vs_avg_PCT_PTS_3PT_y','vs_avg_PCT_AST_3PM_y','vs_avg_PCT_UAST_3PM_y']]

    dfPAINT = df[['GAMECODE','TEAM_ABBREVIATION_x_x','avg_OREB_x','avg_OREB%_x','avg_PCT_PTS_PAINT_x','avg_PTS_2ND_CHANCE_x','avg_PTS_PAINT_x','avg_Fouls_Drawn_x',
    'avg_times_blocked_x','vs_avg_DREB_y','vs_avg_DREB%_y','vs_avg_PCT_PTS_PAINT_y','vs_avg_PTS_2ND_CHANCE_y','vs_avg_PTS_PAINT_y','vs_avg_Fouls_y',
    'vs_avg_BLK_y']]

    dfFF = df[['GAMECODE','TEAM_ABBREVIATION_x_x','avg_AST_x','avg_AST_PCT_x','avg_AST_TOV_x','avg_AST_RATIO_x','avg_EFG%_x','avg_TS_x','avg_FGA_x',
    'avg_FG%_x','avg_FTA_x','avg_FT%_x','avg_FTA_RATE_x','avg_FT_RATE_x','avg_Fouls_Drawn_x','avg_Steals_x','avg_times_blocked_x','avg_TO_x','avg_TO%_x',
    'avg_OREB_x','avg_OREB%_x','avg_PCT_FGA_2PT_x','avg_PCT_PTS_FT_x','avg_PCT_PTS_OFF_TOV_x','avg_PCT_AST_2PM_x','avg_PCT_AST_FGM_x','avg_PTS_OFF_TOV_x',
    'vs_avg_AST_y','vs_avg_AST_PCT_y','vs_avg_AST_TOV_y','vs_avg_AST_RATIO_y','vs_avg_EFG%_y','vs_avg_TS_y','vs_avg_FGA_y','vs_avg_FG%_y','vs_avg_FTA_y'
    ,'vs_avg_FT%_y','vs_avg_FTA_RATE_y','vs_avg_FT_RATE_y','vs_avg_Fouls_Drawn_y','vs_avg_Steals_y','vs_avg_times_blocked_y','vs_avg_TO_y','vs_avg_TO%_y',
    'vs_avg_OREB_y','vs_avg_OREB%_y','vs_avg_PCT_FGA_2PT_y','vs_avg_PCT_PTS_FT_y','vs_avg_PCT_PTS_OFF_TOV_y','vs_avg_PCT_AST_2PM_y','vs_avg_PCT_AST_FGM_y'
    ,'vs_avg_PTS_OFF_TOV_y','avg_DREB_y','avg_DREB%_y']]

    dfs = [dfbase,df3PT,dfPAINT,dfFF]
    df_final = reduce(lambda left,right: pd.merge(left,right,on=['GAMECODE','TEAM_ABBREVIATION_x_x']), dfs)


    saveToExcel(df_final,'DataForModel_'+ year +'.xlsx','Master')
    return df_final

def getLeagueAvg(df,year):
    df['GAMELINK'] = df['GAMECODE'].str[:8]
    df1 = df[['OFF_RATING_x','GAMELINK']]
    df2 = df[['DEF_RATING_x','GAMELINK']]
    df3 = df[['PACE_x','GAMELINK']]
    df1 = df1.groupby(['GAMELINK'],as_index=False)['OFF_RATING_x'].mean()
    df2 = df2.groupby(['GAMELINK'],as_index=False)['DEF_RATING_x'].mean()
    df3 = df3.groupby(['GAMELINK'],as_index=False)['PACE_x'].mean()

    df4 = pd.merge(df1, df2, on='GAMELINK',how='outer')
    df5 = pd.merge(df4, df3, on='GAMELINK',how='outer')
    df5['LeagueAvgORTG'] = df5['OFF_RATING_x'].expanding().mean()
    df5['LeagueAvgDRTG'] = df5['DEF_RATING_x'].expanding().mean()
    df5['LeagueAvgPace'] = df5['PACE_x'].expanding().mean()
    df5['LeagueAvgORTG'] = df5['LeagueAvgORTG'].shift(1)
    df5['LeagueAvgDRTG'] = df5['LeagueAvgDRTG'].shift(1)
    df5['LeagueAvgPace'] = df5['LeagueAvgPace'].shift(1)

    df5 = df5[['GAMELINK','LeagueAvgORTG','LeagueAvgDRTG','LeagueAvgPace']]
    df6 = pd.merge(df, df5, on='GAMELINK',how='outer')

    saveToExcel(df6,'LeagueAvg_'+ year + '.xlsx','Master')
    return df6


def CalcStats(year):
    # ---------------- Split data out by team and calculate stats ------------------------
    # dataset = getDataSet('AllStats_'+ year + '.xlsx')
    # SplitTeams(dataset,year)
    # ---------------- Consume team data from tabs to make one dataset and combine both teams onto one line------------------------
    d1 = getFirstSplit(year)
    # #---------------- Cosolidate split team data, remove first 6 rows without Last 5 calcs, Filter only needed columns------------------------
    teamdata4 = trimDF(d1)
    print('Stat Calculation Complete')

year = '2017'
# CalcStats(year)


teamdata1 = getFirstSplit(year)
teamdata2 = getLeagueAvg(teamdata1,year)

# dataset = getDataSet('AllStats_'+ year + '_.xlsx')
# SplitTeams(dataset,year)
# teamdata1 = getFirstSplit(year)
