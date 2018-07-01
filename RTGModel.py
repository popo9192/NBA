import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.preprocessing import Imputer,LabelEncoder, OneHotEncoder,StandardScaler,PolynomialFeatures
from sklearn.cross_validation import train_test_split
from sklearn.linear_model import LinearRegression
import statsmodels.formula.api as sm
from sklearn.svm import SVR
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor
from pandas import ExcelWriter
from sklearn.metrics import mean_squared_error,mean_absolute_error,explained_variance_score,r2_score
import pickle
import os
from base_functions import (saveToExcel,getDataSet,getDataSetcsv,getDateCutoff,getHomeIndex,getHomeORTG,getAwayORTG,resetRest,VSresetRest,getHomeDRTG,
getAwayDRTG,getLocationORTG,getLocationDRTG)
from Predict import getResults
os.chdir('C:\\Users\\Peter Haley\\Desktop\\Projects\\Data_Science\\Python\\NBA\\Excel')

#-------------Get Dataset-------------

def buildModel(df):
    df = df.dropna()
    # df = df.loc[df['GAME_DATE_EST_x_x'] >= '2017-11-01']
    dfh = df.loc[df['HomeIndex_x_x'] == 0]
    dfa = df.loc[df['HomeIndex_x_x'] == 1]
    y = df[['GAMECODE','TEAM_ABBREVIATION_x_x','PTS_x_x']]
    yh = dfh[['GAMECODE','TEAM_ABBREVIATION_x_x','PTS_x_x']]
    ya = dfa[['GAMECODE','TEAM_ABBREVIATION_x_x','PTS_x_x']]

    # df =df[['avg_FG3A_x','avg_FG3%_x','avg_PCT_FGA_3PT_x','avg_PCT_PTS_3PT_x','avg_PCT_AST_3PM_x','avg_PCT_UAST_3PM_x','vs_avg_FG3A_y','vs_avg_FG3%_y'
    # ,'vs_avg_PCT_FGA_3PT_y','vs_avg_PCT_PTS_3PT_y']]
    #
    # df =df[['AvgORTG_x','AvgNET_x','vs_AvgDRTG_x','vs_AvgNET_x']]

# ----------------------FOR HOME AWAY SPLIT--------------------------------------
    dfh =dfh[['AvgORTG_x','HomeORTG_x','est_avg_Poss_x','vs_AvgDRTG_x','vs_AwayDRTG_x','vs_est_avg_Poss_x']]
    dfa =dfa[['AvgORTG_x','AwayORTG_x','est_avg_Poss_x','vs_AvgDRTG_x','vs_HomeDRTG_x','vs_est_avg_Poss_x']]
    return df, y, dfh, dfa, yh, ya


def Regression(df):
    df, y, dfh, dfa, yh, ya = buildModel(df)
    y_stats = y.iloc[:, 2].values
    # NOT SPLITTING HOME AND AWAY
    # x = df.values
    # y = y.values

    # FOR HOME AND AWAY SPLIT
    df, y, dfh, dfa, yh, ya = buildModel(df)
    xh = dfh.values
    xa = dfa.values
    yh = yh.values
    ya = ya.values

    # -------------Split Train and Test Data-------------
    # NOT SPLITTING HOME AND AWAY
    # x_train, x_test, y_train, y_test = train_test_split(x,y,test_size =0.25, random_state =0)
    # y = y[:,2]
    #
    # y_train = y_train[:,2]
    # y_compare = y_test
    # y_test = y_test[:,2]

    # # FOR HOME AND AWAY SPLIT
    x_train_h, x_test_h, y_train_h, y_test_h = train_test_split(xh,yh,test_size =0.25, random_state =0)
    yh = yh[:,2]
    y_train_h = y_train_h[:,2]
    y_compare_h = y_test_h
    y_test_h = y_test_h[:,2]

    x_train_a, x_test_a, y_train_a, y_test_a = train_test_split(xa,ya,test_size =0.25, random_state =0)
    ya = ya[:,2]
    y_train_a = y_train_a[:,2]
    y_compare_a = y_test_a
    y_test_a = y_test_a[:,2]



    # ------------------Linear--------------------

    # ------------------------------REGRESSION TIME------------------------------
    # NOT SPLITTING HOME AND AWAY
    # regressor = LinearRegression()
    # regressor.fit(x_train, y_train)

    # # FOR HOME AND AWAY SPLIT
    regressor_h = LinearRegression()
    regressor_a = LinearRegression()
    regressor_h.fit(x_train_h, y_train_h)
    regressor_a.fit(x_train_a, y_train_a)


    # #-------------Predict a new result with Random Forest-------------
    # NOT SPLITTING HOME AND AWAY
    # y_pred = regressor.predict(x_test)

     # FOR HOME AND AWAY SPLIT
    y_pred_h = regressor_h.predict(x_test_h)
    y_pred_a = regressor_a.predict(x_test_a)

    # #------------------RANDOM FOREST--------------------
    # regressorRF = RandomForestRegressor(n_estimators=3000, random_state=0)
    #
    # regressorRF.fit(x_train,y_train)
    # y_pred = regressorRF.predict(x_test)
    # r2 = regressorRF.score(x_train, y_train)
    # mae = mean_absolute_error(y_test, y_pred)
    # mse = mean_squared_error(y_test, y_pred)
    # evs = explained_variance_score(y_test, y_pred)
    #
    # print(r2)

    # print('MAE ', mae)
    # print('MSE ', mse)
    # print('Explained Variance ', evs)

    #
    # imp = regressorRF.feature_importances_
    # print(imp)

    # -----------------LINEAR model Scores-------------------
    # import statsmodels.api as sm
    # x = sm.add_constant(x)
    # x_opt = x[:,[0,1]]
    # # print(x_opt.dtype,y.dtype,y_stats.dtype)
    # regressor_ols = sm.OLS(endog = y_stats, exog = x_opt).fit()
    # # print(regressor_ols.summary())

    #-----------------------OUTPUT---------------------
        # NOT SPLITTING HOME AND AWAY
    # df1 = pd.DataFrame({'GAMECODE':y_compare[:,0],'TEAM_ABBREVIATION_x_x':y_compare[:,1],'Actual':y_test,'Predicted':y_pred})
    # df1 = df1.sort_values(by=['GAMECODE'],ascending=[True])
    # # df4 = pd.merge(df3, df1, on=['GAMECODE','TEAM_ABBREVIATION_x'])
    # df1['Mean_Avg_Err'] = (df1['Predicted'] - df1['Actual']).abs()
    # df1['Mean_SQ_Err'] =  df1['Mean_Avg_Err'] * df1['Mean_Avg_Err']
    # df1['Last_10_Avg_Err'] = df1['Mean_Avg_Err'].rolling(window=10).mean()
    # df1['Last_10_SQ_Err'] = df1['Mean_SQ_Err'].rolling(window=10).mean()
    # saveToExcel(df1,'Model_Results.xlsx','Master')
    #
    # print('MAE:',df1['Mean_Avg_Err'].mean())
    # print('MSE:',df1['Mean_SQ_Err'].mean())
    #
    # RTGModelFile = 'RTG_Model.sav'
    # pickle.dump(regressor,open(RTGModelFile,'wb'))


    # -------------------------------FOR HOME AND AWAY SPLIT-----------------------------
    dfh1 = pd.DataFrame({'GAMECODE_x':y_compare_h[:,0],'TEAM_ABBREVIATION_x':y_compare_h[:,1],'Actual':y_test_h,'Predicted':y_pred_h})
    dfh1 = dfh1.sort_values(by=['GAMECODE_x'],ascending=[True])
    # df4 = pd.merge(df3, df1, on=['GAMECODE','TEAM_ABBREVIATION_x'])
    dfh1['Mean_Avg_Err'] = (dfh1['Predicted'] - dfh1['Actual']).abs()
    dfh1['Mean_SQ_Err'] =  dfh1['Mean_Avg_Err'] * dfh1['Mean_Avg_Err']
    dfh1['Last_10_Avg_Err'] = dfh1['Mean_Avg_Err'].rolling(window=10).mean()
    dfh1['Last_10_SQ_Err'] = dfh1['Mean_SQ_Err'].rolling(window=10).mean()


    dfa1 = pd.DataFrame({'GAMECODE_x':y_compare_a[:,0],'TEAM_ABBREVIATION_x':y_compare_a[:,1],'Actual':y_test_a,'Predicted':y_pred_a})
    dfa1 = dfa1.sort_values(by=['GAMECODE_x'],ascending=[True])
    dfa1['Mean_Avg_Err'] = (dfa1['Predicted'] - dfa1['Actual']).abs()
    dfa1['Mean_SQ_Err'] =  dfa1['Mean_Avg_Err'] * dfa1['Mean_Avg_Err']
    dfa1['Last_10_Avg_Err'] = dfa1['Mean_Avg_Err'].rolling(window=10).mean()
    dfa1['Last_10_SQ_Err'] = dfa1['Mean_SQ_Err'].rolling(window=10).mean()

    df2 = dfh1.append(dfa1)
    df2 = df2.sort_values(by=['GAMECODE_x'],ascending=[True])
    # saveToExcel(df2,'Model_Results.xlsx','Master')

    print('Home MAE:',dfh1['Mean_Avg_Err'].mean())
    print('Home MSE:',dfh1['Mean_SQ_Err'].mean())
    print('Away MAE:',dfa1['Mean_Avg_Err'].mean())
    print('Away MSE:',dfa1['Mean_SQ_Err'].mean())

    # awayModelFile = 'Backtest_Away_Model.sav'
    # pickle.dump(regressor_a,open(awayModelFile,'wb'))
    #
    # homeModelFile = 'Backtest_Home_Model.sav'
    # pickle.dump(regressor_h,open(homeModelFile,'wb'))


def RunModels(df,odds):
    dfx, dfy, dfh, dfa, yh, ya = buildModel(df)

    home_model = LoadModel('Backtest_Home_Model.sav')
    away_model = LoadModel('Backtest_Away_Model.sav')

    xh = dfh.values
    xa = dfa.values

    # yh = yh.values
    # ya = ya.values

    homepred = home_model.predict(xh)
    awaypred = away_model.predict(xa)


    dfh1 = pd.DataFrame({'GAMECODE_x':yh['GAMECODE'],'TEAM_ABBREVIATION_x_x':yh['TEAM_ABBREVIATION_x_x'],'Predicted':homepred})
    dfa1 = pd.DataFrame({'GAMECODE_x':ya['GAMECODE'],'TEAM_ABBREVIATION_x_x':ya['TEAM_ABBREVIATION_x_x'],'Predicted':awaypred})
    df2 = dfh1.append(dfa1)
    df3 = pd.merge(df2, odds, on=['GAMECODE_x','TEAM_ABBREVIATION_x_x'],how='left')
    df3['CalculatedSpread'] = df3['Predicted'].shift(1) - df3['Predicted']
    df3['CalculatedOU'] = df3['Predicted'].shift(-1) + df3['Predicted']
    df3['CalculatedLines'] = df3.apply(getCalcedLines,axis=1)
    df3 = df3[['GAMECODE_x','TEAM_ABBREVIATION_x_x','Predicted','CalculatedLines','VegasLines']]
    df3['Difference'] = df3['CalculatedLines'] - df3['VegasLines']
    df3['BetGrade'] = df3.apply(gradeBet,axis=1)

    # df4 = pd.merge(df3, df, on=['GAMECODE_x','TEAM_ABBREVIATION_x'],how='left')
    saveToExcel(df3,'BackTest_Data.xlsx','Master')
    return df3

def getResultSummary(df1):
    dfb = pd.DataFrame(columns=['Total Win%','Over/Under', 'Spread','A','B','C','D','Over/Under-A','Over/Under-B'
    ,'Over/Under-C','Over/Under-D','Spread-A','Spread-B','Spread-C','Spread-D'])
    grades = ['A','B','C','D']
    month = 'total'
    stats={}

    # df1 = df.loc[df['DateTrim'] == month]
    total = (df1['BetGrade']).count()
    losses = (df1['BetGrade'][df1['Correct'] == 'Loss']).count()
    Total_Win_Per = (1-(losses/total))*100
    stats[month +' Win %'] = Total_Win_Per
    # print(total,losses)

    total_OU = (df1['BetGrade'][df1['BetType'] == 'OU']).count()
    losses_OU = df1['BetGrade'][((df1.Correct == 'Loss') & (df1.BetType == 'OU'))].count()
    Total_Win_OU = (1-(losses_OU/total_OU))*100
    stats[month +' Over Under Win %'] = Total_Win_OU

    total_Sp = (df1['BetGrade'][df1['BetType'] == 'Spread']).count()
    losses_Sp = df1['BetGrade'][((df1.Correct == 'Loss') & (df1.BetType == 'Spread'))].count()
    Total_Win_Sp = (1-(losses_Sp/total_Sp))*100
    stats[month +' Spread Win %'] = Total_Win_Sp

    for x in grades:
        total = (df1['BetGrade'][df1['BetGrade'] == x]).count()
        losses = (df1['BetGrade'][((df1.Correct == 'Loss') & (df1.BetGrade == x))]).count()
        Total_Win_Per = (1-(losses/total))*100
        stats[x + ' - ' + month +' Win %'] = Total_Win_Per

        total_OU = df1['BetGrade'][((df1.BetGrade == x) & (df1.BetType == 'OU'))].count()
        losses_OU = df1['BetGrade'][((df1.Correct == 'Loss') & (df1.BetType == 'OU')& (df1.BetGrade == x))].count()
        Total_Win_OU = (1-(losses_OU/total_OU))*100
        stats[x + ' - ' + month +' Over Under Win %'] = Total_Win_OU

        total_sp = df1['BetGrade'][((df1.BetGrade == x) & (df1.BetType == 'Spread'))].count()
        losses_sp = df1['BetGrade'][((df1.Correct == 'Loss') & (df1.BetType == 'Spread')& (df1.BetGrade == x))].count()
        Total_Win_sp = (1-(losses_sp/total_sp))*100
        stats[x + ' - ' + month +' Spread Win %'] = Total_Win_sp

    dfr = pd.DataFrame({'Total Win%':stats[month +' Win %'],'Over/Under':stats[month +' Over Under Win %'],
    'Spread':stats[month +' Spread Win %'],'A':stats['A' + ' - ' + month +' Win %'],'B':stats['B' + ' - ' + month +' Win %'],
    'C':stats['C' + ' - ' + month +' Win %'],'D':stats['D' + ' - ' + month +' Win %'],'Over/Under-A':stats['A' + ' - ' + month +' Over Under Win %'],
    'Over/Under-B':stats['B' + ' - ' + month +' Over Under Win %'],'Over/Under-C':stats['C' + ' - ' + month +' Over Under Win %'],
    'Over/Under-D':stats['D' + ' - ' + month +' Over Under Win %'],'Spread-A':stats['A' + ' - ' + month +' Spread Win %'],
    'Spread-B':stats['B' + ' - ' + month +' Spread Win %'],'Spread-C':stats['C' + ' - ' + month +' Spread Win %'],
    'Spread-D':stats['D' + ' - ' + month +' Spread Win %']}, index=[0])

    dfb = dfb.append(dfr)
    # dfb['Month'] = dfb.apply(convertMonth,axis=1)
    dfb = dfb[['Total Win%','Over/Under', 'Spread','A','B','C','D','Over/Under-A','Over/Under-B'
    ,'Over/Under-C','Over/Under-D','Spread-A','Spread-B','Spread-C','Spread-D']]
    # print(stats)
    # print(dfb.head())
    # saveToExcel(dfb,'Backtest_Summary.xlsx','Master')
    return dfb

monthList = {'10':'October','11':'November','12':'December','01':'January','02':'February','03':'March','04':'April','05':'May',
'06':'June'}

def convertMonth(df):
    return monthList[df['MonthCode']]

def getDataSet(dataset):
    df = pd.read_excel(dataset)
    return df

def LoadModel(filename):
    loaded_model = pickle.load(open(filename, 'rb'))
    return loaded_model


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
    if df['Difference'] >= 20 or df['Difference'] <= -20:
        return 'A'
    elif (df['Difference'] >= 10 and df['Difference'] < 20) or (df['Difference'] <= -10 and df['Difference'] >= -20)  :
        return 'B'
    elif (df['Difference'] > 5 and df['Difference'] < 10) or (df['Difference'] < -5 and df['Difference'] > -10)  :
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

def backtest(year,both):
    odds = getDataSet('Historical_Odds_'+ year +'.xlsx')
    if both:
        odds1 = getDataSet('Historical_Odds_'+ '2015' +'.xlsx')
        odds2 = getDataSet('Historical_Odds_'+ '2016' +'.xlsx')
        frames = [odds1, odds2]
        odds = pd.concat(frames)
    df = getDataSet('DataForModel_'+ year +'.xlsx')
    if both:
        df1 = getDataSet('DataForModel_'+ '2015' +'.xlsx')
        df2 = getDataSet('DataForModel_'+ '2016' +'.xlsx')
        frames = [df1, df2]
        df = pd.concat(frames)
    actual = getDataSet('AllStats_' + year + '.xlsx')
    if both:
        actual1 = getDataSet('AllStats_' + '2015' + '.xlsx')
        actual2 = getDataSet('AllStats_' + '2016' + '.xlsx')
        frames = [actual1, actual2]
        actual = pd.concat(frames)
    Regression(df)
    proj = RunModels(df,odds)
    results = getResults(actual,proj,'Backtest')
    # dfr = pd.merge(dfd, results, on=['GAMECODE_x','TEAM_ABBREVIATION_x'],how='left')
    # saveToExcel(dfr,'BackTest_Data.xlsx','Master')
    dfb = getResultSummary(results)
    if both:
        saveToExcel(dfb,'BackTest_Summary.xlsx','Both')
    if not both:
        saveToExcel(dfb,'BackTest_Summary.xlsx',year)



year = '2015'
both = False
backtest(year,both)
