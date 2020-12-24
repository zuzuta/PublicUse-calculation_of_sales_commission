# 外部公開用 / 一部の固有名詞を匿名化
# 機能：汎用Functionの外部モジュール化



# ライブラリインポート
import sys
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
import calendar
import datetime



# Function
# コミッション対象 該当/非該当 判別
def prePreocessing01(index, dfMs01, dfMs02, dfDb01, dfDb02):
    # 支払サイト入力
    customer = dfMs01.iloc[index, dfMs01.columns.get_loc("客先名")]
    customeridx = dfDb01.query("客先名 == @customer").index
    dfMs01.iloc[index, dfMs01.columns.get_loc("支払サイト")] = pd.Timedelta(dfDb01.iloc[customeridx[0], 1], unit="D")

    # 客先コード入力
    dfMs01.iloc[index, dfMs01.columns.get_loc("客先コード")] = dfDb01.iloc[customeridx[0], dfDb01.columns.get_loc("得意先コード")]
            
    # 入金予定日入力
    dfMs01.iloc[index, dfMs01.columns.get_loc("入金予定日")] = dfMs01.iloc[index, dfMs01.columns.get_loc("出荷日")] + dfMs01.iloc[index, dfMs01.columns.get_loc("支払サイト")]

    # コミッション対象Rep有無の判断
    customer = dfMs01.iloc[index, dfMs01.columns.get_loc("客先名")]
    parts = dfMs01.iloc[index, dfMs01.columns.get_loc("品番")]
    dbidx = (dfDb02[(dfDb02["客先名"] == customer) & (dfDb02["品番"] == parts)]).index

    # Databaseに登録がない客先/品番組み合わせ場合、[コミッション対象Rep数]colにエラー
    if dbidx.size > 0:
        dfMs01.iloc[index, dfMs01.columns.get_loc("コミッション対象Rep数")] = dfDb02.iloc[dbidx[0], dfDb02.columns.get_loc("コミッション対象Rep数")]
    elif dbidx.size == 0:
        dfMs01.iloc[index, dfMs01.columns.get_loc("コミッション対象Rep数")] = "未登録客先/品番"

    # MasterのSheet1からコミッション対象分をSheet2にコピー, ダブルコミッションは2列分コピー
    if dfMs01.iloc[index, dfMs01.columns.get_loc("コミッション対象Rep数")] == 1:
        dfMs02 = pd.concat([dfMs02, dfMs01.iloc[[index]]])
            
    elif dfMs01.iloc[index, dfMs01.columns.get_loc("コミッション対象Rep数")] == 2:
        dfMs02 = pd.concat([dfMs02, dfMs01.iloc[[index]]])
        dfMs02 = pd.concat([dfMs02, dfMs01.iloc[[index]]])
    
    return [dfMs01, dfMs02]



# Function
# コミッション対象オーダーに対する処理（入金完了前）
def prePreocessing02(index, dfMs02, dfDb03, dfDb04, idx=0):
    # コミッション対象Rep入力
    if idx == 0:
        customer = dfMs02.iloc[index, dfMs02.columns.get_loc("客先名")]
        parts = dfMs02.iloc[index, dfMs02.columns.get_loc("品番")]
        idx = (dfDb03[(dfDb03["客先名"] == customer) & (dfDb03["品番"] == parts)]).index
    
    dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")] = dfDb03.iloc[idx[0], dfDb03.columns.get_loc("コミッション対象Rep")]

    # エンドカスタマー入力
    dfMs02.iloc[index, dfMs02.columns.get_loc("エンドカスタマー")] = dfDb03.iloc[idx[0], dfDb03.columns.get_loc("エンドカスタマー")]

    # コミッションレート入力
    dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")] = dfDb03.iloc[idx[0], dfDb03.columns.get_loc("コミッションレート")]

    # Rep地域コード入力
    rep = dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")]
    repidx = (dfDb04[(dfDb04["コミッション対象Rep"] == rep)]).index
    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep地域コード")] = dfDb04.iloc[repidx[0], dfDb04.columns.get_loc("Rep地域コード")]

    # コミッション金額入力
    # 日本円
    if dfMs02.iloc[index, dfMs02.columns.get_loc("通貨")] == "JPY":
        dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションJPY金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]
    # USドル
    else:
        dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションUSD金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]

    # Repへの支払予定月入力
    payday = dfDb04.iloc[repidx[0], dfDb04.columns.get_loc("Rep支払日")]
    epm = dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")].month
    epy = dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")].year

    ### Rep支払日1 -> 1/25, 4/25, 7/25, 10/25 （地域A）
    if payday == 1:
        if 1 <= epm <= 3:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy, month=4, day=25).date()
        elif 4 <= epm <= 6:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy, month=7, day=25).date()
        elif 7 <= epm <= 9:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy, month=10, day=25).date()    
        else:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy + 1, month=1, day=25).date()             

    ### Rep支払日2 -> 1/30, 4/30, 7/30, 10/30 （地域B）
    elif payday == 2:
        if 1 <= epm <= 3:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy, month=4, day=30).date()
        elif 4 <= epm <= 6:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy, month=7, day=30).date()
        elif 7 <= epm <= 9:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy, month=10, day=30).date()      
        else:
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=epy + 1, month=1, day=30).date()  
            
    ### Rep支払日3 -> 毎月、2か月後25日 （地域C）
    elif payday == 3:
        temp = dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")] + relativedelta(months=2)
        tempm = temp.month
        tempy = temp.year
        dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = pd.Timestamp(year=tempy, month=tempm, day=25).date()

    return dfMs02



# Function
# コミッション対象オーダーに対する処理（入金完了後）
def commissionPaymentDateCalculation(dfMs01, dfMs02, dfDb04):
    # 入金日入力
    for index, row in dfMs02.iterrows():
        if pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")]):
            idx = (dfMs01[(dfMs01["Invoice"] == dfMs02.iloc[index, dfMs02.columns.get_loc("Invoice")])]).index
            dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] = dfMs01.iloc[idx[0], dfMs01.columns.get_loc("入金日")]

    # Rep支払確定日入力
    for index, row in dfMs02.iterrows():
        # [Rep支払確定月]colが空欄、且つ[入金日]が入力済の場合
        if pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")]) and pd.notnull(dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")]):
            rep = dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")]
            idx = (dfDb04[(dfDb04["コミッション対象Rep"] == rep)]).index
            payday = dfDb04.iloc[idx[0], dfDb04.columns.get_loc("Rep支払日")]
            pm = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")].month
            py = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")].year

            ### Rep支払日1 -> 1/25, 4/25, 7/25, 10/25 （地域A）
            if payday == 1:
                if 1 <= pm <= 3:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=4, day=25).date()
                elif 4 <= pm <= 6:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=7, day=25).date()
                elif 7 <= pm <= 9:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=10, day=25).date()    
                else:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy + 1, month=1, day=25).date()

            ### Rep支払日2 -> 1/30, 4/30, 7/30, 10/30 （地域B）
            elif payday == 2:
                if 1 <= pm <= 3:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=4, day=30).date()
                elif 4 <= pm <= 6:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=7, day=30).date()
                elif 7 <= pm <= 9:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=10, day=30).date()      
                else:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy + 1, month=1, day=30).date()  
            
            ### Rep支払日3 -> 毎月、2か月後25日 （地域C） + 海外子会社→Rep支払月も同時に入力
            elif payday == 3:
                temp = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] + relativedelta(months=2)
                tempm = temp.month
                tempy = temp.year
                dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=tempy, month=tempm, day=25).date()
                
                temp = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] + relativedelta(months=1)
                tempm = temp.month
                tempy = temp.year
                dfMs02.iloc[index, dfMs02.columns.get_loc("海外子会社→Rep支払月")] = pd.Timestamp(year=tempy, month=tempm, day=calendar.monthrange(tempy, tempm)[1]).date()
    
    return dfMs02



# Function
# エクセルでのdatetime.datetime表示をdatetime.date表示に変換 for Sheet1
def datetimeToDateOnMaster01(dfMs01):
    for index, row in dfMs01.iterrows():
        if isinstance(dfMs01.iloc[index, dfMs01.columns.get_loc("出荷日")], datetime.datetime):
            dfMs01.iloc[index, dfMs01.columns.get_loc("出荷日")] = dfMs01.iloc[index, dfMs01.columns.get_loc("出荷日")].date()
        if isinstance(dfMs01.iloc[index, dfMs01.columns.get_loc("入金予定日")], datetime.datetime):
            dfMs01.iloc[index, dfMs01.columns.get_loc("入金予定日")] = dfMs01.iloc[index, dfMs01.columns.get_loc("入金予定日")].date()
        if isinstance(dfMs01.iloc[index, dfMs01.columns.get_loc("入金日")], datetime.datetime):
            dfMs01.iloc[index, dfMs01.columns.get_loc("入金日")] = dfMs01.iloc[index, dfMs01.columns.get_loc("入金日")].date()
    return dfMs01



# Function
# エクセルでのdatetime.datetime表示をdatetime.date表示に変換 for Sheet2
def datetimeToDateOnMaster02(dfMs02):
    for index, row in dfMs02.iterrows():
        if isinstance(dfMs02.iloc[index, dfMs02.columns.get_loc("出荷日")], datetime.datetime):
            dfMs02.iloc[index, dfMs02.columns.get_loc("出荷日")] = dfMs02.iloc[index, dfMs02.columns.get_loc("出荷日")].date()
        if isinstance(dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")], datetime.datetime):
            dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")] = dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")].date()
        if isinstance(dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")], datetime.datetime):
            dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")].date()
        if isinstance(dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")], datetime.datetime):
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")].date()
        if isinstance(dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")], datetime.datetime):
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")].date()
        if isinstance(dfMs02.iloc[index, dfMs02.columns.get_loc("海外子会社→Rep支払月")], datetime.datetime):
            dfMs02.iloc[index, dfMs02.columns.get_loc("海外子会社→Rep支払月")] = dfMs02.iloc[index, dfMs02.columns.get_loc("海外子会社→Rep支払月")].date()
    return dfMs02