# 外部公開用 / 一部の固有名詞を匿名化
# 機能：例外計上コミッション処理その2を、マスターファイルへ反映



# ライブラリのインポート
import sys
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
import datetime
import calendar

from External_Module import datetimeToDateOnMaster01
from External_Module import datetimeToDateOnMaster02



# コマンドライン引数引き受け（ファイルパス）
mspass = sys.argv[1]
dbpass = sys.argv[2]
dipass = sys.argv[3]

# コマンドライン引数引き受け（入力日 or 対象日）
inputdate = datetime.datetime.strptime(sys.argv[4], "%Y/%m/%d")

tempy = inputdate.year
tempm = inputdate.month

inputdate = pd.Timestamp(year=tempy, month=tempm, day=calendar.monthrange(tempy, tempm)[1])
modate = inputdate.strftime("S/R of %b/%y  (%m/%d)")
exceptional_date = "例外計上 at {0}年{1}月".format(tempy, tempm)



# Function
# エクセルファイルをpandas.DataFrame型へ読み込み
def readMasterFile01(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=0)
    return dfMs

def readMasterFile02(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=1)
    return dfMs

def readDbFile04(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=3)
    return dfDb

def readDbFile05(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=4)
    return dfDb

def readDistributorFile01(difile):
    diDi = pd.read_excel(difile, sheet_name=0)
    return diDi



# Function
# 例外計上コミッション処理その2 / Distoributorからの売上報告ファイルのうち、コミッション対象アイテムをマスターファイルへ反映
def distributorToMaster(dfMs02, dfDb05, diDi01):
    # マスターファイルのcol毎に空リスト作成
    a = [] #出荷日
    b = [] #オーダーNo
    c = [] #客先コード
    d = [] #客先名
    e = [] #エンドカスタマー
    f = [] #品種
    g = [] #品番
    h = [] #数量
    i = [] #通貨
    j = [] #単価
    k = [] #金額
    l = [] #Invoice
    m = [] #支払サイト
    n = [] #入金予定日
    o = [] #入金日
    p = [] #コミッション対象Rep数

    # 品番毎の総量集計行をdropする
    diDi01 = diDi01[~diDi01["CUSTOMER"].str.contains("TOTAL", na=True)]

    for index, row in diDi01.iterrows():
        # 該当月の出荷数量が0より大きい場合
        if diDi01.iloc[index, diDi01.columns.get_loc(modate)] > 0:
            flag = False
            customer = diDi01.iloc[index, diDi01.columns.get_loc("CUSTOMER")]
            parts = diDi01.iloc[index, diDi01.columns.get_loc("P/N")]

            for dbidx, dbrow in dfDb05.iterrows():
                # エンドカスタマー&品番が一致し、コミッション対象Repが存在する場合（＝データベース登録済かつコミッション対象の場合）
                if ((dfDb05.iloc[dbidx, dfDb05.columns.get_loc("エンドカスタマー")] == customer) &
                (dfDb05.iloc[dbidx, dfDb05.columns.get_loc("品番")] == parts) &
                pd.notnull(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("コミッション対象Rep")])):

                    a.append(inputdate)
                    b.append(exceptional_date)
                    c.append(99999)
                    d.append("特殊計上02")
                    e.append(customer)
                    f.append(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("品種")])
                    g.append(parts)
                    h.append(diDi01.iloc[index, diDi01.columns.get_loc(modate)] * 1000)
                    i.append(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("特殊計上02通貨")])
                    j.append(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("特殊計上02単価")])
                    k.append(diDi01.iloc[index, diDi01.columns.get_loc(modate)] * 1000 * dfDb05.iloc[dbidx, dfDb05.columns.get_loc("特殊計上02単価")])
                    l.append(customer + "/" + parts)
                    m.append(exceptional_date)
                    n.append(inputdate)
                    o.append(inputdate)
                    p.append(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("コミッション対象Rep数")])
                    flag = True
                    break

                # エンドカスタマー&品番が一致し、コミッション対象Repが存在しない場合（＝データベース登録済かつコミッション対象外の場合）
                elif ((dfDb05.iloc[dbidx, dfDb05.columns.get_loc("エンドカスタマー")] == customer) &
                (dfDb05.iloc[dbidx, dfDb05.columns.get_loc("品番")] == parts) &
                pd.isnull(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("コミッション対象Rep")])): 
                    flag = True
                    break
            
            # データベース未登録のエンドカスタマー/品番組み合わせの場合
            if not flag:
                a.append(inputdate)
                b.append(exceptional_date)
                c.append(99999)
                d.append("特殊計上02")
                e.append(customer)
                f.append(dfDb05.iloc[dbidx, dfDb05.columns.get_loc("品種")])
                g.append(parts)
                h.append(diDi01.iloc[index, diDi01.columns.get_loc(modate)] * 1000)
                i.append("新規登録")
                j.append("新規登録")
                k.append("新規登録")
                l.append("新規登録")
                m.append("新規登録")
                n.append(inputdate)
                o.append(inputdate)
                p.append("新規登録")             

    temp = pd.DataFrame(data={"出荷日" : a, "オーダーNo" : b, "客先コード" : c, "客先名" : d, "エンドカスタマー" : e, "品種" : f, "品番" : g, "数量" : h,
    "通貨" : i, "単価" : j, "金額" : k, "Invoice" : l, "支払サイト" : m, "入金予定日" : n, "入金日" : o, "コミッション対象Rep数" : p} )
    
    dfMs02 = pd.concat([dfMs02, temp])
    dfMs02.reset_index(inplace=True, drop=True)
    
    return dfMs02



# Function
# 例外計上コミッション処理その2 / マスターファイルへ新規登録したコミッション対象オーダーのコミッション額計算処理
def commissionAmountCalculation(dfMs02, dfDb05):
    for index, row in dfMs02.iterrows():
        # [コミッション対象Rep]colが空欄、かつ新規登録エラーでない場合
        if pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")]) and (dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] != "新規登録"):
            
            # コミッション対象Rep入力
            customer = dfMs02.iloc[index, dfMs02.columns.get_loc("客先名")]
            parts = dfMs02.iloc[index, dfMs02.columns.get_loc("品番")]
            endcustomer = dfMs02.iloc[index, dfMs02.columns.get_loc("エンドカスタマー")]
            idx = (dfDb05[(dfDb05["客先名"] == customer) & (dfDb05["品番"] == parts) & (dfDb05["エンドカスタマー"] == endcustomer)]).index
            dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")] = dfDb05.iloc[idx[0], dfDb05.columns.get_loc("コミッション対象Rep")]

            # コミッションレート入力
            dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")] = dfDb05.iloc[idx[0], dfDb05.columns.get_loc("コミッションレート")]

            # コミッション額計算
            if dfMs02.iloc[index, dfMs02.columns.get_loc("通貨")] == "JPY":
                dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションJPY金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]
            else:
                dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションUSD金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]

            # ダブルコミッション時の動作
            if dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] == 2:
                
                # マスターファイル/Sheet2における、現在行の次行index（＝ダブルコミッションは連続行登録されている）
                index += 1

                # 現在行におけるデータベースの[ダブルコミッションindex]が1の場合
                if dfDb05.iloc[idx[0], dfDb05.columns.get_loc("ダブルコミッションindex")] == 1:
                    # [ダブルコミッションindex]2を探す
                    newidx = (dfDb05[(dfDb05["客先名"] == customer) & (dfDb05["品番"] == parts) & (dfDb05["エンドカスタマー"] == endcustomer) & (dfDb05["ダブルコミッションindex"] == 2)]).index
                    dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")] = dfDb05.iloc[newidx[0], dfDb05.columns.get_loc("コミッション対象Rep")] # コミッション対象Rep
                    dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")] = dfDb05.iloc[newidx[0], dfDb05.columns.get_loc("コミッションレート")] # コミッションレート
                    # コミッション額計算
                    if dfMs02.iloc[index, dfMs02.columns.get_loc("通貨")] == "JPY":
                        dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションJPY金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]
                    else:
                        dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションUSD金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]

                # 現在行におけるデータベースの[ダブルコミッションindex]が2の場合
                elif dfDb05.iloc[idx[0], dfDb05.columns.get_loc("ダブルコミッションindex")] == 2:
                    # [ダブルコミッションindex]1を探す
                    newidx = (dfDb05[(dfDb05["客先名"] == customer) & (dfDb05["品番"] == parts) & (dfDb05["エンドカスタマー"] == endcustomer) & (dfDb05["ダブルコミッションindex"] == 1)]).index
                    dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")] = dfDb05.iloc[newidx[0], dfDb05.columns.get_loc("コミッション対象Rep")] # コミッション対象Rep
                    dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")] = dfDb05.iloc[newidx[0], dfDb05.columns.get_loc("コミッションレート")] # コミッションレート
                    # コミッション額計算
                    if dfMs02.iloc[index, dfMs02.columns.get_loc("通貨")] == "JPY":
                        dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションJPY金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]
                    else:
                        dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションUSD金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] * dfMs02.iloc[index, dfMs02.columns.get_loc("コミッションレート")]

    return dfMs02



# Function
# 例外計上コミッション処理その2 / マスターファイルへ新規登録したコミッション対象オーダーの支払日計算処理
def commissionScheduleCalculation(dfMs02, dfDb04):
    for index, row in dfMs02.iterrows():
        # Rep支払確定月、Rep支払予定月、Rep地域コードが空欄、且つ[入金日]が入力済、且つ新規登録エラーでないの場合
        if pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")]) and pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")]) and \
            pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("Rep地域コード")]) and pd.notnull(dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")]) and \
            (dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] != "新規登録"):

            # Rep支払確定月入力
            rep = dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")]
            idx = (dfDb04[(dfDb04["コミッション対象Rep"] == rep)]).index
            payday = dfDb04.iloc[idx[0], dfDb04.columns.get_loc("Rep支払日")]

            epm = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")].month
            epy = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")].year

            ### Rep支払日1 -> 1/25, 4/25, 7/25, 10/25 （地域A）
            if payday == 1:
                if 1 <= epm <= 3:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=4, day=25)
                elif 4 <= epm <= 6:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=7, day=25)
                elif 7 <= epm <= 9:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=10, day=25)      
                else:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy + 1, month=1, day=25)              

            ### Rep支払日2 -> 1/30, 4/30, 7/30, 10/30 （地域B）
            elif payday == 2:
                if 1 <= epm <= 3:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=4, day=30)
                elif 4 <= epm <= 6:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=7, day=30)
                elif 7 <= epm <= 9:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy, month=10, day=30)      
                else:
                    dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=epy + 1, month=1, day=30)  
            
            ### Rep支払日3 -> 毎月、2か月後25日 （地域C） + 海外子会社→Rep支払月も同時に入力
            elif payday == 3:
                temp = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] + relativedelta(months=2)
                tempm = temp.month
                tempy = temp.year
                dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")] = pd.Timestamp(year=tempy, month=tempm, day=25)
                
                temp = dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] + relativedelta(months=1)
                tempm = temp.month
                tempy = temp.year
                dfMs02.iloc[index, dfMs02.columns.get_loc("海外子会社→Rep支払月")] = pd.Timestamp(year=tempy, month=tempm, day=calendar.monthrange(tempy, tempm)[1])
        
            # Rep支払予定付月 
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払予定月")] = dfMs02.iloc[index, dfMs02.columns.get_loc("Rep支払確定月")]

            # Rep地域コード
            dfMs02.iloc[index, dfMs02.columns.get_loc("Rep地域コード")] = dfDb04.iloc[idx[0], dfDb04.columns.get_loc("Rep地域コード")]

    return dfMs02



# 実行
dfMs01 = readMasterFile01(mspass)
dfMs02 = readMasterFile02(mspass)
dfDb04 = readDbFile04(dbpass)
dfDb05 = readDbFile05(dbpass)
diDi01 = readDistributorFile01(dipass)

result01 = distributorToMaster(dfMs02, dfDb05, diDi01)
result02 = commissionAmountCalculation(result01, dfDb05)
result03 = commissionScheduleCalculation(result02, dfDb04)

dfMs01 = datetimeToDateOnMaster01(dfMs01)
result03 = datetimeToDateOnMaster02(result03)



# エクセルファイル書き出し
with pd.ExcelWriter("{} Master.xlsx".format(datetime.date.today())) as writer:
    dfMs01.to_excel(writer, index=False, sheet_name="Sheet1")
    result03.to_excel(writer, index=False, sheet_name="Sheet2")