# 外部公開用 / 一部の固有名詞を匿名化
# 機能：例外計上コミッション処理その1を、マスターファイルへ反映



# ライブラリのインポート
import sys
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
import calendar
import datetime

from External_Module import prePreocessing02
from External_Module import commissionPaymentDateCalculation
from External_Module import datetimeToDateOnMaster01
from External_Module import datetimeToDateOnMaster02



# コマンドライン引数引き受け
mspass = sys.argv[1]
dbpass = sys.argv[2]
expass = sys.argv[3]

# inputdateの変数変換
inputdate = datetime.datetime.strptime(sys.argv[4], "%Y/%m/%d")
tempy = inputdate.year
tempm = inputdate.month
inputdate = pd.Timestamp(year=tempy, month=tempm, day=calendar.monthrange(tempy, tempm)[1])



# Function
# エクセルファイル読み込み
def readMasterFile01(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=0)
    return dfMs

def readMasterFile02(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=1)
    return dfMs

def readDbFile03(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=2)
    return dfDb

def readDbFile04(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=3)
    return dfDb

def readExceptionFile01(exfile):
    dfEx = pd.read_excel(exfile, sheet_name=0)
    return dfEx



# Function
# 例外計上コミッション処理その1 / マスターファイルへコピー
def copyExceptionalCommission(dfMs02, dfEx01):
    for index, row in dfEx01.iterrows():
        if dfEx01.iloc[index, dfEx01.columns.get_loc("数量")] > 0:
            # 例外計上登録ファイルから、該当分をマスターファイルへconcat（シングルコミッションは1列分コピー）
            if dfEx01.iloc[index, dfEx01.columns.get_loc("コミッション対象Rep数")] == 1:
                dfMs02 = pd.concat([dfMs02, dfEx01.iloc[[index]]])
            
            # 例外計上登録ファイルから、該当分をマスターファイルへconcat（ダブルコミッションは2列分コピー）
            elif dfEx01.iloc[index, dfEx01.columns.get_loc("コミッション対象Rep数")] == 2:
                dfMs02 = pd.concat([dfMs02, dfEx01.iloc[[index]]])
                dfMs02 = pd.concat([dfMs02, dfEx01.iloc[[index]]])
    
    dfMs02.reset_index(inplace=True, drop=True) 
        return dfMs02



# Function
# 例外計上コミッション処理その1 / マスターファイルへ登録後、計算
def inputExceptionalCommission(dfMs02, dfDb03, dfDb04):
    for index, row in dfMs02.iterrows():
        # 例外計上直後のコミッション対象オーダーを対象に空欄埋め処理（1週目）
        if pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("出荷日")]):
            dfMs02.iloc[index, dfMs02.columns.get_loc("出荷日")] = inputdate
            dfMs02.iloc[index, dfMs02.columns.get_loc("オーダーNo")] = "例外計上 at {0}年{1}月".format(tempy, tempm)
            dfMs02.iloc[index, dfMs02.columns.get_loc("金額")] = dfMs02.iloc[index, dfMs02.columns.get_loc("数量")] * dfMs02.iloc[index, dfMs02.columns.get_loc("単価")]
            dfMs02.iloc[index, dfMs02.columns.get_loc("Invoice")] = dfMs02.iloc[index, dfMs02.columns.get_loc("客先名")] + " / " + dfMs02.iloc[index, dfMs02.columns.get_loc("品番")]
            dfMs02.iloc[index, dfMs02.columns.get_loc("支払サイト")] = "例外計上 at {0}年{1}月".format(tempy, tempm)
            dfMs02.iloc[index, dfMs02.columns.get_loc("入金予定日")] = inputdate
            dfMs02.iloc[index, dfMs02.columns.get_loc("入金日")] = inputdate


            # シングルコミッション時の動作（この時点でRep支払確定月まで計算処理する）
            if dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] == 1:
                dfMs02 = prePreocessing02(index, dfMs02, dfDb03, dfDb04)

            # ダブルコミッション時の動作（1週目はオーダーNoとInvoiceをPrimary Key用に空欄埋め）
            elif dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] == 2:
                continue
    
    tempdate = "例外計上 at {0}年{1}月".format(tempy, tempm)

    for index, row in dfMs02.iterrows():
        # 例外計上直後のコミッション対象オーダーを対象に空欄埋め処理（2週目）
        # ダブルコミッションを、Rep支払確定月まで計算処理する
        if dfMs02.iloc[index, dfMs02.columns.get_loc("オーダーNo")] == tempdate and dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] == 2:
            customer = dfMs02.iloc[index, dfMs02.columns.get_loc("客先名")]
            parts = dfMs02.iloc[index, dfMs02.columns.get_loc("品番")]

            # オーダーNoとInvoiceをPrimary Keyとして該当オーダー特定
            lotno = dfMs02.iloc[index, dfMs02.columns.get_loc("オーダーNo")]
            invoice = dfMs02.iloc[index, dfMs02.columns.get_loc("Invoice")]

            tempidx = (dfMs02[(dfMs02["オーダーNo"] == lotno) & (dfMs02["Invoice"] == invoice)]).index
            
            # DB/Sheet3の[ダブルコミッションindex]が1のRep支払確定月計算処理
            idx = (dfDb03[(dfDb03["客先名"] == customer) & (dfDb03["品番"] == parts) & (dfDb03["ダブルコミッションindex"] == 1)]).index
            dfMs02 = prePreocessing02(tempidx[0], dfMs02, dfDb03, dfDb04, idx)

            # DB/Sheet3の[ダブルコミッションindex]が2のRep支払確定月計算処理
            idx = (dfDb03[(dfDb03["客先名"] == customer) & (dfDb03["品番"] == parts) & (dfDb03["ダブルコミッションindex"] == 2)]).index
            dfMs02 = prePreocessing02(tempidx[1], dfMs02, dfDb03, dfDb04, idx)            

    return dfMs02



# 実行
dfMs01 = readMasterFile01(mspass)
dfMs02 = readMasterFile02(mspass)
dfDb03 = readDbFile03(dbpass)
dfDb04 = readDbFile04(dbpass)
dfEx01 = readExceptionFile01(expass)

result01 = copyExceptionalCommission(dfMs02, dfEx01)
result02 = inputExceptionalCommission(result01, dfDb03, dfDb04)
result03 = commissionPaymentDateCalculation(dfMs01, result02, dfDb04)

dfMs01 = datetimeToDateOnMaster01(dfMs01)
result03 = datetimeToDateOnMaster02(result03)



# エクセルファイル書き出し
with pd.ExcelWriter("{} Master.xlsx".format(datetime.date.today())) as writer:
    dfMs01.to_excel(writer, index=False, sheet_name="Sheet1")
    result03.to_excel(writer, index=False, sheet_name="Sheet2")