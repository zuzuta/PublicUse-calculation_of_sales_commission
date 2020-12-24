# 外部公開用 / 一部の固有名詞を匿名化
# 機能：マスターファイルSheet1の入金情報をSheet2へ反映



# ライブラリのインポート
import sys
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
import calendar

from External_Module import commissionPaymentDateCalculation
from External_Module import datetimeToDateOnMaster01
from External_Module import datetimeToDateOnMaster02



# コマンドライン引数引き受け
mspass = sys.argv[1]
dbpass = sys.argv[2]



# Function
# エクセルファイル読み込み
def readMasterFile01(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=0)
    return dfMs

def readMasterFile02(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=1)
    return dfMs

def readDbFile04(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=3)
    return dfDb



# 実行
dfMs01 = readMasterFile01(mspass)
dfMs02 = readMasterFile02(mspass)
dfDb04 = readDbFile04(dbpass)

result01 = commissionPaymentDateCalculation(dfMs01, dfMs02, dfDb04)

dfMs01 = datetimeToDateOnMaster01(dfMs01)
result01 = datetimeToDateOnMaster02(result01)



# エクセルファイル書き出し
with pd.ExcelWriter("Master.xlsx") as writer:
    dfMs01.to_excel(writer, index=False, sheet_name="Sheet1")
    result01.to_excel(writer, index=False, sheet_name="Sheet2")