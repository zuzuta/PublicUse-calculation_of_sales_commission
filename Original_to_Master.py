# 外部公開用 / 一部の固有名詞を匿名化
# 機能：出荷履歴ファイルを取り込み、マスターファイルに出力



# ライブラリインポート
import sys
import pandas as pd
import numpy as np
from dateutil.relativedelta import relativedelta
import calendar

from external_module import prePreocessing01
from external_module import prePreocessing02
from external_module import datetimeToDateOnMaster01
from external_module import datetimeToDateOnMaster02



# コマンドライン引数引き受け
cspass = sys.argv[1]
mspass = sys.argv[2]
dbpass = sys.argv[3]



# Function
# エクセルファイル読み込み
def readCsDownloadedFile(csfile):
    dfCs = pd.read_excel(csfile)
    return dfCs

def readMasterFile01(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=0)
    return dfMs

def readMasterFile02(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=1)
    return dfMs

def readDbFile01(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=0)
    return dfDb

def readDbFile02(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=1)
    return dfDb

def readDbFile03(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=2)
    return dfDb

def readDbFile04(dbfile):
    dfDb = pd.read_excel(dbfile, sheet_name=3)
    return dfDb



# Function
# 出荷履歴ファイル→マスターファイル変換
def concatCsToMaster(dfCs, dfMs01):
    # 不要なcolを削除
    dfCs.drop(["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"], axis=1, inplace=True)

    # col名変更
    dfCs.rename(columns={"得意先":"客先名", "オーダＮＯ":"オーダーNo", "品名":"品番", "品種名":"品種", "個数":"数量", "InvoiceNo":"Invoice"},
    inplace=True)

    # 特定オーダー削除
    dfCs.dropna(subset=["Invoice"], inplace=True)
    dfCs.reset_index(drop=True, inplace=True)

    # 二重登録防止
    droplist = []
    for index, row in dfCs.iterrows():
        if (dfMs01["Invoice"].isin([dfCs.iloc[index, dfCs.columns.get_loc("Invoice")]])).any():
            droplist.append(index)
    dfCs.drop(index=droplist, inplace=True)

    # 通貨振り分け
    dfCs["通貨"] = "JPY"
    
    dfCs["通貨"].where(dfCs["外貨単価"] == 0, other="USD", inplace=True)
    dfCs["単価"].where(dfCs["外貨単価"] == 0, other=dfCs["外貨単価"], inplace=True)
    dfCs["金額"].where(dfCs["外貨金額"] == 0, other=dfCs["外貨金額"], inplace=True)
    
    dfCs.drop(["外貨単価", "外貨金額"], axis=1, inplace=True)

    # dfCsをdfMsにconcat
    dfMs01 = pd.concat([dfMs01, dfCs])
    dfMs01.reset_index(inplace=True, drop=True)
    return dfMs01



def calculateMaster(dfMs01, dfMs02, dfDb01, dfDb02, dfDb03, dfDb04):
    # コミッション対象情報入力1
    for index, row in dfMs01.iterrows():
        if pd.isnull(dfMs01.iloc[index, dfMs01.columns.get_loc("支払サイト")]):
            dfMs01 = prePreocessing01(index, dfMs01, dfMs02, dfDb01, dfDb02)[0]
            dfMs02 = prePreocessing01(index, dfMs01, dfMs02, dfDb01, dfDb02)[1]
    
    dfMs02.reset_index(inplace=True, drop=True)

    # コミッション対象情報入力2
    for index, row in dfMs02.iterrows():
        if pd.isnull(dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep")]):

            # コミッション対象Repが1社の場合の動作
            if dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] == 1:
                dfMs02 = prePreocessing02(index, dfMs02, dfDb03, dfDb04)
            
            # コミッション対象Repが2社の場合の動作
            elif dfMs02.iloc[index, dfMs02.columns.get_loc("コミッション対象Rep数")] == 2:
                customer = dfMs02.iloc[index, dfMs02.columns.get_loc("客先名")]
                parts = dfMs02.iloc[index, dfMs02.columns.get_loc("品番")]

                # オーダーNoとInvoiceの組み合わせをPrimary Keyとして、ダブルコミッション対象行を特定
                lotno = dfMs02.iloc[index, dfMs02.columns.get_loc("オーダーNo")]
                invoice = dfMs02.iloc[index, dfMs02.columns.get_loc("Invoice")]

                # tempidxにリスト形式で該当indexが入る
                tempidx = (dfMs02[(dfMs02["オーダーNo"] == lotno) & (dfMs02["Invoice"] == invoice)]).index
                
                idx = (dfDb03[(dfDb03["客先名"] == customer) & (dfDb03["品番"] == parts) & (dfDb03["ダブルコミッションindex"] == 1)]).index
                dfMs02 = prePreocessing02(tempidx[0], dfMs02, dfDb03, dfDb04, idx)

                idx = (dfDb03[(dfDb03["客先名"] == customer) & (dfDb03["品番"] == parts) & (dfDb03["ダブルコミッションindex"] == 2)]).index
                dfMs02 = prePreocessing02(tempidx[1], dfMs02, dfDb03, dfDb04, idx)

    return [dfMs01, dfMs02]



# 実行
dfCs = readCsDownloadedFile(cspass)
dfMs01 = readMasterFile01(mspass)
dfMs02 = readMasterFile02(mspass)
dfDb01 = readDbFile01(dbpass)
dfDb02 = readDbFile02(dbpass)
dfDb03 = readDbFile03(dbpass)
dfDb04 = readDbFile04(dbpass)
result01 = concatCsToMaster(dfCs, dfMs01)
result02 = calculateMaster(result01, dfMs02, dfDb01, dfDb02, dfDb03, dfDb04)

result02a = datetimeToDateOnMaster01(result02[0])
result02b = datetimeToDateOnMaster02(result02[1])



# エクセルファイル書き出し
with pd.ExcelWriter("Master.xlsx") as writer:
    result02a.to_excel(writer, index=False, sheet_name="Sheet1")
    result02b.to_excel(writer, index=False, sheet_name="Sheet2")