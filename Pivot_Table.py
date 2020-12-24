# 外部公開用 / 一部の固有名詞を匿名化
# 機能：Rep毎のコミッション支払予定、支払確定を月別にPivot Tableへ書き出し



# ライブラリのインポート
import sys
import pandas as pd
import numpy as np
import datetime



# コマンドライン引数引き受け
mspass = sys.argv[1]



# Function
# エクセルファイル読み込み
def readMasterFile02(msfile):
    dfMs = pd.read_excel(msfile, sheet_name=1)
    return dfMs



# Function
# Pivot Table作成
def pivotTable(dfMs02):
    result01 = pd.pivot_table(dfMs02, values="コミッションUSD金額", index="コミッション対象Rep", columns=pd.Grouper(freq="M", key="Rep支払予定月"), aggfunc=np.sum)
    sumByMonth = result01.sum()
    sumByMonth.name = "月別USD合計"
    result01 = result01.append(sumByMonth)

    result02 = pd.pivot_table(dfMs02, values="コミッションJPY金額", index="コミッション対象Rep", columns=pd.Grouper(freq="M", key="Rep支払予定月"), aggfunc=np.sum)
    sumByMonth = result02.sum()
    sumByMonth.name = "月別JPY合計"
    result02 = result02.append(sumByMonth)

    result03 = pd.pivot_table(dfMs02, values="コミッションUSD金額", index="コミッション対象Rep", columns=pd.Grouper(freq="M", key="Rep支払確定月"), aggfunc=np.sum)
    sumByMonth = result03.sum()
    sumByMonth.name = "月別USD合計"
    result03 = result03.append(sumByMonth)

    result04 = pd.pivot_table(dfMs02, values="コミッションJPY金額", index="コミッション対象Rep", columns=pd.Grouper(freq="M", key="Rep支払確定月"), aggfunc=np.sum)
    sumByMonth = result04.sum()
    sumByMonth.name = "月別JPY合計"
    result04 = result04.append(sumByMonth)

    result05 = pd.pivot_table(dfMs02, values="コミッションUSD金額", index="コミッション対象Rep", columns=pd.Grouper(freq="M", key="海外子会社→Rep支払月"), aggfunc=np.sum)
    sumByMonth = result05.sum()
    sumByMonth.name = "月別USD合計"
    result05 = result05.append(sumByMonth)

    estimated = pd.concat([result01, result02])
    fixed = pd.concat([result03, result04])
    via_subsidiary = result05

    return [estimated, fixed, via_subsidiary]



# 実行
dfMs02 = readMasterFile02(mspass)
result = pivotTable(dfMs02)

result[0].fillna(value=0, inplace=True)
result[1].fillna(value=0, inplace=True)
result[2].fillna(value=0, inplace=True)



# datetime.datetime→datetime.date変換 for Sheet1 (Rep支払予定月)
resulta = result[0]
tempdt = resulta.columns.values
tempdd = []

for i in range(len(tempdt)):
    date = pd.to_datetime(str(tempdt[i]))
    tempdd.append(date.strftime('%Y-%m'))

resulta.columns = tempdd



# datetime.datetime→datetime.date変換 for Sheet2 (Rep支払確定月)
resultb = result[1]
tempdt = resultb.columns.values
tempdd = []

for i in range(len(tempdt)):
    date = pd.to_datetime(str(tempdt[i]))
    tempdd.append(date.strftime('%Y-%m'))

resultb.columns = tempdd



# datetime.datetime→datetime.date変換 for Sheet3 (海外子会社→Rep支払月)
resultc = result[2]
tempdt = resultc.columns.values
tempdd = []

for i in range(len(tempdt)):
    date = pd.to_datetime(str(tempdt[i]))
    tempdd.append(date.strftime('%Y-%m'))

resultc.columns = tempdd



# エクセルファイル書き出し
with pd.ExcelWriter("Pivot_Table.xlsx") as writer:
    resulta.to_excel(writer, index=True, sheet_name="Rep支払予定月")
    resultb.to_excel(writer, index=True, sheet_name="Rep支払確定月")
    resultc.to_excel(writer, index=True, sheet_name="海外子会社→Rep支払月")