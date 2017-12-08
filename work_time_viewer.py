import pandas as pd
from copy import deepcopy


def addYMD(x, sheetName):
    formatStr = sheetName + " " + x.strftime("%H:%M:%S")
    return pd.to_datetime(formatStr)


def dataCleansing(df, sheetName):
    cleanedData = df.copy()

    print(df)

    cleanedData["日付"] = cleanedData.apply(func=lambda x: pd.to_datetime(sheetName), axis=1)
    cleanedData["YYYY-MM-DD_hh:mm:ss_startTime"] = cleanedData.apply(func=lambda x: addYMD(x["開始時刻"], sheetName), axis=1)
    cleanedData["YYYY-MM-DD_hh:mm:ss_endTime"] = cleanedData.apply(func=lambda x: addYMD(x["終了時刻"], sheetName), axis=1)
    cleanedData["workTime"] = cleanedData.apply(func=lambda x: x["YYYY-MM-DD_hh:mm:ss_endTime"] - x["YYYY-MM-DD_hh:mm:ss_startTime"], axis=1)

    return cleanedData


def getSequentialData(FilePath):

    sheets = pd.ExcelFile(FilePath)
    sheetNames = sheets.sheet_names
    targetSheets = deepcopy(sheetNames)
    targetSheets.remove("テンプレ")

    sequentialData = [dataCleansing(sheets.parse(sheet_name=sheetName), sheetName)
                      for sheetName in targetSheets]
    # print(sequentialData)

    return pd.concat(sequentialData)

def analyzeSequentialData(sequentialData):
    print(sequentialData)
    ana1 = sequentialData.groupby("タスク名")["workTime"].sum()
    ana1.to_excel('taskSum.xlsx', index=True, encoding="UTF-8")
    print(ana1)
    print(sequentialData.resample(rule="D", on="日付").mean())



def exportSequentialData(analyzed):
    pass


def main():
    testpath = "test2.xlsx"
    sequentialData = getSequentialData(testpath)
    sequentialData.to_excel('sequentialData.xlsx', index=False, encoding="UTF-8")
    analyzeSequentialData(sequentialData)
    # print(sequentialData)

if __name__ == '__main__':
    main()
