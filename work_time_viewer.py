import pandas as pd
from copy import deepcopy


def addYMD(x, sheetName):
    formatStr = sheetName + " " + x.strftime("%H:%M:%S")
    return pd.to_datetime(formatStr)


def addTaskGroup(recode):
    if pd.isnull(recode["タスク分類"]):
        print(recode)
        if recode["タスク名"].find("育成面談") is not -1:
            return "その他"
        if recode["タスク名"].find("回答作成") is not -1:
            return "窓口回答作成"
        if recode["タスク名"].find("回答修正") is not -1:
            return "窓口回答作成"
        if recode["タスク名"].find("休憩") is not -1:
            return "休憩"
        if recode["タスク名"].find("更新") is not -1:
            return "窓口チケット更新"
        if recode["タスク名"].find("回答送付") is not -1:
            return "窓口回答送付"
        if recode["タスク名"].find("JIRA") is not -1:
            return "PoC"
        if recode["タスク名"].find("起票") is not -1:
            return "窓口チケット起票"
        if recode["タスク名"].find("回答相談") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("対応相談") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("相談メール") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("PoC打ち合わせ") is not -1:
            return "PoC"
        if recode["タスク名"].find("メール確認") is not -1:
            return "窓口メール確認"
        if recode["タスク名"].find("メールチェック") is not -1:
            return "窓口メール確認"
        if recode["タスク名"].find("共通業務") is not -1:
            return "共通業務"
        if recode["タスク名"].find("朝会") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("夕会") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("クローズ確認") is not -1:
            return "窓口チケットクローズ"
        if recode["タスク名"].find("Crowd") is not -1:
            return "PoC"
        if recode["タスク名"].find("庶務") is not -1:
            return "その他"
        if recode["タスク名"].find("FW定例") is not -1:
            return "その他"
        if recode["タスク名"].find("管理") is not -1:
            return "窓口チケット更新"
        if recode["タスク名"].find("受付") is not -1:
            return "窓口問い合わせ受付"
        if recode["タスク名"].find("育成") is not -1:
            return "OJT"
        if recode["タスク名"].find("OJT") is not -1:
            return "OJT"
        if recode["タスク名"].find("状況確認") is not -1:
            return "窓口状況確認"
        if recode["タスク名"].find("アクション会議") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("リックソフト") is not -1:
            return "PoC"
        if recode["タスク名"].find("窓口メール") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("ServiceNow") is not -1:
            return "PoC"
        if recode["タスク名"].find("Macchinetta") is not -1:
            return "窓口付随タスク"
        if recode["タスク名"].find("Zendesk稼働報告") is not -1:
            return "窓口付随タスク"
        if recode["タスク名"].find("Zendesk環境構築") is not -1:
            return "PoC"
        if recode["タスク名"].find("Zendesk検証作業") is not -1:
            return "PoC"
        if recode["タスク名"].find("Zendesk打ち合わせ") is not -1:
            return "PoC"
        if recode["タスク名"].find("調査") is not -1:
            return "窓口回答作成"
        if recode["タスク名"].find("窓口内部連携対応") is not -1:
            return "窓口回答相談"
        if recode["タスク名"].find("IBT") is not -1:
            return "その他"
        if recode["タスク名"].find("年末年始対応") is not -1:
            return "窓口付随タスク"
        if recode["タスク名"].find("窓口体制調整") is not -1:
            return "窓口付随タスク"
        if recode["タスク名"].find("ITM") is not -1:
            return "PoC"
        if recode["タスク名"].find("AP基盤チーム打ち合わせ") is not -1:
            return "その他"
    else:
        return recode["タスク分類"]


def dataCleansing(df, sheetName):
    cleanedData = df.copy()

    # print(df)


    cleanedData["日付"] = cleanedData.apply(func=lambda x: pd.to_datetime(sheetName), axis=1)
    cleanedData["YYYY-MM-DD_hh:mm:ss_startTime"] = cleanedData.apply(func=lambda x: addYMD(x["開始時刻"], sheetName), axis=1)
    cleanedData["YYYY-MM-DD_hh:mm:ss_endTime"] = cleanedData.apply(func=lambda x: addYMD(x["終了時刻"], sheetName), axis=1)
    cleanedData["workTime"] = cleanedData.apply(func=lambda x: x["YYYY-MM-DD_hh:mm:ss_endTime"] - x["YYYY-MM-DD_hh:mm:ss_startTime"], axis=1)
    cleanedData["タスク分類"] = cleanedData.apply(func=lambda x: addTaskGroup(x), axis=1)


    # print(cleanedData)
    return cleanedData


def getSequentialData(FilePath):

    sheets = pd.ExcelFile(FilePath)
    sheetNames = sheets.sheet_names
    targetSheets = deepcopy(sheetNames)
    targetSheets.remove("テンプレ")

    sequentialData = []

    for sheet in targetSheets:
        data = sheets.parse(sheetname=sheet)
        cleanedData = dataCleansing(data, sheet)
        sequentialData.append(cleanedData)

    return pd.concat(sequentialData)


def analyzeSequentialData(sequentialData):
    # print(sequentialData)

    ana1 = sequentialData.groupby("タスク分類")["workTime"].sum()
    ana1.to_excel('taskGroupSum.xlsx', index=True, encoding="UTF-8")
    print(ana1)

    ana2 = sequentialData.groupby("タスク名")["workTime"].sum()
    ana2.to_excel('taskNameSum.xlsx', index=True, encoding="UTF-8")
    # print(ana2)

    ana3 = sequentialData.groupby(["タスク分類", "タスク名"])["workTime"].sum()
    ana2.to_excel('taskGroupNameSum.xlsx', index=True, encoding="UTF-8")
    # print(ana3)

    # print(sequentialData.resample(rule="D", on="日付").mean())


def exportSequentialData(analyzed):
    pass


def createDataAndAnalyze():
    testpath = "test.xlsx"
    sequentialData = getSequentialData(testpath)
    sequentialData.to_excel('sequentialData.xlsx', index=False, encoding="UTF-8")
    analyzeSequentialData(sequentialData)
    # print(sequentialData)


def DataAnalyze():
    path = "sequentialData.xlsx"
    df = pd.read_excel(path)
    analyzeSequentialData(df)

if __name__ == '__main__':
    createDataAndAnalyze()
    # DataAnalyze()
