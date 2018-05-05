"""
see
https://github.com/fprimex/zdesk/blob/master/zdesk/zdesk_api.py
https://developer.zendesk.com/rest_api/docs/help_center/articles
"""

from ast import literal_eval
from pprint import pprint


import pandas as pd
from zdesk import Zendesk


def articlesDataFactory(DataSourceFilePath):

    df = pd.ExcelFile(DataSourceFilePath)
    articleDataSource = df.parse("articleData")

    articlesData = []

    baseProperties = ["locale", "comments_disabled", "draft", "title", "body"]
    additionalProperties = list(
        set(articleDataSource.columns).difference(set(baseProperties)))

    for key, val in articleDataSource.iterrows():
        dataDicStr = ""
        articleData = {"article": {}}

        for baseProperty in baseProperties:
            dataDicStr += '"' + baseProperty + '"' + ": " + \
                '"' + str(val[baseProperty]) + '"' + ", "
        else:
            dataDicStr = "{" + dataDicStr + "}"
            articleData["article"] = literal_eval(dataDicStr)

        for additionalProperty in additionalProperties:
            header = "<h1>" + additionalProperty + "</h1>"
            contents = val[additionalProperty]
            articleData["article"]["body"] += header + contents

        articlesData.append(articleData)

    return articlesData


def taggerFieldDataFactory(DataSourceFilePath):

    df = pd.ExcelFile(DataSourceFilePath)
    fieldProperties = df.parse("FieldProperties")
    fieldOptions = df.parse("FieldOptions")

    taggerFieldData = {"ticket_field": {}}
    taggerFieldOptionsData = []

    for key, val in fieldProperties.iterrows():
        taggerFieldData["ticket_field"] = {
            "type": val["type"],
            "title": val["title"],
            "required": val["required"]}

    for key, val in fieldOptions.iterrows():

        name = ""

        if pd.notnull(val["NameDepth1"]):
            name += val["NameDepth1"]
        if pd.notnull(val["NameDepth2"]):
            name += "::" + val["NameDepth2"]
        if pd.notnull(val["NameDepth3"]):
            name += "::" + val["NameDepth3"]
        if pd.notnull(val["NameDepth4"]):
            name += "::" + val["NameDepth4"]
        if pd.notnull(val["NameDepth5"]):
            name += "::" + val["NameDepth5"]
        if pd.notnull(val["NameDepth6"]):
            name += "::" + val["NameDepth6"]

        taggerFieldOptionData = {
            "name": name,
            "value": val["value"]}

        taggerFieldOptionsData.append(taggerFieldOptionData)

    taggerFieldData["ticket_field"][
        "custom_field_options"] = taggerFieldOptionsData

    return taggerFieldData


def getTargetSectionID(zendesk):

    sectionsList = zendesk.help_center_sections_list()

    pprint(sectionsList)
    print("セクションのIDを指定して下さい")
    targetSectionID = int(input())

    return targetSectionID


def articlesCreateOnInstance(zendesk, articlesDataList):
    targetSectionID = getTargetSectionID(zendesk)

    for articleData in articlesDataList:
        zendesk.help_center_section_article_create(
            targetSectionID, articleData)


def taggerFieldCreate(zendesk, taggerFieldData):
    result = zendesk.ticket_field_create(taggerFieldData)
    pprint(result)


def makeTestEnv(zendesk):
    testCategoryData = {
        "category":
            {
                "position": 2,
                "locale": "jp",
                "name":  "テスト用",
                "description": "テスト用カテゴリ"
            }
    }
    zendesk.help_center_category_create(data=testCategoryData)

    sectionId = 360000342113
    testSectionData = {
        "section":
            {
                "position": 2,
                "locale": "jp",
                "name": "テスト用セクション",
                "description": "テスト用セクション"
            }}

    zendesk.help_center_category_section_create(sectionId, testSectionData)

if __name__ == '__main__':
    zendesk = Zendesk('https://yourcompany.zendesk.com',
                      'you@yourcompany.com', 'passwd')

    articlesData = articlesDataFactory("sampleArticleData.xlsx")
    # taggerFieldData = taggerFieldDataFactory("sampleFieldData.xlsx")

    articlesCreateOnInstance(zendesk, articlesData)
    # taggerFieldCreate(zendesk, taggerFieldData)
