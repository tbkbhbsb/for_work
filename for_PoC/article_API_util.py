"""
article_API_util.py

記事を作る為の機能を書くモジュール

参考
[pythonAPI]:https://github.com/fprimex/zdesk/blob/master/zdesk/zdesk_api.py
[articleAPI]https://developer.zendesk.com/rest_api/docs/help_center/articles
"""

from ast import literal_eval
from pprint import pprint

import sys
import pandas as pd
from zdesk import Zendesk


"""
sampleArticleData.xlsxのような形式のファイルを読み込み
辞書型のデータとして返す

ファイルから読み込んだデータは、内部的に
    baseProperties: [articleAPI]が定義するデータ項目として扱うことを前提としたもの
    additionalProperties: basePropertiesでないもの。これは本文中に埋め込まれる
に分かれる
"""
def articlesDataFactory(DataSourceFilePath):

    df = pd.ExcelFile(DataSourceFilePath)
    # 読み込むシート名の指定
    articleDataSource = df.parse("articleData")

    articlesData = []

    # basePropertiseの指定
    baseProperties = ["locale", "comments_disabled", "draft", "title", "body"]

    # additionalPropertiesの算出
    additionalProperties = list(
        set(articleDataSource.columns).difference(set(baseProperties)))

    # ファイルが含むデータを辞書型のデータに変換
    # ファイルを行ごとにイテレーションして処理する
    for key, val in articleDataSource.iterrows():
        dataDicStr = ""
        articleData = {"article": {}}

        # basePropertiesの部分について、文字列として組み立て
        for baseProperty in baseProperties:
            dataDicStr += '"' + baseProperty + '"' + ": " + \
                '"""' + str(val[baseProperty]) + '"""' + ", "
        # ループ終了時に実行
        else:
            dataDicStr = "{" + dataDicStr + "}"
            # 文字列を辞書として解釈
            articleData["article"] = literal_eval(dataDicStr)

        # additionalProperties を本文に埋め込み
        for additionalProperty in additionalProperties:
            header = "<h1>" + additionalProperty + "</h1>"
            contents = val[additionalProperty]
            articleData["article"]["body"] += header + contents

        articlesData.append(articleData)

    return articlesData


"""
対話的にセクションのIDを取得する
"""
def getTargetSectionID(zendesk):

    sectionsList = zendesk.help_center_sections_list()

    pprint(sectionsList)
    print("セクションのIDを指定して下さい")
    targetSectionID = int(input())

    return targetSectionID


"""
articlesDataFactoryが組み立てた辞書型のデータを(おそらく)[PythonAPI]が
Jsonに変換し、zendeskのAPIを叩いて記事を作成
"""
def articlesCreateOnInstance(zendesk, articlesDataList):
    targetSectionID = getTargetSectionID(zendesk)

    for articleData in articlesDataList:
        result = zendesk.help_center_section_article_create(
            targetSectionID, articleData)
        # デバッグ用
        # print(result)


"""
このファイルをpythonインタプリタの引数として与えた時、これが実行される
"""
if __name__ == '__main__':
    # Usage : <this> https://yourcompany.zendesk.com you@yourcompany.com passwd
    argvs = sys.argv
    argc = len(argvs)
    if (argc != 5):
        print("Usage : <this> https://yourcompany.zendesk.com you@yourcompany.com passwd <filepath>")
        quit()
    
    zendesk = Zendesk(argvs[1], # site
                      argvs[2], # Username
                      argvs[3]) # Password

    articlesData = articlesDataFactory(argvs[4])

    articlesCreateOnInstance(zendesk, articlesData)
