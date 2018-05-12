"""
field_API_util.py

カスタムフィールドを作る為の機能を書くモジュール

参考
[pythonAPI]:https://github.com/fprimex/zdesk/blob/master/zdesk/zdesk_api.py
[fieldAPI]https://developer.zendesk.com/rest_api/docs/core/ticket_fields
"""


"""
tagger形式のフィールドを作る

sampleFieldData.xlsxのような形式のファイルを読み込み
辞書型のデータとして返す

読み込んだファイルは、以下のシートを持っている事を前提としている
    FieldProperties: [fieldAPI]に定義された項目を記載する
    FieldOptions: taggerの項目定義
"""
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
        # 空白セルは無視する
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

"""
taggerFieldDataFactoryが組み立てた辞書型のデータを(おそらく)[PythonAPI]が
Jsonに変換し、zendeskのAPIを叩いてフィールドを作成
"""
def taggerFieldCreate(zendesk, taggerFieldData):
    result = zendesk.ticket_field_create(taggerFieldData)
    pprint(result)


"""
このファイルをpythonインタプリタの引数として与えた時、これが実行される
"""
if __name__ == '__main__':
    zendesk = Zendesk('https://yourcompany.zendesk.com',
                      'you@yourcompany.com', 'passwd')

    taggerFieldData = taggerFieldDataFactory("sampleFieldData.xlsx")

    taggerFieldCreate(zendesk, taggerFieldData)