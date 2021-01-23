# -*- coding: utf-8 -*-
# /*
# 作者：XXX
# 创建时间：XXXX年XX月XX日
# */

# 返回几天后的日期字符串  格式：yyyy-mm-dd
def GetNextDate(days = 1):
    import datetime
    return (datetime.datetime.now() + datetime.timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")[:10]

# 返回几天后的日期字符串  格式：yyyy年mm月dd日
def GetNextDateCn(days = 1):
    import datetime
    # strftime("%Y-%m-%d %H:%M:%S")
    return (datetime.datetime.now() + datetime.timedelta(days=days)).strftime("%Y") + "年" + \
    (datetime.datetime.now() + datetime.timedelta(days=days)).strftime("%m") + "月" + \
    (datetime.datetime.now() + datetime.timedelta(days=days)).strftime("%d") + "日"


# 获取天气数据
def GetWeather():
    import requests
    url = "http://wthrcdn.etouch.cn/weather_mini?citykey=101010100"
    res=requests.get(url)
    sJson = res.text.encode("raw_unicode_escape").decode()
    import json
    return json.loads(sJson)

#将所有航班信息写入excel
def ToExcel(flights=[],columns=[],filePath=""):
    import pandas as pd
    df = pd.DataFrame(flights)
    df["序号"] = df.index + 1
    df.to_excel(filePath,index=False,columns = columns)

#用正则获取文本中的“yyyy-mm-dd hh:mm”格式日期，并转为 “yyyy年mm月dd日 hh:mm”
def GetDatetimeStr(dateTimestr = ""):
    # 2021 - 01 - 16 00:05
    import re
    dateAll = re.findall(r"(\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2})", dateTimestr)
    date = dateAll[0][0:4] + "年" + dateAll[0][5:7] + "月" + dateAll[0][8:10] + "日" + dateAll[0][10:]
    return [date,dateAll[0]]

