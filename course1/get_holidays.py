import requests
import re
from bs4 import BeautifulSoup
from datetime import date, timedelta, datetime
import pandas as pd

def allweekends(year):
   d = date(year, 1, 1)                    # January 1st
   d += timedelta(days = 6 - d.weekday())  # First Sunday
   while d.year == year:
      yield (d - timedelta(days = 1 )).strftime("%Y/%m/%d"), d.strftime("%Y/%m/%d")
      d += timedelta(days = 7)

response = requests.get("https://holidays-calendar.net/calendar_zh_tw/china_zh_tw.html")
response.encoding = 'utf8'
soup = BeautifulSoup(response.text, "lxml")
#print(soup.prettify())  #輸出排版後的HTML內容
#df = pd.DataFrame(columns = ["日期","星期","假日/國定假日"])
DAY_OF_WEEK = {
    "Monday": "一", "Tuesday": "二", "Wednesday": "三", "Thursday": "四", "Friday": "五", "Saturday": "六", "Sunday": "日"
}
NUMBER_OF_MONTH = {
    1:"jan", 2:"feb", 3:"mar", 4:"apr", 5:"may", 6:"jun", 7:"jul", 8:"aug", 9:"sep", 10:"oct", 11:"nov", 12:"dec"
}
year = soup.find(class_='site-title')
year = re.findall(r'\d+', year.text)[0]
list = [["日期","星期","假日/國定假日"]]
for month in range(1,12):
    jan = soup.find_all('table', id=NUMBER_OF_MONTH[month])
    for days in jan:
        for day in days.find_all(class_='hol'):
            list.append(["{}/{}/{}".format(year, month, day.text.strip()), 
                         DAY_OF_WEEK[datetime(int(year), int(month), int(day.text.strip())).strftime('%A')],
                         '是'])
for weekend in allweekends(int(year)):
    list.append([weekend[0], '六', '是'])
    list.append([weekend[1], '日', '是'])
df = pd.DataFrame(list)
df, df.columns = df[1:] , df.iloc[0]
df = df.set_index('日期')

writer = pd.ExcelWriter('CN國定假日.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='國定假日')
writer.save()
df.to_csv("CN國定假日.csv",encoding='utf-8-sig')