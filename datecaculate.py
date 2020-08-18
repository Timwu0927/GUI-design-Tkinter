
import datetime
import time
import calendar



def GetToday():
    return datetime.date.today()


def GetDaysList():#輸入要找幾日前的int，輸出一個資料為date格式的陣列
    howmanydays=30
    total=[]
    today=GetToday()
    total.append(today)
    while (howmanydays!=0):
        oneday = datetime.timedelta(days=1)
        yesterday = today - oneday

        total.append(yesterday)
        today=yesterday
        howmanydays=howmanydays-1
  
    return total


def StringToDate(string,format)  :#左側輸入時間字串 右側輸入時間字串原本的格式   輸出時間，格式為2020-07-13
    string=datetime.datetime.strptime(string,format)

    return string

def DateToString(date)  :#輸入時間 輸出字串 格式為2020.07.13
    string=date.strftime('%Y.%m.%d')
    return string

def StringToString(string,format) :#左側輸入時間字串 右側輸入時間字串原本的格式   輸出字串，格式為2020.07.13
    string=StringToDate(string,format)
    string=string.strftime('%Y.%m.%d')
    return string


def GetMSformatToday():
    today=GetToday()
    today=DateToString(today)
    todaylist=today.split('.')
    today=todaylist[0]+'.'+todaylist[1]
    return today
