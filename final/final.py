import requests
import json
import time
from selenium import webdriver
import pandas as pd
import re
import os


def auto_save_file(path):
    directory, file_name = os.path.split(path)
    while os.path.isfile(path):
        pattern = '(\d+)\)\.'
        if re.search(pattern, file_name) is None:
            file_name = file_name.replace('.', '(0).')
        else:
            current_number = int(re.findall(pattern, file_name)[-1])
            new_number = current_number + 1
            file_name = file_name.replace(
                f'({current_number}).', f'({new_number}).')
        path = os.path.join(directory + os.sep + file_name)
    return path


def change_time_to_TW(year, month, day, hour, minute):
    timeString = year+'/'+str(month)+'/' + str(day)+' ' + \
        str(hour)+':'+minute  # 時間格式為字串
    struct_time = time.strptime(timeString, "%Y/%m/%d %H:%M")  # 轉成時間元組
    time_stamp = int(time.mktime(struct_time))
    hour_stamp = 28800
    time_stamp += hour_stamp
    return_time = time.localtime(time_stamp)  # 轉成時間元組
    timeString = time.strftime("%Y/%m/%d %H:%M", return_time)
    return timeString


localtime = time.localtime()
result = time.strftime("%Y-%m-%d %I:%M:%S %p", localtime)
fullyear = result[0:4]
halfyear = result[2:4]

df = pd.read_excel("C:/typhoon/NewPath/newpath.xlsx", sheet_name='CODE')
step = 0
while 1:
    try:
        df.at[step, 'code_tw']
    except:
        step -= 1
        break
    step += 1
TWnum = str(df.at[step, 'code_tw'])
HKnum = str(df.at[step, 'code_hk'])
CNnum = str(df.at[step, 'code_cn'])
JPnum = str(df.at[step, 'code_jp'])
USnum = str(df.at[step, 'code_us'])
KRnum = str(df.at[step, 'code_kr'])

# US
urlUS = 'https://www.metoc.navy.mil/jtwc/products/wp' + \
    USnum+halfyear+'web.txt'  # 目標下載連結

r = requests.get(urlUS).text

lst = []
begin = 0

while r[begin:-1].find('VALID AT:') != -1:
    index = r[begin:-1].find('VALID AT:')
    lst.append(index+begin)
    begin = begin+index+1

timeUS = []
typeUS = []
latUS = []
lngUS = []
radiusUS = []
pressureUS = []

localtime = time.localtime()
result = time.strftime("%Y-%m-%d %I:%M:%S %p", localtime)

nowstr = 'WARNING POSITION:'

now = r.find(nowstr)

year = result[0:4]
month = result[5:7]
day = r[now+22:now+24]
hour = r[now+24:now+26]
minute = r[now+26:now+28]

timeString = change_time_to_TW(year, month, day, hour, minute)

timeUS.append(timeString)
typeUS.append('Analysis')
latUS.append(r[now+39:now+43])
lngUS.append(r[now+45:now+50])
radiusUS.append(None)
pressureUS.append(None)

temp = 0

for i in lst:
    day = r[i+14:i+16]
    hour = r[i+16:i+18]
    minute = r[i+18:i+20]
    if(int(day) < temp):
        month = int(month)+1
    if int(month) > 12:
        month -= 12

    timeString = change_time_to_TW(year, month, day, hour, minute)
    temp = int(day)

    timeUS.append(timeString)

    typeUS.append('Forecast')
    latUS.append(r[i+26:i+30])
    lngUS.append(r[i+32:i+37])
    radiusUS.append(None)
    pressureUS.append(None)

df_US = pd.DataFrame({
    'time': timeUS, 'type': typeUS, 'lat': latUS, 'lng': lngUS, 'radius': radiusUS, 'pressure': pressureUS,
})


# CHINA
urlCN = "https://typhoon.slt.zj.gov.cn/Api/TyphoonInfo/"+fullyear+CNnum
htmlCN = requests.get(urlCN).text[1:-1]
dct = json.loads(htmlCN)

length = len(dct['points'])

timeCN = []
typeCN = []
latCN = []
lngCN = []
pressureCN = []
radiusCN = []

for i in range(len(dct['points'][length-1]['forecast'])):
    if (dct['points'][length-1]['forecast'][i]['tm'] == '中国'):
        oo = len(dct['points'][length-1]['forecast'][i]['forecastpoints'])
        for z in range(oo):
            timeCN.append(dct['points'][length-1]['forecast']
                          [i]['forecastpoints'][z]['time'])
            latCN.append(dct['points'][length-1]['forecast']
                         [i]['forecastpoints'][z]['lat'])
            lngCN.append(dct['points'][length-1]['forecast']
                         [i]['forecastpoints'][z]['lng'])
            pressureCN.append(dct['points'][length-1]['forecast']
                              [i]['forecastpoints'][z]['pressure'])
            if('radius' in dct):
                radiusCN.append(dct['points'][length-1]['forecast']
                                [i]['forecastpoints'][0]['radius'])
            else:
                radiusCN.append(None)
            if('type' in dct):
                typeCN.append(dct['points'][length-1]['forecast']
                              [i]['forecastpoints'][0]['type'])
            else:
                typeCN.append('Forecast')

typeCN[0] = 'Analysis'

df_CN = pd.DataFrame({
    'time': timeCN, 'type': typeCN, 'lat': latCN, 'lng': lngCN, 'radius': radiusCN, 'pressure': pressureCN,
})


# HONGKONG
urlHK = "https://typhoon.slt.zj.gov.cn/Api/TyphoonInfo/" + fullyear+HKnum
htmlHK = requests.get(urlHK).text[1:-1]
dct = json.loads(htmlHK)
length = len(dct['points'])

timeHK = []
typeHK = []
latHK = []
lngHK = []
pressureHK = []
radiusHK = []


for i in range(len(dct['points'][length-1]['forecast'])):
    if (dct['points'][length-1]['forecast'][i]['tm'] == '中国香港'):
        oo = len(dct['points'][length-1]['forecast'][i]['forecastpoints'])
        for z in range(oo):
            timeHK.append(dct['points'][length-1]['forecast']
                          [i]['forecastpoints'][z]['time'])
            latHK.append(dct['points'][length-1]['forecast']
                         [i]['forecastpoints'][z]['lat'])
            lngHK.append(dct['points'][length-1]['forecast']
                         [i]['forecastpoints'][z]['lng'])
            pressureHK.append(dct['points'][length-1]['forecast']
                              [i]['forecastpoints'][z]['pressure'])
            if('radius' in dct):
                radiusHK.append(dct['points'][length-1]['forecast']
                                [i]['forecastpoints'][0]['radius'])
            else:
                radiusHK.append(None)
            if('type' in dct):
                typeHK.append(dct['points'][length-1]['forecast']
                              [i]['forecastpoints'][0]['type'])
            else:
                typeHK.append('Forecast')

typeHK[0] = 'Analysis'

df_HK = pd.DataFrame({
    'time': timeHK, 'type': typeHK, 'lat': latHK, 'lng': lngHK, 'radius': radiusHK, 'pressure': pressureHK,
})

# JAPAN
urlJP = "https://www.jma.go.jp/bosai/typhoon/data/TC" + \
    halfyear+JPnum+"/specifications.json"
htmlJP = requests.get(urlJP).text
dct = json.loads(htmlJP)

timeJP = []
typeJP = []
latJP = []
lngJP = []
pressureJP = []
radiusJP = []

settime = dct[1]['validtime']['UTC']

year = settime[0:4]
month = settime[5:7]
day = settime[8:10]
hour = settime[11:13]
minute = settime[14:16]

timeString = change_time_to_TW(year, month, day, hour, minute)

typeJP.append(dct[1]['part']['en'][0:8])
latJP.append(dct[1]['position']['deg'][0])
lngJP.append(dct[1]['position']['deg'][1])
timeJP.append(timeString)
pressureJP.append(None)
radiusJP.append(None)

for i in dct[2:-1]:
    typeJP.append(i['part']['en'][0:8])
    latJP.append(i['position']['deg'][0])
    lngJP.append(i['position']['deg'][1])
    settimeJP = i['validtime']['UTC']
    year = settimeJP[0:4]
    month = settimeJP[5:7]
    day = settimeJP[8:10]
    hour = settimeJP[11:13]
    minute = settimeJP[14:16]

    timeString = change_time_to_TW(year, month, day, hour, minute)

    timeJP.append(timeString)
    radiusJP.append(i['probabilityCircleRadius']['km'])
    pressureJP.append(i['pressure'])

df_JP = pd.DataFrame({
    'time': timeJP, 'type': typeJP, 'lat': latJP, 'lng': lngJP, 'radius': radiusJP, 'pressure': pressureJP,
})

# KOREA
pathKR = 'C:/typhoon/chromedriver_win32/chromedriver'
driver = webdriver.Chrome(pathKR)
driver.get('https://www.kma.go.kr/eng/weather/typoon/typhoon_5days.jsp?tIdx=0')

driver.maximize_window()
time.sleep(30)

datas = driver.find_elements("xpath", '//*[@id="contents"]/table/tbody/tr')

time_data = []
type_data = []
lat_data = []
lng_data = []
radius_data = []
pressure_data = []

for data in datas[0:-1]:
    timeKR = data.find_element("xpath", './td[1]').get_attribute('textContent')
    year = timeKR[0:4]
    month = timeKR[5:7]
    day = timeKR[8:10]
    hour = timeKR[12:14]
    minute = timeKR[15:17]

    timeString = change_time_to_TW(year, month, day, hour, minute)

    latKR = data.find_element("xpath", './td[2]').get_attribute('textContent')
    lngKR = data.find_element("xpath", './td[3]').get_attribute('textContent')
    pressureKR = data.find_element(
        "xpath", './td[4]').get_attribute('textContent')
    radiusKR = data.find_element(
        "xpath", './td[12]').get_attribute('textContent')

    time_data.append(timeString)
    type_data.append(timeKR[18:-1])
    lat_data.append(latKR)
    lng_data.append(lngKR)
    radius_data.append(radiusKR)
    pressure_data.append(pressureKR)


time.sleep(5)
driver.quit()

df_KR = pd.DataFrame({
    'time': time_data, 'type': type_data, 'lat': lat_data, 'lng': lng_data, 'radius': radius_data, 'pressure': pressure_data,
})


# TW
path = 'C:/typhoon/chromedriver_win32/chromedriver'
driver = webdriver.Chrome(path)
driver.get('https://www.cwb.gov.tw/V8/C/P/Typhoon/TY_NEWS.html')

driver.maximize_window()
time.sleep(5)

datas = driver.find_elements("xpath", '//*[@id="collapse-B1"]/div/div[2]/ul')
nowdatas = driver.find_elements(
    "xpath", '//*[@id="collapse-B1"]/div/div[1]')

time_data = []
type_data = []
lat_data = []
lng_data = []
pressure_data = []
radius_data = []


for nowdata in nowdatas:
    timeTW = nowdata.find_element(
        "xpath", './p[1]').get_attribute('textContent')
    position = nowdata.find_element(
        "xpath", './ul[1]/li[1]').get_attribute('textContent')
    pressure = nowdata.find_element(
        "xpath", './ul[1]/li[4]').get_attribute('textContent')
    radius = nowdata.find_element(
        "xpath", './ul[1]/li[7]').get_attribute('textContent')

    timeTW = re.findall(r'-?\d+\.?\d*', timeTW)
    position = re.findall(r'-?\d+\.?\d*', position)
    pressure = re.findall(r'-?\d+\.?\d*', pressure)
    radius = re.findall(r'-?\d+\.?\d*', radius)

    time_data.append(str(timeTW[0]) +
                     '/'+str(timeTW[1])+'/'+str(timeTW[2]) +
                     ' '+str(timeTW[3]+':00'))
    type_data.append('Analysis')
    lat_data.append(position[0])
    lng_data.append(position[1])
    pressure_data.append(pressure[0])
    radius_data.append(None)

for data in datas:
    timeTW = data.find_element("xpath", './li[2]').get_attribute('textContent')
    position = data.find_element(
        "xpath", './li[3]').get_attribute('textContent')
    pressure = data.find_element(
        "xpath", './li[4]').get_attribute('textContent')
    radius = data.find_element("xpath", './li[7]').get_attribute('textContent')

    timeTW = re.findall(r'-?\d+\.?\d*', timeTW)
    position = re.findall(r'-?\d+\.?\d*', position)
    pressure = re.findall(r'-?\d+\.?\d*', pressure)
    radius = re.findall(r'-?\d+\.?\d*', radius)

    time_data.append('2022/'+str(timeTW[0]) +
                     '/'+str(timeTW[1])+' '+str(timeTW[2])+':00')
    type_data.append('forecast')
    lat_data.append(position[0])
    lng_data.append(position[1])
    pressure_data.append(pressure[0])
    radius_data.append(None)

time.sleep(5)
driver.quit()

df_TW = pd.DataFrame({
    'time': time_data, 'type': type_data, 'lat': lat_data, 'lng': lng_data, 'radius': radius_data, 'pressure': pressure_data,
})

storepath = 'C:/typhoon/NewPath/newpath.xlsx'
oldpath = auto_save_file('C:/typhoon/OldPath/path.xlsx')

if os.path.exists(storepath):
    os.rename(storepath, oldpath)

with pd.ExcelWriter(storepath) as writer:
    df.to_excel(writer, sheet_name='CODE',
                index=False, header=True)
    df_CN.to_excel(writer, sheet_name='CN',
                   index=False, header=True)
    df_HK.to_excel(writer, sheet_name='HK',
                   index=False, header=True)
    df_TW.to_excel(writer, sheet_name='TW',
                   index=False, header=True)
    df_US.to_excel(writer, sheet_name='US',
                   index=False, header=True)
    df_JP.to_excel(writer, sheet_name='JP',
                   index=False, header=True)
    df_KR.to_excel(writer, sheet_name='KR',
                   index=False, header=True)
