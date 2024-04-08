import rainFn
from PIL import Image
import numpy as np
import matplotlib.pyplot as plt
import time
import winctrl as wcon
from pathlib import Path
import pytesseract
import collections
import collections.abc
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
import pptxFn as pxn
from bs4 import BeautifulSoup
import re


def splitFn(inp, param):
    r = inp.split(param[0])
    if len(param[1]) == 0:
        result = r
    else:
        result = []
        for p in param[1]:
            result.append(r[p])
    return result


def numberFn(inp, param):
    r = re.findall(r"(?:\d*\.\d+|\d+)", inp)
    if len(param) == 0:
        result = r
    else:
        result = []
        for p in param:
            result.append(r[p])
    return result


mappingKeys = {
    '中心位置': ['center', numberFn, []],
    '過去移動方向': ['direction', splitFn, [' ', [1]]],
    '中心氣壓': ['pressure', numberFn, [0]],
    '時速': ['speed', numberFn, [0]],
    '瞬間最大陣風': ['gust', numberFn, [0]],
    '中心最大風速': ['windspeed', numberFn, [0]],
    '七級風暴風半徑': ['radius7', numberFn, [0]],
    '十級風暴風半徑': ['radius10', numberFn, [0]],
    '緩慢移動': ['speed', numberFn, []],
    # '月': ['time', numberFn, []],
    '70%機率半徑': ['radius70p', numberFn, [1]],
}

# 下載颱風資料
downloadPath = Path('C:/Users/AQUA/Downloads')
chromepath = r"C:/Program Files/Google/Chrome/Application/chrome.exe"
filename = 'typhoon.html'
url = 'https://www.cwb.gov.tw/V8/C/P/Typhoon/TY_NEWS.html'

# wcon.downloadWebPage(url, chromepath, downloadPath, filename, loadwait=30, savewait=30)

# 整理下載好的資料, 內含頁面中所有颱風的資枓
with open(downloadPath / filename, encoding="utf-8") as fp:
    soup = BeautifulSoup(fp, "html.parser")
data = soup.find('div', attrs={
                 'class': 'panel-group acc-v1 vision_1 margin-top-30', 'id': 'accordion-2'})
# data = data.prettify()
spans = [span.text for span in data.find_all('span')]  # 所有颱風的基本資料
# 颱風資料應有的欄位
infos = ['time', 'type', 'description', 'longitude', 'latitude', 'speed', 'direction',
         'pressure', 'windspeed', 'gust', 'radius7', 'radius10', 'radius70p']
typInfo = {}  # 所有的颱風資料
typType = {}  # 颱風的描述, 如: 輕度颱風, 熱帶性低氣壓等, 先記起來, 等一下再填入typInfo中
for i in range(len(spans)//4):
    name = spans[i*4+1]
    typInfo[name] = {info: [] for info in infos}
    typType[name] = spans[i*4]
uls = data.find_all('ul')  # 所有颱風的現況與預測路徑
ps = [p.text for p in data.find_all('p')]  # 所有颱風的現況與預測路徑標題, 對應每個現況與預測路徑
# 把所有欄位資料都設為None
typNames = iter(list(typInfo.keys()))
for p in ps:
    nRec = 0
    if '預測' not in p:
        name = next(typNames)
    for info in infos:
        typInfo[name][info].append(None)

# 找出所有颱風, 並填入對應的欄位
typNames = iter(list(typInfo.keys()))
for i, ul in enumerate(uls):
    lis = [li.text for li in ul.find_all('li')]
    if lis[0].startswith('中心'):  # 最近颱風實際資料
        nRec = 0
        name = next(typNames)
        lat, lon = re.findall(r"[-+]?(?:\d*\.\d+|\d+)", lis[0])
        dates = numberFn(ps[i], [])
        yr = dates[0]  # 預測資料沒有年的資料
        typInfo[name]['description'][nRec] = typType[name]
        typInfo[name]['type'][nRec] = 'analysis'
        typInfo[name]['time'][nRec] = f'{dates[0]}/{dates[1]}/{dates[2]} {dates[3]}:00'
        typInfo[name]['longitude'][nRec] = lon
        typInfo[name]['latitude'][nRec] = lat
        for li in lis[1:]:
            for k in list(mappingKeys.keys()):
                if k in li:
                    val = mappingKeys[k][1](li, mappingKeys[k][2])
                    typInfo[name][mappingKeys[k][0]][nRec] = val[0]
        nRec = nRec + 1
    else:  # 颱風預測資料
        typInfo[name]['type'][nRec] = 'forecast'
        dates = numberFn(lis[1], [])
        typInfo[name]['time'][nRec] = f'{yr}/{dates[0]}/{dates[1]} {dates[2]}:00'
        move = splitFn(lis[0], [' ', []])
        typInfo[name]['direction'][nRec] = move[0]
        typInfo[name]['speed'][nRec] = move[1]
        for m in move[1:]:
            try:
                val = int(m)
                typInfo[name]['speed'][nRec] = m
            except Exception as ex:
                pass
        lat, lon = numberFn(lis[2], [])
        typInfo[name]['longitude'][nRec] = lon
        typInfo[name]['latitude'][nRec] = lat
        for li in lis[3:]:
            for k in list(mappingKeys.keys()):
                if k in li:
                    val = mappingKeys[k][1](li, mappingKeys[k][2])
                    typInfo[name][mappingKeys[k][0]][nRec] = val[0]
        nRec = nRec + 1

# 計算颱風與嘉義巿之間的相對位置關係: 距離、方位
chiayi = pxn.appoachChiayi(typInfo)

# 將資料填入pptx
for typName in list(typInfo.keys()):
    templatePPTX = "./templateR3.pptx"
    blankPPTX = './blank.pptx'
    pptT = pptx.Presentation(templatePPTX)
    pptB = pptx.Presentation(blankPPTX)
    allImg = {}
    typ = typInfo[typName]
    chi = chiayi[typName]
    velo = (chi['distance'][1] - chi['distance'][0]) / 6
    appVelo = f'{np.abs(velo): 4.3f}'
    if velo > 0:
        approach = '遠離'
    elif velo == 0:
        approach = '滯留'
    else:
        approach = '接近'
    texts = [
        typName,
        typ['time'][0][8:10],
        typ['time'][0][11:13],
        typ['latitude'][0],
        typ['longitude'][0],
        chi['compass'][0],
        str(chi['distance'][0]),
        appVelo,
        approach,
        typ['speed'][1],
        typ['direction'][0],
        typ['pressure'][0],
        typ['windspeed'][0],
        pxn.getWindScale(float(typ['windspeed'][0])),
        typ['gust'][0],
        pxn.getWindScale(float(typ['gust'][0])),
        typ['radius7'][0],
        typ['radius10'][0]
    ]
    nS = 1
    pxn.replaceText(nS, pptB.slides, texts)
    pptB.save(f"Slide4_{typName}.pptx")
