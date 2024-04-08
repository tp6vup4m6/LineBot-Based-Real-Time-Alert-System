import rainFn
from PIL import Image
import numpy as np
import matplotlib.pyplot as plt
import time
from datetime import datetime, timedelta
import winctrl as wcon
from pathlib import Path
import pytesseract
import collections
import collections.abc
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt, Inches
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.effect import ShadowFormat
import pptxFn as pxn
import os

downloadPath = 'C:/Users/a9536/Downloads/'
chromepath = r"C:/Program Files/Google/Chrome/Application/chrome.exe"

url = 'https://watch.ncdr.nat.gov.tw/watch_tfrain_fst'
filename = 'numerical.html'
wcon.downloadWebPage(url, chromepath, downloadPath,
                     filename, loadwait=30, savewait=30)
files = os.listdir(downloadPath+filename[:-5]+'_files')

tms = list(range(0, 6, 1))
date3 = list(filter(lambda x: x.endswith('.gif'), files))[
    0].split('_')[-3][:-4]
date4 = list(filter(lambda x: x.endswith('.gif'), files))[
    0].split('_')[-3]

for i in tms:
    date1 = list(filter(lambda x: x.endswith('.gif'), files))[i].split('_')[-2]
    date2 = list(filter(lambda x: x.endswith('.gif'), files))[i].split('_')[-1]
    url = f'https://watch.ncdr.nat.gov.tw/00_Wxmap/5F4_CWBWRF_MultiModels/{date3}/{date4}/G01_{date4}_{date1}_{date2}'
    file = url[-26:]
    wcon.downloadWebPage(url, chromepath, downloadPath,
                         file, loadwait=5, savewait=5)

templatePPTX = 'C:/work/typhoon/pptx/templateR3.pptx'
blankPPTX = 'C:/work/typhoon/pptx/blank.pptx'
pptT = pptx.Presentation(templatePPTX)
pptB = pptx.Presentation(blankPPTX)
nS = 5
allImg = pxn.getImgInfo(nS, pptT.slides)
# print(allImg)

for i in range(0, 6, 1):
    frm = i+1
    date11 = list(filter(lambda x: x.endswith('.gif'), files))[
        i].split('_')[-2]
    date12 = list(filter(lambda x: x.endswith('.gif'), files))[
        i].split('_')[-1]

    img = downloadPath + f'/G01_{date4}_{date11}_{date12}'
    pic = pptB.slides[nS-1].shapes.add_picture(img,
                                               allImg[frm]['left'],
                                               allImg[frm]['top'],
                                               width=allImg[frm]['width'],
                                               height=allImg[frm]['height'])

pptB.save(f"Slide05.pptx")
