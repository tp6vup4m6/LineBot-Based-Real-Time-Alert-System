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

url1 = 'https://watch.ncdr.nat.gov.tw/watch_typhoon'
url2 = 'https://www.cwb.gov.tw/V8/C/P/Typhoon/TY_NEWS.html'
filename1 = 'numerical1.html'
filename2 = 'numerical2.html'
# wcon.downloadWebPage(url1, chromepath, downloadPath,
#                      filename1, loadwait=30, savewait=30)
# wcon.downloadWebPage(url2, chromepath, downloadPath,
#                      filename2, loadwait=30, savewait=30)
files1 = os.listdir(downloadPath+filename1[:-5]+'_files')
files2 = os.listdir(downloadPath+filename2[:-5]+'_files')

pic1 = list(filter(lambda x: x.endswith('.png'), files1))[19]
url = f'https://watch.ncdr.nat.gov.tw/00_Wxmap/5F14_CWB_OFFICITAL_TRACK/202308/'+pic1
# wcon.downloadWebPage(url, chromepath, downloadPath,
#                      pic1, loadwait=5, savewait=5)

pic2 = list(filter(lambda x: x.endswith('.png'), files2))[1]
url = f'https://www.cwb.gov.tw/Data/typhoon/TY_NEWS/'+pic2
# wcon.downloadWebPage(url, chromepath, downloadPath,
#                      pic2, loadwait=5, savewait=5)

templatePPTX = 'C:/work/typhoon/pptx/templateR3.pptx'
blankPPTX = 'C:/work/typhoon/pptx/blank.pptx'
pptT = pptx.Presentation(templatePPTX)
pptB = pptx.Presentation(blankPPTX)
nS = 2
allImg = pxn.getImgInfo(nS, pptT.slides)
# print(allImg)

img1 = downloadPath + f'/'+pic1
pic11 = pptB.slides[nS-1].shapes.add_picture(img1,
                                             allImg[1]['left'],
                                             allImg[1]['top'],
                                             width=allImg[1]['width'],
                                             height=allImg[1]['height'])

img2 = downloadPath + f'/'+pic2
pic12 = pptB.slides[nS-1].shapes.add_picture(img2,
                                             allImg[2]['left'],
                                             allImg[2]['top'],
                                             width=allImg[2]['width'],
                                             height=allImg[2]['height'])

pptB.save(f"Slide02.pptx")
