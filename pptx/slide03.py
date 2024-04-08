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

url = 'https://npd.cwb.gov.tw/NPD/products_display/product?menu_index=1'
filename = 'numerical.html'

# wcon.downloadWebPage(url, chromepath, downloadPath,
#                      filename, loadwait=30, savewait=30)
files = os.listdir(downloadPath+filename[:-5]+'_files')
dates = list(filter(lambda x: x.endswith('.gif'), files))[0].split('_')[-2]
tms = list(range(0, 45, 6))  # + list(range(90, 121, 6))
gifs = [f'WWW_WRF_F2_{tm:03d}_{dates}_B622A01X2B.gif' for tm in tms]
for fn, tm in zip(gifs, tms):
    url = f'https://npd.cwb.gov.tw/NPD/irisme_data/grapher/gifdir/Nwww/WRFM04/{dates}/WWW_WRF_F2_{tm:03d}_{dates}_B622A01X2B.gif'
    # wcon.downloadWebPage(url, chromepath, downloadPath,
    #                      fn, loadwait=5, savewait=5)
print('---------------------------------------------------------------------------')
# 存成gif動畫
# imgs = (Image.open('C:/Users/a9536/Downloads/'+gif) for gif in gifs)
# print(list(imgs)[0])
# img = next(imgs)
# img.save(fp='slide03.gif', format='GIF', append_images=imgs,
#          save_all=True, duration=1000, loop=0)

# 存成靜態圖
templatePPTX = 'C:/work/typhoon/pptx/templateR3.pptx'
blankPPTX = 'C:/work/typhoon/pptx/blank.pptx'
pptT = pptx.Presentation(templatePPTX)
pptB = pptx.Presentation(blankPPTX)
nS = 3
allImg = pxn.getImgInfo(nS, pptT.slides)
# print(allImg)
t0 = datetime.strptime('20'+dates, '%Y%m%d%H')
for i in range(0, 43, 6):
    frm = (i//6)+1
    # print(frm)
    row = ((i//6)//4)
    col = ((i//6) % 4)
    ti = t0+timedelta(hours=i)
    img = downloadPath + f'/WWW_WRF_F2_{i:03d}_{dates}_B622A01X2B.gif'
    pic = pptB.slides[nS-1].shapes.add_picture(img,
                                               allImg[frm]['left'],
                                               allImg[frm]['top'],
                                               width=allImg[frm]['width'],
                                               height=allImg[frm]['height'])
    timeBBox = pptB.slides[nS-1].shapes.add_textbox(
        Inches(0.8+(col*2.88)), Inches(1.6+(row*3.5)), Inches(1.4), Inches(0.4))
    timeBBox.fill.solid()
    timeBBox.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
    timeBFrame = timeBBox.text_frame
    timeBFrame.vertical_anchor = MSO_ANCHOR.TOP
    timeBPara = timeBFrame.paragraphs[0]
    timeBPara.text = ti.strftime("%m%d %H%M")
    timeBPara.font.size = Pt(20)
    timeBPara.font.bold = True
    timeBPara.font.color.rgb = RGBColor(0, 0, 255)
pptB.save(f"Slide03.pptx")
