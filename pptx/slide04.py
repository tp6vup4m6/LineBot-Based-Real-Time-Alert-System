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
pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR\\tesseract.exe'

downloadPath = Path('C:/Users/a9536/Downloads/')
chromepath = r"C:/Program Files/Google/Chrome/Application/chrome.exe"
urls = [
    f'https://www.cwb.gov.tw/Data/fcst_img/QPF_ChFcstPrecip_6_{i:02d}.png' for i in range(6, 25, 6)]
files = [url[-25:] for url in urls]

# for url in urls:
#     wcon.downloadWebPage(url, chromepath, downloadPath,
#                          files, loadwait=5, savewait=5)

validTime = []
rainDepthW = []
rainDepthE = []
rainDepth = []
images = []
for f in files:
    img = Image.open(downloadPath / f)
    images.append(img)
    npImg = np.asarray(img)
    img = npImg[52:144, 246:900]
    img = Image.fromarray(np.uint8(img))
    ocr = pytesseract.image_to_string(img, lang="eng")
    dates = ocr.split('\n')
    validTime.append(dates[1][:13]+'-'+dates[1][-5:-3])
    # 1: 降雨圖高解析度; 2: 系集圖低解析度
    rainDepthW.append(rainFn.calRainDepth(npImg, rainFn.ChiayiMaskW, 1))
    rainDepthE.append(rainFn.calRainDepth(npImg, rainFn.ChiayiMaskE, 1))
    rainDepth.append(rainFn.calRainDepth(npImg, rainFn.ChiayiMask, 1))

# print(validTime)
# print(rainDepthW)
# print(rainDepthE)

mergedImg = Image.new('RGB', (4*1245, 1500))
for i, img in enumerate(images):
    mergedImg.paste(img, (1245*i, 0))
mergedImg.save(downloadPath / "mergedPrepDepth.gif", "GIF")
time.sleep(3)
# mergedImg.show()

templatePPTX = "C:/work/typhoon/pptx/templateR3.pptx"
blankPPTX = 'C:/work/typhoon/pptx/blank.pptx'
pptT = pptx.Presentation(templatePPTX)
pptB = pptx.Presentation(blankPPTX)
allImg = {}

# # 投影片4
nS = 4
allImg[nS] = pxn.getImgInfo(nS, pptT.slides)
nI = 1
pic = pptB.slides[nS-1].shapes.add_picture(
    f'./image/slide{nS}_{nI}.png', allImg[nS][nI]['left'], allImg[nS][nI]['top'], width=allImg[nS][nI]['width'])
pic.line.color.rgb = RGBColor(0, 76, 153)
pic.line.width = Pt(3)

pxn.replaceText(nS, pptB.slides, validTime+rainDepth)
pptB.save("Slide04.pptx")
