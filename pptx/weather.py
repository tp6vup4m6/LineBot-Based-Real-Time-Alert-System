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
import os

downloadPath = Path('C:/Users/a9536/Downloads')
chromepath = r"C:/Program Files/Google/Chrome/Application/chrome.exe"
url = 'https://www.cwb.gov.tw/V8/C/P/QPF.html'

wcon.loadWebPage(chromepath, url)
wcon.ctrlStrike('s')
time.sleep(1)
wcon.singleStrike('tab')
time.sleep(1)
wcon.singleStrike('tab')
time.sleep(1)
wcon.singleStrike('tab')
time.sleep(1)
wcon.singleStrike('enter')
time.sleep(10)
files = os.listdir(downloadPath / '定量降水預報_files')
dates = list(filter(lambda x: x.endswith('.gif'), files))[0].split('_')[-2]
