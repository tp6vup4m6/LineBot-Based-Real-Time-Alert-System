import collections
import collections.abc
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
import numpy as np
import pyproj

chiayiXY = (120.4, 23.53)
windScale = {
    '0': [0, 0.3], '1': [0.3, 1.6], '2': [1.6, 3.4], '3': [3.4, 5.5], '4': [5.5, 8.0],
    '5': [8.0, 10.8], '6': [10.8, 13.9], '7': [13.9, 17.2], '8': [17.2, 20.8], '9': [20.8, 24.5],
    '10': [24.5, 28.5], '11': [28.5, 32.7], '12': [32.7, 37.0], '13': [37.0, 41.5], '14': [41.5, 46.2],
    '15': [46.2, 51.0], '16': [51.0, 56.1], '17': [56.1, 61.3], '18': [61.3, 1000]
}


def compass(backAzimuth):

    dct = {
        '南南西': [-168.75, -146.25],
        '西南': [-146.25, -123.75],
        '西南西': [-123.75, -101.25],
        '西': [-101.25, -78.75],
        '西北西': [-78.75, -56.25],
        '西北': [-56.25, -33.75],
        '北北西': [-33.75, -11.25],
        '北': [-11.25, 11.25],
        '北北東': [11.25, 33.75],
        '東北': [33.75, 56.25],
        '東北東': [56.25, 78.75],
        '東': [78.75, 101.25],
        '東南東': [101.25, 123.75],
        '東南': [123.75, 146.25],
        '南南東': [146.25, 168.75],
    }
    val = '南'
    for c in list(dct.keys()):
        if backAzimuth >= dct[c][0] and backAzimuth < dct[c][1]:
            val = c
            break
    return val


def getImgInfo(nS, slides):
    dct = {}
    nImg = 1
    for s in slides[nS-1].shapes:
        if type(s) is pptx.shapes.picture.Picture:
            dct[nImg] = {}
            dct[nImg]['top'] = s.top
            dct[nImg]['left'] = s.left
            dct[nImg]['width'] = s.width
            dct[nImg]['height'] = s.height
            nImg = nImg+1
    return dct


def replaceText(nS, slides, texts):
    dct = {'TEXT'+f'{i+1:02d}': t for i, t in enumerate(texts)}
    for shape in slides[nS-1].shapes:
        for k in dct.keys():
            if shape.has_text_frame:
                # shape.text: 能看到整個text_frame的文字, 但可能只是view而已
                # 用shape.text變更文字並無效果
                if (shape.text.find(k)) != -1:
                    # 取得shape內的text_frame
                    # text_frame也有text屬性, 能看到文字, 但可能只是view而已
                    # 用text_frame.text變更文字並無效果
                    textFrame = shape.text_frame
                    # paragragphs: 能看到text_frame內所有段落
                    # paragraph.text: 能看到每個段落的文字, 但可能只是view而已
                    # 用paragraph.text變更文字並無效果
                    for paragraph in textFrame.paragraphs:
                        # run是一個段落中的再被細切的文字, 中文細切的原則未知
                        # 必須用run來替換文字才有效果
                        for run in paragraph.runs:
                            if k in run.text:
                                try:
                                    run.text = run.text.replace(k, dct[k])
                                except Exception as ex:
                                    run.text = run.text.replace(k, '無')


def toChiayi(info):
    geodesic = pyproj.Geod(ellps='WGS84')
    cX = chiayiXY[0]
    cY = chiayiXY[1]
    dct = {}
    for k in list(info.keys()):
        dct[k] = {}
        data = info[k]['time00']
        tX = float(data['center']['longitude'])
        tY = float(data['center']['latitude'])
        fwd, back, distance = geodesic.inv(tX, tY, cX, cY)
        dct[k]['distance'] = int(distance / 1000)
        dct[k]['compass'] = compass(back)
    return dct


def appoachChiayi(info):
    geodesic = pyproj.Geod(ellps='WGS84')
    cX = chiayiXY[0]
    cY = chiayiXY[1]
    dct = {}
    for k in list(info.keys()):
        dct[k] = {'distance': [], 'compass': []}
        data = info[k]
        for tX, tY in zip(data['longitude'], data['latitude']):
            tX = float(tX)
            tY = float(tY)
            fwd, back, distance = geodesic.inv(tX, tY, cX, cY)
            dct[k]['distance'].append(int(distance / 1000))
            dct[k]['compass'].append(compass(back))
    return dct


def getWindScale(wspd):
    for w in list(windScale.keys()):
        if wspd >= windScale[w][0] and wspd < windScale[w][1]:
            return w
