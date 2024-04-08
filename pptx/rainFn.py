import numpy as np
from PIL import Image

RainScale = {
    '~0.0 mm': (237, 249, 254, 255, 0.0),
    '<1.0 mm': (194, 194, 194, 255, 0.5),
    '<2.0 mm': (156, 252, 255, 255, 1.5),
    '<5.0 mm': (3, 200, 255, 255, 3.5),
    '< 10 mm': (5, 155, 255, 255, 7.5),
    '< 15 mm': (3, 99, 255, 255, 12.5),
    '< 20 mm': (5, 153, 2, 255, 17.5),
    '< 30 mm': (57, 255, 3, 255, 25.0),
    '< 40 mm': (255, 251, 3, 255, 35.0),
    '< 50 mm': (255, 200, 0, 255, 45.0),
    '< 70 mm': (255, 149, 0, 255, 60.0),
    '< 90 mm': (255, 0, 0, 255, 80.0),
    '<110 mm': (204, 0, 0, 255, 100.0),
    '<130 mm': (153, 0, 0, 255, 120.0),
    '<150 mm': (150, 0, 0, 255, 140.0),
    '<200 mm': (201, 0, 204, 255, 175.0),
    '<300 mm': (251, 0, 255, 255, 250.0),
    '>300 mm': (253, 201, 255, 255, 300.0),
}


def ChiayiMaskW(id):
    if id == 1:
        height = 1500
        width = 1245
        mask = np.zeros(shape=(height, width, 4))
        mask[863:867, 528:534] = (1, 1, 1, 1)
        mask[867:871, 517:534] = (1, 1, 1, 1)
        mask[871:872, 517:534] = (1, 1, 1, 1)
        mask[872:878, 521:534] = (1, 1, 1, 1)
        mask[878, 517:534] = (1, 1, 1, 1)
        mask[879, 531:534] = (1, 1, 1, 1)
        mask[880:883, 531:534] = (1, 1, 1, 1)
        mask[883:886, 531:534] = (1, 1, 1, 1)
        pixel = []
        for y in range(850, 890):
            for x in range(510, 560):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel
    if id == 2:
        height = 780
        width = 700
        mask = np.zeros(shape=(height, width, 4))
        mask[434, 243] = (1, 1, 1, 1)
        mask[435, 241:244] = (1, 1, 1, 1)
        mask[436, 241:245] = (1, 1, 1, 1)
        mask[437, 244] = (1, 1, 1, 1)
        pixel = []
        for y in range(429, 443):
            for x in range(238, 251):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel
    if id == 3:
        height = 642
        width = 315
        mask = np.zeros(shape=(height, width, 4))
        mask[338, 63:66] = (1, 1, 1, 1)
        mask[339:344, 61:67] = (1, 1, 1, 1)
        mask[344, 65:68] = (1, 1, 1, 1)
        mask[345, 66] = (1, 1, 1, 1)
        pixel = []
        for y in range(330, 355):
            for x in range(55, 80):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel


def ChiayiMaskE(id):
    if id == 1:
        height = 1500
        width = 1245
        mask = np.zeros(shape=(height, width, 4))
        mask[863:867, 534:539] = (1, 1, 1, 1)
        mask[867:871, 534:546] = (1, 1, 1, 1)
        mask[871:872, 534:553] = (1, 1, 1, 1)
        mask[872:878, 534:553] = (1, 1, 1, 1)
        mask[878, 534:553] = (1, 1, 1, 1)
        mask[879, 534:539] = (1, 1, 1, 1)
        mask[879, 549:553] = (1, 1, 1, 1)
        mask[880:883, 534:539] = (1, 1, 1, 1)
        mask[880:883, 549:553] = (1, 1, 1, 1)
        mask[883:886, 534:535] = (1, 1, 1, 1)
        pixel = []
        for y in range(850, 890):
            for x in range(510, 560):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel
    if id == 2:
        height = 780
        width = 700
        mask = np.zeros(shape=(height, width, 4))
        mask[434, 244] = (1, 1, 1, 1)
        mask[435, 244:247] = (1, 1, 1, 1)
        mask[436, 245:248] = (1, 1, 1, 1)
        pixel = []
        for y in range(429, 443):
            for x in range(238, 251):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel
    if id == 3:
        height = 642
        width = 315
        mask = np.zeros(shape=(height, width, 4))
        mask[339, 71:73] = (1, 1, 1, 1)
        mask[340, 69:75] = (1, 1, 1, 1)
        mask[341:343, 69:76] = (1, 1, 1, 1)
        mask[343, 70:76] = (1, 1, 1, 1)
        mask[344, 73:75] = (1, 1, 1, 1)
        mask[345, 74] = (1, 1, 1, 1)
        pixel = []
        for y in range(330, 355):
            for x in range(55, 80):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel


def ChiayiMask(id):
    if id == 1:
        height = 1500
        width = 1245
        mask = np.zeros(shape=(height, width, 4))
        mask[863:867, 528:539] = (1, 1, 1, 1)
        mask[867:871, 517:546] = (1, 1, 1, 1)
        mask[871:872, 517:553] = (1, 1, 1, 1)
        mask[872:878, 521:553] = (1, 1, 1, 1)
        mask[878, 517:553] = (1, 1, 1, 1)
        mask[879, 531:539] = (1, 1, 1, 1)
        mask[879, 549:553] = (1, 1, 1, 1)
        mask[880:883, 531:539] = (1, 1, 1, 1)
        mask[880:883, 549:553] = (1, 1, 1, 1)
        mask[883:886, 531:535] = (1, 1, 1, 1)
        pixel = []
        for y in range(850, 890):
            for x in range(510, 560):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel
    if id == 2:
        height = 780
        width = 700
        mask = np.zeros(shape=(height, width, 4))
        mask[434, 243:245] = (1, 1, 1, 1)
        mask[435, 241:247] = (1, 1, 1, 1)
        mask[436, 241:248] = (1, 1, 1, 1)
        mask[437, 244] = (1, 1, 1, 1)
        pixel = []
        for y in range(429, 443):
            for x in range(238, 251):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel
    if id == 3:
        height = 642
        width = 315
        mask = np.zeros(shape=(height, width, 4))
        mask[339, 71:73] = (1, 1, 1, 1)
        mask[340, 69:75] = (1, 1, 1, 1)
        mask[341:343, 69:76] = (1, 1, 1, 1)
        mask[343, 70:76] = (1, 1, 1, 1)
        mask[344, 73:75] = (1, 1, 1, 1)
        mask[345, 74] = (1, 1, 1, 1)
        mask[338, 63:66] = (1, 1, 1, 1)
        mask[339:344, 61:67] = (1, 1, 1, 1)
        mask[344, 65:68] = (1, 1, 1, 1)
        mask[345, 66] = (1, 1, 1, 1)
        pixel = []
        for y in range(330, 355):
            for x in range(55, 80):
                if np.allclose(mask[y, x], np.array((1, 1, 1, 1))):
                    pixel.append((y, x))
        return mask, pixel


def calRainDepth(img, masking, id):
    npImg = np.asarray(img)
    mask, pixel = masking(id)
    npImg = mask*npImg
    depth = 0
    for p in pixel:
        simMin = 99999999
        data = npImg[p[0], p[1]][:3]
        for scale in RainScale.keys():
            level = RainScale[scale][:3]
            sim = np.sum(np.abs(data-level))
            if sim < simMin:
                simMin = sim
                scaleKey = scale
        depth = depth + RainScale[scaleKey][4]
    errMin = 999999
    depth = depth/len(pixel)
    for k in RainScale.keys():
        err = np.abs(RainScale[k][-1]-depth)
        if err < errMin:
            errMin = err
            level = k
    return level
