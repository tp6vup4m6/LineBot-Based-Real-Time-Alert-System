from PIL import Image
import numpy as np

im = Image.open('C:/Users/AQUA/Downloads/satellite_files/LCC_IR1_CR_1000.jpg')
im_crop = im.crop((200, 300, 800, 700))
im_crop.save('C:/Users/AQUA/Downloads/satellite_files/LCC_IR1_CR_1000_crop.jpg', quality=100)
