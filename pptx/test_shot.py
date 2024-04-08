import pyscreenshot as ImageGrab

# 擷取指定範圍畫面（影像大小為 500x500 像素）
im = ImageGrab.grab(
    bbox=(20,   # X1
          20,   # Y1
          520,   # X2
          520))  # Y2

# 儲存檔案
im.save("box.png")
