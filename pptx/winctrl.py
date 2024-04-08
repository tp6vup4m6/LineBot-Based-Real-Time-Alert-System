import win32api
import win32gui
import win32con
from win32con import WM_INPUTLANGCHANGEREQUEST
import time
import ctypes
import subprocess
import os
import pyperclip as pc
from pathlib import Path

caps = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*()_+{}:<>?~'
keys = {
    'backspace': 0x08, 'tab': 0x09, 'clear': 0x0C, 'enter': 0x0D,
    'shift': 0x10, 'ctrl': 0x11, 'alt': 0x12, 'pause': 0x13,
    'caps_lock': 0x14, 'esc': 0x1B, 'spacebar': 0x20,
    'page_up': 0x21, 'page_down': 0x22,
    'end': 0x23, 'home': 0x24,
    'left_arrow': 0x25, 'up_arrow': 0x26, 'right_arrow': 0x27, 'down_arrow': 0x28,
    'select': 0x29, 'print': 0x2A, 'execute': 0x2B, 'print_screen': 0x2C,
    'ins': 0x2D, 'del': 0x2E, 'help': 0x2F,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45,
    'f': 0x46, 'g': 0x47, 'h': 0x48, 'i': 0x49, 'j': 0x4A,
    'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E, 'o': 0x4F,
    'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54,
    'u': 0x55, 'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59,
    'z': 0x5A,
    'A': 0x41, 'B': 0x42, 'C': 0x43, 'D': 0x44, 'E': 0x45,
    'F': 0x46, 'G': 0x47, 'H': 0x48, 'I': 0x49, 'J': 0x4A,
    'K': 0x4B, 'L': 0x4C, 'M': 0x4D, 'N': 0x4E, 'O': 0x4F,
    'P': 0x50, 'Q': 0x51, 'R': 0x52, 'S': 0x53, 'T': 0x54,
    'U': 0x55, 'V': 0x56, 'W': 0x57, 'X': 0x58, 'Y': 0x59,
    'Z': 0x5A,
    'numpad_0': 0x60, 'numpad_1': 0x61, 'numpad_2': 0x62,
    'numpad_3': 0x63, 'numpad_4': 0x64, 'numpad_5': 0x65,
    'numpad_6': 0x66, 'numpad_7': 0x67, 'numpad_8': 0x68,
    'numpad_9': 0x69,
    'multiply_key': 0x6A, 'add_key': 0x6B, 'separator_key': 0x6C,
    'subtract_key': 0x6D, 'decimal_key': 0x6E, 'divide_key': 0x6F,
    'F1': 0x70, 'F2': 0x71, 'F3': 0x72, 'F4': 0x73, 'F5': 0x74,
    'F6': 0x75, 'F7': 0x76, 'F8': 0x77, 'F9': 0x78, 'F10': 0x79,
    'F11': 0x7A, 'F12': 0x7B, 'F13': 0x7C, 'F14': 0x7D, 'F15': 0x7E,
    'F16': 0x7F, 'F17': 0x80, 'F18': 0x81, 'F19': 0x82, 'F20': 0x83,
    'F21': 0x84, 'F22': 0x85, 'F23': 0x86, 'F24': 0x87,
    'num_lock': 0x90, 'scroll_lock': 0x91,
    'left_shift': 0xA0, 'right_shift ': 0xA1, 'left_control': 0xA2, 'right_control': 0xA3,
    'left_menu': 0xA4, 'right_menu': 0xA5,
    'browser_back': 0xA6, 'browser_forward': 0xA7, 'browser_refresh': 0xA8,
    'browser_stop': 0xA9, 'browser_search': 0xAA, 'browser_favorites': 0xAB,
    'browser_start_and_home': 0xAC,
    'volume_mute': 0xAD, 'volume_Down': 0xAE, 'volume_up': 0xAF,
    'next_track': 0xB0, 'previous_track': 0xB1,
    'stop_media': 0xB2, 'play/pause_media': 0xB3,
    'start_mail': 0xB4, 'select_media': 0xB5,
    'start_application_1': 0xB6, 'start_application_2': 0xB7,
    'attn_key': 0xF6, 'crsel_key': 0xF7, 'exsel_key': 0xF8,
    'play_key': 0xFA, 'zoom_key': 0xFB, 'clear_key': 0xFE,
    '+': 0xBB, '=': 0xBB, ',': 0xBC, '<': 0xBC, '-': 0xBD, '_': 0xBD,
    '.': 0xBE, '>': 0xBE, '/': 0xBF, '?': 0xBF, '`': 0xC0, '~': 0xC0,
    ';': 0xBA, ':': 0xBA, '[': 0xDB, '{': 0xDB, ']': 0xDD, '}': 0xDD,
    '&': 0x37, '\\': 0xDC, '|': 0xDC, "'": 0xDE, '"': 0xDE
}


# 模擬單一按鍵
def singleStrike(st):
    win32api.keybd_event(keys[st], 0, 0, 0)
    win32api.keybd_event(keys[st], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(1)


# 模擬control組合鍵
def ctrlStrike(st):
    win32api.keybd_event(keys['left_control'], 0, 0, 0)
    win32api.keybd_event(keys[st], 0, 0, 0)
    win32api.keybd_event(keys[st], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(keys['left_control'], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(1)


# 模擬shift組合鍵, 大小寫也由此函式處理
def shftStrike(st):
    win32api.keybd_event(keys['shift'], 0, 0, 0)
    win32api.keybd_event(keys[st], 0, 0, 0)
    win32api.keybd_event(keys[st], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(keys['shift'], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(1)


# 一次輸入許多按鍵
# 雖然予許一次給定一個字串, 但若有大小寫混合時, 會無法處理
# 所以還是一次處理一個字元, 若是大寫就呼叫shift組合鍵
def multiStrike(st):
    for c in st:
        if c in caps:
            shftStrike(c)
        else:
            singleStrike(c)
        time.sleep(0.2)
    time.sleep(1)


# 模擬mouse移到某個位置, 按左鍵
def moveMouseLClick(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # left down
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # left up
    time.sleep(0.2)


# 模擬mouse移到某個位置, 按右鍵
def moveMouseRClick(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(
        win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)  # left down
    ctypes.windll.user32.mouse_event(
        win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)  # left up
    time.sleep(0.2)


# 廢棄, 但還留著
def downloadchrome(chromepath, url):
    subprocess.Popen([chromepath])
    time.sleep(2)
    ctrlStrike('l')
    singleStrike('backspace')
    # win32api.LoadKeyboardLayout('00000409', 1)
    urls = url.split(':')
    for i, u in enumerate(urls):
        multiStrike(u)
        if i < len(urls)-1:
            shftStrike(':')
    singleStrike('enter')
    time.sleep(2)
    singleStrike('enter')
    time.sleep(2)
    os.system("taskkill /im chrome.exe /f")


# 廢棄, 但還留著
def savechrome(chromepath, url):
    subprocess.Popen([chromepath])
    time.sleep(2)
    ctrlStrike('l')
    singleStrike('backspace')
    for u in url:
        if u in caps:
            shftStrike(u)
        else:
            singleStrike(u)
    singleStrike('enter')
    time.sleep(2)
    ctrlStrike('s')
    singleStrike('tab')
    singleStrike('tab')
    singleStrike('tab')
    singleStrike('enter')
    time.sleep(10)
    os.system("taskkill /im chrome.exe /f")


# 取得剪貼簿的內容, 回傳值為自串
def getClipboard():
    CF_TEXT = 1
    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    user32 = ctypes.windll.user32
    user32.GetClipboardData.restype = ctypes.c_void_p
    user32.OpenClipboard(0)
    try:
        if user32.IsClipboardFormatAvailable(CF_TEXT):
            data = user32.GetClipboardData(CF_TEXT)
            data_locked = kernel32.GlobalLock(data)
            text = ctypes.c_char_p(data_locked)
            value = text.value
            kernel32.GlobalUnlock(data_locked)
    finally:
        user32.CloseClipboard()
    try:
        val = value.decode("utf-8")
    except Exception as ex:
        val = value.decode("cp950")
    return val  # value是bytes資料型態, 要轉為字串


# 設定鍵盤語言, 現在只支援英文與預設的中文鍵盤
def setKeyboard(language='en'):
    code = {'en': 0x4090409, 'zh': 0x4040404}
    handle = win32gui.GetForegroundWindow()
    win32api.SendMessage(handle, WM_INPUTLANGCHANGEREQUEST, 0, code[language])


def getChrome():
    return r"C:/Program Files/Google/Chrome/Application/chrome.exe"


# 只載入網頁, 無後續動作
def loadWebPage(chrome, url):
    pc.copy(url)
    subprocess.Popen([chrome])
    time.sleep(2)
    ctrlStrike('l')
    singleStrike('backspace')
    ctrlStrike('v')
    # chrome會自動產生類似的網址, 如果以前有比現在網址長的網址記錄
    # chrome會自動把網址補成之前常用的網址,
    # 用delete把自動補進來的部分刪除
    singleStrike('del')
    singleStrike('enter')
    time.sleep(2)


# chromepath: chrome執行檔
# targetFolder: 欲存檔的資料匣(現在還沒有此功能, 必須是chrome預試的下載資料匣)
# filename: 欲存檔案名稱, 空字串為chrome預設值
# loadwait: 載入網頁等待時間
# savewait: 存檔等待時間
def downloadWebPage(url, chromepath, targetFolder, filename, loadwait=90, savewait=90):
    setKeyboard('en')
    loadWebPage(chromepath, url)
    time.sleep(loadwait)
    pc.copy(filename)
    ctrlStrike('s')
    time.sleep(2)
    ctrlStrike('v')
    folderfiles = os.listdir(targetFolder)
    if filename in folderfiles:
        singleStrike('enter')
        time.sleep(1)
        singleStrike('y')
        time.sleep(savewait)
    else:
        singleStrike('tab')
        singleStrike('tab')
        singleStrike('tab')
        singleStrike('enter')
        time.sleep(savewait)
    os.system("taskkill /im chrome.exe /f")  # 暴力關閉chrome, 但清的最乾淨, 還在測試其他方法
