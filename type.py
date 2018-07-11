import win32api
import win32con
import pytesseract
from PIL import Image
from PIL import ImageGrab
import os, time
import pyautogui as pag

###Alt+tab 切换至金山打字通
win32api.keybd_event(18, 0, 0, 0)
win32api.keybd_event(9, 0, 0, 0)
win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
time.sleep(1)

index = 1
while True:
    ###截取当前单词存为图片
    im = ImageGrab.grab((245, 188, 1114, 216))
    im.save('D:\\pycode\\image\\a' + str(index) + '.jpg')
    # try:
    #     while True:
    #             print("Press Ctrl-C to end")
    #             x,y = pag.position() #返回鼠标的坐标
    #             posStr="Position:"+str(x).rjust(4)+','+str(y).rjust(4)
    #             print(posStr)#打印坐标
    #             time.sleep(0.2)
    #             os.system('cls')#清楚屏幕
    # except  KeyboardInterrupt:
    #     print('end....')
    # 245, 188  ,  1114, 216
    time.sleep(1)
    ###开始识别，调用OCR
    image = Image.open('D:\\pycode\\image\\a' + str(index) + '.jpg')
    code = pytesseract.image_to_string(image)
    for char in code:
        if char is ' ':
            win32api.keybd_event(32, 0, 0, 0)
            win32api.keybd_event(32, 0, win32con.KEYEVENTF_KEYUP, 0)
        elif char.isupper():
            win32api.keybd_event(16, 0, 0, 0)
            win32api.keybd_event(65 + ord(char) - ord('A'), 0, 0, 0)
            win32api.keybd_event(16, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(65 + ord(char) - ord('A'), 0, win32con.KEYEVENTF_KEYUP, 0)
        else:
            win32api.keybd_event(65 + ord(char) - ord('a'), 0, 0, 0)
            win32api.keybd_event(65 + ord(char) - ord('a'), 0, win32con.KEYEVENTF_KEYUP, 0)
    index += 1
    time.sleep(0.5)
# win32api.keybd_event(65,0,0,0)  #ctrl键位码是17
# win32api.keybd_event(65,0,win32con.KEYEVENTF_KEYUP,0) #释放按键
# win32api.keybd_event(66,0,0,0)  #v键位码是86
# win32api.keybd_event(66,0,win32con.KEYEVENTF_KEYUP,0)
time.sleep(1)
win32api.keybd_event(18, 0, 0, 0)
win32api.keybd_event(9, 0, 0, 0)
win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)
