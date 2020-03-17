# -*- coding: utf-8 -*-
# Created by liuchang on 2020/3/12
# Copyright (c) 2020 liuchang. All rights reserved.


import pyautogui as pag


import  win32api
import win32gui
import win32clipboard as clip
import win32con
import  win32com.client
from io import BytesIO
from PIL import Image,ImageGrab
import  sys
import pyperclip
hwnd_title=dict()
def SendMsg(pict,act):
    box= pag.locateOnScreen(pict)
    if box is None:
        return None
    if act=="click":
        pag.click(pag.center(box))
    else:
        pag.doubleClick(pag.center(box))
    return box
def Show_Desktop():
    pass
def Show_Window(title):
    if len(hwnd_title.items())<=0:
        win32gui.EnumWindows(get_all_hwnd, 0)
    for h, t in hwnd_title.items():
        if t.lower().find(title.lower()) >= 0:
            print(h,t)
            老刘
            win32gui.ShowWindow(h,win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(h)
            return h
    return None



def get_all_hwnd(hwnd, mouse):
    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})
def setImage(filename):
    img = Image.open(filename)  # Image.open可以打开网络图片与本地图片。
    output = BytesIO()  # BytesIO实现了在内存中读写bytes
    img.convert("RGB").save(output, "BMP")  # 以RGB模式保存图像
    data = output.getvalue()[14:]
    output.close()
    clip.OpenClipboard() #打开剪贴板
    clip.EmptyClipboard()  #先清空剪贴板
    clip.SetClipboardData(win32con.CF_DIB, data)  #将图片放入剪贴板
    clip.CloseClipboard()
def Window_Snapshot(hwnd,filename):
    a, b, c, d = win32gui.GetWindowRect(hwnd)

    pag.screenshot(filename, (a, b, c - a, d - b))



if __name__=="__main__":
    h = Show_Window("微信")

    print(h)
   # pag.sleep(1)
  #  h = Show_Window("mybase")
    #pag.PAUSE = 1

   # h=Show_Window("招商证券")
   # print(h)
   # pag.sleep(1)
   # Window_Snapshot(h,"股票截图.png")
   # pag.sleep(1)
   #  i=1
   #  while i<=1:
   #      print(Show_Window("word"))
   #      pag.sleep(1)
   #      ret = SendMsg("weixin_search.png", "click")
   #
   #      print(ret)
   #     #  secs_between_keys = 0.1
   #     #
   #      pyperclip.copy("老婆")
   #      pag.hotkey('ctrl', 'v')
   #      pag.sleep(1)
   #
   #      pag.press("enter")
   #      box =pag.center(ret)
   #      pag.moveTo(box.x,box.y + 80)
   #      pag.click()
   #     #  pag.sleep(1)
   #
   #      pyperclip.copy("郭惠英大笨猪")
   #      pag.hotkey('ctrl', 'v')
   #      pag.press("enter")
   #      i+=1
   # # ret = SendMsg("群1.png", "click")
   #  setImage("股票截图.png")
   #  pag.sleep(1)
   #
   #  #pag.typewrite('老刘\n', interval=secs_between_keys)
   #  pag.sleep(1)
   #
   #  #pyperclip.copy("老流氓")
   #  pag.hotkey('ctrl', 'v')
   #
   #  pag.press("enter")