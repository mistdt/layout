# -*- coding: utf-8 -*-
# Created by liuchang on 2020/3/13
# Copyright (c) 2020 liuchang. All rights reserved.
import win32gui, win32con
import pyautogui as pag
import os
import win32com.client
from io import BytesIO
from PIL import Image,ImageGrab
import win32clipboard as clip
import pyperclip
import datetime
import schedule as sc
class Outsidewindow():
    def __init__(self):
        self.__hwnd=None
        self.__Title=None
        self.__delay=1
        self.__Imagename=type(self).__name__ + ".png"
        if os.path.isfile(self.__Imagename):
            os.remove(self.__Imagename)
        self.__imaged=False

    @property
    def Hwnd(self):
        return self.__hwnd
    @property
    def Imaged(self):
        return self.__imaged
    @property
    def Delay(self):
        return  self.__delay
    @Delay.setter
    def Delay(self,delay):
        self.__delay=0 if delay is None or delay<=0 else delay
    @property
    def Imagename(self):
        return self.__Imagename




    def _window_enum_callback(self, hwnd, regex):

         # if str(win32gui.GetWindowText(hwnd)).lower().find(regex.lower())>=0:
         #     print("句柄={0},iswindow={1},enabled={2},winvisible={3}".format(hwnd, win32gui.IsWindow(hwnd),win32gui.IsWindowEnabled(hwnd),win32gui.IsWindowVisible(hwnd)))
         #     print(win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_SYSMENU  ==win32con.WS_SYSMENU )
            #'''Pass to win32gui.EnumWindows() to check all open windows'''

        if self.__hwnd is None and str(win32gui.GetWindowText(hwnd)).lower().find(regex.lower())>=0:
                 if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE) & win32con.WS_SYSMENU  ==win32con.WS_SYSMENU:
                     #print(win32gui.GetWindowText(hwnd))
                     self.__hwnd = hwnd

    def FindWindow(self,regex):
        self.__hwnd = None
        self.__Title=regex
        win32gui.EnumWindows(self._window_enum_callback, regex)
        if self.__hwnd is not None:
            return True
        else:
            return  False
    def ShowWindow(self,state=win32con.SW_SHOW):
        #shell = win32com.client.Dispatch("WScript.Shell")
        #shell.SendKeys('%')
        if self.IsLive()==False:
            raise Exception("窗口已经关闭")

        if win32gui.IsIconic(self.__hwnd):
            win32gui.ShowWindow(self.__hwnd,win32con.SW_RESTORE)

        if win32gui.GetForegroundWindow()!=self.__hwnd:
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.ShowWindow(self.__hwnd, state)
            win32gui.SetForegroundWindow(self.__hwnd)


    def IsLive(self):
        if self.Hwnd is None:
            return False
        else:
            return win32gui.IsWindow(self.__hwnd)
    def ISvisible(self):
        return win32gui.IsWindowVisible(self.__hwnd)
    def Rebindhwnd(self):
        self.__imaged = False
        return self.FindWindow(self.__Title)
    def CopyWindow(self,filename=None,delay=None):
        if delay is  None:
            delay = self.Delay
        if delay>0:
            pag.sleep(delay)
        if filename is None:
            filename=self.__Imagename
        a, b, c, d = win32gui.GetWindowRect(self.__hwnd)

        pag.screenshot(filename, (a, b, c - a, d - b))
        self.__Imagename = filename
        self.__imaged = True
    def ImagetoClippboard(self):
        if self.Imaged==False:
            raise Exception("未拷贝图像")


        img = Image.open(self.__Imagename)  # Image.open可以打开网络图片与本地图片。
        output = BytesIO()  # BytesIO实现了在内存中读写bytes
        img.convert("RGB").save(output, "BMP")  # 以RGB模式保存图像
        data = output.getvalue()[14:]
        output.close()
        clip.OpenClipboard()  # 打开剪贴板
        clip.EmptyClipboard()  # 先清空剪贴板
        clip.SetClipboardData(win32con.CF_DIB, data)  # 将图片放入剪贴板
        clip.CloseClipboard()
    def Locate(self,imagename,delay=None):
        if os.path.isfile(imagename)==False:
            raise Exception("图像搜索：未找到待搜索的图像原文件")
        if delay is None:
            delay=self.Delay
        if delay>0:
            pag.sleep(delay)
        box = pag.locateOnScreen(imagename)
        if box is None:
            return None

        return pag.center(box)
    def ClickWindow(self,point,dc=False):
        if dc==True:
            pag.doubleClick(point.x,point.y)
        else:
            pag.click(point.x,point.y)
    def SendMsg(self,str,keyenter=True,delay=None):
        if self.IsLive()==False:
            return False
        if delay is None:
            delay=self.Delay
        if  delay>0:
            pag.sleep(delay)
        pyperclip.copy(str)
        pag.hotkey('ctrl', 'v')
        if keyenter==True:
            if  delay>0:
                pag.sleep(delay)
            pag.press("enter")

i=0
def Job():
    print("开始执行任务")
    wx = Outsidewindow()
    wx.FindWindow("微信")
    tdx = Outsidewindow()
    tdx.FindWindow("招商证券智")

    if tdx.IsLive():
        tdx.ShowWindow()
        tdx.CopyWindow()
        #win32gui.ShowWindow(tdx.Hwnd,win32con.SW_MINIMIZE)
    else:
        print("未找到通达信")
        return sc.CancelJob()
    wx.ShowWindow()
    box = wx.Locate("weixin_search.png")
    if box is not  None:

        wx.ClickWindow(box)
        wx.SendMsg("老刘")
        pag.sleep(1)
        pag.click(box.x, box.y + 80)
        tdx.ImagetoClippboard()
        pag.hotkey('ctrl', 'v')
        pag.sleep(1)
        pag.press("enter")
        pag.sleep(1)

    else:
        return sc.CancelJob()
    if datetime.datetime.now().minute>35:
        return sc.CancelJob()







# if wx.IsLive():
#     wx.ShowWindow()
#     box = wx.Locate("weixin_search.png")
#     if box is not  None:
#         pag.click(box)
#         wx.SendMsg("老刘")
#         tdx.ImagetoClippboard()
#         pag.hotkey('ctrl', 'v')
#         pag.sleep(1)
#         pag.press("enter")

if __name__ == '__main__':

    # wx=Outsidewindow()
    #
    #

    #
    # wx.FindWindow("微信")
    # tdx=Outsidewindow()
    # tdx.FindWindow("招商证券智")
    # print(tdx.Hwnd)
    # if tdx.IsLive():
    #     tdx.ShowWindow()
    #     tdx.CopyWindow()
    # else:
    #     print("未找到通达信")
    #
    # if wx.IsLive():
    #     wx.ShowWindow()
    #     box = wx.Locate("weixin_search.png")
    #     if box is not  None:
    #         pag.click(box)
    #         wx.SendMsg("老刘")
    #         tdx.ImagetoClippboard()
    #         pag.hotkey('ctrl', 'v')
    #         pag.sleep(1)
    #         pag.press("enter")


    a=sc.Scheduler()

    c=a.every(5).seconds.do(Job)
    c.tag("lwc")
    #a.cancel_job()
    a.clear("lwc")
    sc.every()
    print(a.jobs)
    #a.run_pending()
    # a.cancel_job(c)
    # print(a.jobs)
    # while True:
    #     a.run_pending()
    #     if len(a.jobs) <= 0:
    #         print("任务终结")
    #         break
    #     pag.sleep(1)
