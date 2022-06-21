from time import time
from turtle import setposition
from tsjPython.tsjCommonFunc import *
from config import *
import re
import time

import win32con
import win32gui
import win32com.client

def openTarget(reName):
    import os
    if searchRegexProcess(reName)==None:
        os.startfile(programPath)
        passPrint('try to Open Program')
    else:
        passPrint('Already Open Program')

def searchRegexProcess(reName):
    hWndList = []
    win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
    # print(hWndList)
    
    for hwnd in hWndList:
        clsname = win32gui.GetClassName(hwnd)
        title = win32gui.GetWindowText(hwnd)
        if(re.match("(.)*{}(.)*".format(reName), clsname, re.IGNORECASE) or re.match("(.)*{}(.)*".format(reName), title, re.IGNORECASE)):
            print("~~~~~~~~~~~")
            print(clsname)
            print(title)
            return hwnd
    return None

def setPositionSize(x,y,height,width,reName):
    count=0
    while count==0:
        hwnd=searchRegexProcess(reName)
        if hwnd!=None:
            count+=1
            print("height {} width {}".format(height, width))
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, x,
                                y, width, height, win32con.SWP_SHOWWINDOW)
            win32gui.BringWindowToTop(hwnd)
            # 先发送一个alt事件，否则会报错导致后面的设置无效：pywintypes.error: (0, 'SetForegroundWindow', 'No error message is available')
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(hwnd)
            passPrint('Open Program')
        else:
            errorPrint('wait Program to Open')
            time.sleep(1)
    

    # # 找出窗体的编号：
    # QQwin = win32gui . FindWindow("TXGuiFoundation", "QQ")  # 写类和标题，中间逗号隔开
    # # 控制窗体的位置和大小：
    # win32gui.SetWindowPos(QQwin, win32con.HWND_TOPMOST, 100,
    #                     100, 600, 600, win32con.SWP_SHOWWINDOW)
def calculateXY():
    scale=max(ipadResolutionRatioX/(DisplayX-fixOffset),ipadResolutionRatioY/(DisplayY-fixOffset))
    return [int(ipadResolutionRatioX/scale),int(ipadResolutionRatioY/scale)]

if __name__ == "__main__":
    openTarget(regexName)
    [allowableX,allowableY]=calculateXY()
    setPositionSize(positionsX,positionsY,allowableY,allowableX,regexName)
    