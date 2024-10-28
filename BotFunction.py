import pyautogui as Driver
import autoit as At
import os
import time
import ObjPool as Pool
import pandas as pd
import logging as Log

def ClickOnScreen(ImageLocation,ButtonORType):

    res = Driver.locateOnScreen(ImageLocation,confidence=0.9)
    ImageCoordinates = Driver.center(res)
    Driver.moveTo(ImageCoordinates)
    print(ImageCoordinates)
    time.sleep(0.5)
    if(ButtonORType=="Left"):
        Driver.leftClick()
    elif(ButtonORType=="Right"):
        Driver.rightClick()
    elif(ButtonORType=="Double"):
        Driver.doubleClick()
    else:
        print("Invalid Click Button")

def ClickRelOnScreen(ImageLocation,ManualX,ManualY,ButtonORType):
    res = Driver.locateOnScreen(ImageLocation,confidence=0.9)
    ImageCoordinates = Driver.center(res)
    Driver.moveTo(ImageCoordinates)
    Driver.moveRel(ManualX,ManualY)
    time.sleep(0.5)
    if(ButtonORType=="Left"):
        Driver.leftClick()
    elif(ButtonORType=="Right"):
        Driver.rightClick()
    elif(ButtonORType=="Double"):
        Driver.doubleClick()
    else:
        print("Invalid Click Button")

def killApp(AppName):
    os.system("taskkill /f /im "+AppName)

def ExportFromSAP(SavePath,OverWrite):
    #Converting Overwrite Input to Lowercase
    OverWrite= str(OverWrite).lower()

    Exportbtn = Pool.ExportBtn
    SpreadOption = Pool.SpreadOption
    ClickOnScreen(Exportbtn ,"Left")
    time.sleep(1)
    Driver.press("S")
    Driver.hotkey("enter")
    time.sleep(2)
    i = 0
    while i<=10:
        SpdOption = At.win_exists("Save As")
        if SpdOption ==1:
            break
        Driver.click()
        Driver.press("S")
        Driver.hotkey("enter")
        time.sleep(10)
        try:
            At.win_active("Save As")
        except Exception:
            print("Trying to Find Save As window")
        i =i+1
         

    At.win_wait_active("Save As",10000)
    At.control_click("Save As","Edit1")
    At.control_set_text("Save As","Edit1",SavePath)
    At.control_click("Save As","Button2")
    
    #Invoking OverWrite the exisiting file if it is marked yes
    if OverWrite =="yes":
        time.sleep(2)
        try:
            ConfirmSaveWin=At.win_exists("Confirm Save As")
            if ConfirmSaveWin ==1:
                At.control_click("Confirm Save As","Button1")
        except:
            None
    
    #Allowing Security to Save the file
    try: 
        for _ in range(2):
            time.sleep(1)
            At.win_wait_active("SAP GUI Security",30)
            time.sleep(1)
            At.control_click("SAP GUI Security","Button2")
    except:
        None

def KillWins(WindowName):
    At.win_kill(WindowName)
    At.win_wait_close(WindowName,30)

def MergerExcel(Excel1,Excel2,Destination):
    df1 = pd.read_excel(Excel1)
    df2 = pd.read_excel(Excel2,header= None)
    merged_df = pd.concat([df1, df2])
    merged_df.to_excel(Destination,index=False)

Log.basicConfig(filename='Process.log', level= Log.INFO , format= "%(asctime)s %(levelname)s %(message)s")
def TrackBot(Msg):
    Log.info(Msg)

def JustSaveFromSAP(SavePath,OverWrite):
    #Converting Overwrite Input to Lowercase
    OverWrite= str(OverWrite).lower()
    time.sleep(2)
    i = 0
    while i<=10:
        SpdOption = At.win_exists("Save As")
        if SpdOption ==1:
            break
        Driver.click()
        Driver.press("S")
        Driver.hotkey("enter")
        time.sleep(10)
        try:
            At.win_active("Save As")
        except Exception:
            print("Trying to Find Save As window")
        i =i+1
         

    At.win_wait_active("Save As",10000)
    At.win_active("Save As")
    At.control_click("Save As","Edit1")
    At.control_set_text("Save As","Edit1",SavePath)
    time.sleep(1)
    At.control_click("Save As","Button2")
    time.sleep(1)
    At.control_click("Save As","Button2")
    
    #Invoking OverWrite the exisiting file if it is marked yes
    
    if OverWrite =="yes":
        time.sleep(2)
        try:
            ConfirmSaveWin=At.win_exists("Confirm Save As")
            if ConfirmSaveWin ==1:
                At.control_click("Confirm Save As","Button1")
        except:
            None
    
    #Allowing Security to Save the file
    try: 
        for _ in range(2):
            time.sleep(1)
            At.win_wait_active("SAP GUI Security",30)
            time.sleep(1)
            At.control_click("SAP GUI Security","Button2")
    except:
        None

def SelectWin(WinName,Timeout): 
    At.win_wait(WinName,30)
    At.win_activate(WinName)


