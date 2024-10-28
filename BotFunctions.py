import pyautogui as Driver
import autoit as At
import os
import time
#import ObjPool as Pool
import pandas as pd
import logging as Log


def ClickOnScreen(ImageLocation, ButtonORType):

    res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
    ImageCoordinates = Driver.center(res)
    Driver.moveTo(ImageCoordinates)
    print(ImageCoordinates)
    time.sleep(0.5)
    if (ButtonORType == "Left"):
        Driver.leftClick()
    elif (ButtonORType == "Right"):
        Driver.rightClick()
    elif (ButtonORType == "Double"):
        Driver.doubleClick()
    else:
        print("Invalid Click Button")


def ClickRelOnScreen(ImageLocation, ManualX, ManualY, ButtonORType):
    res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
    ImageCoordinates = Driver.center(res)
    Driver.moveTo(ImageCoordinates)
    Driver.moveRel(ManualX, ManualY)
    time.sleep(0.5)
    if (ButtonORType == "Left"):
        Driver.leftClick()
    elif (ButtonORType == "Right"):
        Driver.rightClick()
    elif (ButtonORType == "Double"):
        Driver.doubleClick()
    else:
        print("Invalid Click Button")


def killApp(AppName):
    os.system("taskkill /f /im "+AppName)


def ExportFromSAP(SavePath, OverWrite):
    # Converting Overwrite Input to Lowercase
    OverWrite = str(OverWrite).lower()

    #Exportbtn = Pool.ExportBtn
    #SpreadOption = Pool.SpreadOption
    #ClickOnScreen(Exportbtn, "Left")
    time.sleep(1)
    Driver.press("S")
    Driver.hotkey("enter")
    time.sleep(2)
    i = 0
    while i <= 10:
        SpdOption = At.win_exists("Save As")
        if SpdOption == 1:
            break
        Driver.click()
        Driver.press("S")
        Driver.hotkey("enter")
        time.sleep(10)
        try:
            At.win_active("Save As")
        except Exception:
            print("Trying to Find Save As window")
        i = i+1

    At.win_wait_active("Save As", 10000)
    At.control_click("Save As", "Edit1")
    At.control_set_text("Save As", "Edit1", SavePath)
    At.control_click("Save As", "Button2")

    # Invoking OverWrite the exisiting file if it is marked yes
    if OverWrite == "yes":
        time.sleep(2)
        try:
            ConfirmSaveWin = At.win_exists("Confirm Save As")
            if ConfirmSaveWin == 1:
                At.control_click("Confirm Save As", "Button1")
        except:
            None

    # Allowing Security to Save the file
    try:
        for _ in range(2):
            time.sleep(1)
            At.win_wait_active("SAP GUI Security", 30)
            time.sleep(1)
            At.control_click("SAP GUI Security", "Button2")
    except:
        None


def KillWins(WindowName):
    At.win_kill(WindowName)


def MergerExcel(Excel1, Excel2, Destination):
    df1 = pd.read_excel(Excel1)
    df2 = pd.read_excel(Excel2, header=None)
    merged_df = pd.concat([df1, df2])
    merged_df.to_excel(Destination, index=False)


Log.basicConfig(filename='Process.log', level=Log.INFO,
                format="%(asctime)s %(levelname)s %(message)s")


def TrackBot(Msg):
    Log.info(Msg)


def JustSaveFromSAP(SavePath, OverWrite):
    # Converting Overwrite Input to Lowercase
    OverWrite = str(OverWrite).lower()
    time.sleep(2)
    i = 0
    while i <= 10:
        SpdOption = At.win_exists("Save As")
        if SpdOption == 1:
            break
        Driver.click()
        Driver.press("S")
        Driver.hotkey("enter")
        time.sleep(10)
        try:
            At.win_active("Save As")
        except Exception:
            print("Trying to Find Save As window")
        i = i+1

    At.win_wait_active("Save As", 10000)
    At.control_click("Save As", "Edit1")
    At.control_set_text("Save As", "Edit1", SavePath)
    time.sleep(1)
    At.control_click("Save As", "Button2")
    time.sleep(1)
    At.control_click("Save As", "Button2")

    # Invoking OverWrite the exisiting file if it is marked yes
    if OverWrite == "yes":
        time.sleep(2)
        try:
            ConfirmSaveWin = At.win_exists("Confirm Save As")
            if ConfirmSaveWin == 1:
                At.control_click("Confirm Save As", "Button1")
        except:
            None

    # Allowing Security to Save the file
    try:
        for _ in range(2):
            time.sleep(1)
            At.win_wait_active("SAP GUI Security", 30)
            time.sleep(1)
            At.control_click("SAP GUI Security", "Button2")
    except:
        None


def autoclick(ImageLocation, MouseButton, clicks):
    try:
        res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
        ImageCoordinates = Driver.center(res)

        if ImageCoordinates is not None:
            x, y = ImageCoordinates.x, ImageCoordinates.y
        else:
            print("Image not found")

        At.mouse_click(MouseButton, x, y,clicks,1)

        time.sleep(0.5)
    except Exception as e:
        print("Cannot click on object: " + str(e))
        exit()


def FindElement(ImageLocation, Timeout):
    while True:
        try:
            res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
            if res is not None:
                time.sleep(1)
                return 1
        except:
            time.sleep(Timeout)
            break

def FindElement2(ImageLocation):
            time.sleep(1)
            res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
            if res is not None:
                time.sleep(1)
                return 1
            else:
                 return 0

            


def SelectWin(WinName, Timeout):
    At.win_wait(WinName, 30)
    At.win_activate(WinName)


def autoRelclick(ImageLocation, xaxis, yaxis, MouseButton):
    try:
        res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
        ImageCoordinates = Driver.center(res)

        if ImageCoordinates is not None:
            x, y = ImageCoordinates.x, ImageCoordinates.y
            x = x+xaxis
            y = y+yaxis
        else:
            print("Image not found")

        At.mouse_click(MouseButton, x, y, 1, 1)
        time.sleep(0.5)
    except Exception as e:
        print("Cannot click on object: " + str(e))
        exit()
def autoRelDoubleclick(ImageLocation, xaxis, yaxis, MouseButton):
    try:
        res = Driver.locateOnScreen(ImageLocation, confidence=0.9)
        ImageCoordinates = Driver.center(res)

        if ImageCoordinates is not None:
            x, y = ImageCoordinates.x, ImageCoordinates.y
            x = x+xaxis
            y = y+yaxis
        else:
            print("Image not found")

        At.mouse_click(MouseButton, x, y, 2,1)
        time.sleep(0.5)
    except Exception as e:
        print("Cannot click on object: " + str(e))
        exit()

def GetObject(ElementName):
    try:
        cwd = os.getcwd()
        folder_path = cwd+'\\Assets\\' # Replace with the path to your folder

        if os.path.exists(folder_path) and os.path.isdir(folder_path):
            # List all files in the folder
            files = os.listdir(folder_path)
            
            # Store the filenames in a list
            filenames = []
            for filename in files:
                filenames.append(filename)

            itemcount = 1
            for item in filenames:
                if ElementName in item:
                    AssetName = (folder_path +str(item))
                    
                    return AssetName
                else:
                    if itemcount > len(filenames):
                        print("Element Item not found in assets")


        else:
            print(f"The folder '{folder_path}' does not exist or is not a directory.")
    except Exception as exp:
        TrackBot("There was some error while getting object from assets: " + str(exp))
