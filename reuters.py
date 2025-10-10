from configparser import ConfigParser
import subprocess
import pyautogui
import win32gui
from time import sleep


def reuters_sign_in(path):
    config_object = ConfigParser()
    config_object.read(path)
    
    username = config_object['REUTERS']['Username']
    password = config_object['REUTERS']['Password']
    
    eikon_terminal = subprocess.run(["C:\Example\Thomson Reuters\Eikon\Eikon.exe"],
                                     capture_output=True)
    if eikon_terminal.returncode == 0:
        sleep(7)
        print('Eikon terminal opened succefully.')
    else:
        print('Eikon terminal failed to open.')
        return
    
    screenWidth, screenHeight = pyautogui.size()
    
    eikon_num = win32gui.FindWindow(None, 'REFINITIV EIKON')
    if eikon_num != 0:
        win32gui.SetForegroundWindow(eikon_num)
    else:
        print('Eikon application not found running.')
        return
    
    currentMouseX, currentMouseY = pyautogui.position()
    
    #UserID
    pyautogui.moveTo(843, 481)
    pyautogui.click() 
    pyautogui.write(username)
    
    #Password
    pyautogui.moveTo(836, 525)
    pyautogui.click()
    pyautogui.write(password)
    
    #SignIn
    sleep(0.5)
    pyautogui.moveTo(1053, 573)
    pyautogui.click()
    
    pyautogui.moveTo(currentMouseX, currentMouseY)
