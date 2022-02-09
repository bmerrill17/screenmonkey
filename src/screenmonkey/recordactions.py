
import pandas as pd
import numpy as np
import datetime as dt
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
import copy
import os
import matplotlib.pyplot as plt
import xlsxwriter
import pyautogui
from pynput import mouse

def checkTime(prev):
    timeDif = dt.datetime.now() - prev
    return round(timeDif.total_seconds(), 3)

def on_click(x, y, button, pressed):
    global mosPos
    global prev
    action = 'Pressed' if pressed else 'Released'
    buttonStr = 'Left' if str(button) == 'Button.left' else 'Right'
    print('{0} at {1} on {2}'.format(action,(x, y), buttonStr))
    currentDf = {'X' : x, 'Y' : y, 'Action' : action, 'Button' : buttonStr, 'Seconds' : checkTime(prev), 'Type' : 'Mouse'}
    prev = dt.datetime.now()
    mosPos = mosPos.append(currentDf, ignore_index=True)
    print(mosPos)
    if (x <= 0) and (y <= 0):
        # Stop listener
        mosPos.to_excel('screenCoordinates.xlsx')
        return False

def on_scroll(x, y, dx, dy):
    print('Scrolled {0} at {1}'.format(
        'down' if dy < 0 else 'up',
        (x, y)))


# Collect events until released
prev = dt.datetime.now()
mosPos = pd.DataFrame(columns=['X', 'Y', 'Action', 'Button', 'Seconds', 'Type'])
with mouse.Listener(
        on_click=on_click,
        on_scroll=on_scroll) as listener:
    listener.join()