
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
import time
from pynput.keyboard import Key, Controller as KeyboardController
from pynput.mouse import Button, Controller as MouseController

keyboard = KeyboardController()
mouse = MouseController()

df = pd.read_excel('automation/updateForecasts.xlsx')
len = df.shape[0]
for i in range(0, len):
    time.sleep(df['Seconds'][i])
    if df['Type'][i] == 'Mouse':
        mouse.position = (df['X'][i], df['Y'][i])
        if df['Action'][i] == 'Pressed':
            if df['Button'][i] == 'Left':
                mouse.press(Button.left)
            elif df['Button'][i] == 'Right':
                mouse.press(Button.right)
        elif df['Action'][i] == 'Released':
            if df['Button'][i] == 'Left':
                mouse.release(Button.left)
            elif df['Button'][i] == 'Right':
                mouse.release(Button.right)
        print('The current pointer position is {0}'.format(
            mouse.position))
    elif df['Type'][i] == 'Keyboard':
        with keyboard.pressed(Key.shift):
            with keyboard.pressed(Key.ctrl):
                if df['Action'][i] == 'Pressed':
                    if df['Button'][i] == 'Right':
                        keyboard.press(Key.right)
                    elif df['Button'][i] == 'Down':
                        keyboard.press(Key.down)
                    print('Pressed button: {0}'.format(df['Button'][i]))
                elif df['Action'][i] == 'Released':
                    if df['Button'][i] == 'Right':
                        keyboard.release(Key.right)
                    elif df['Button'][i] == 'Down':
                        keyboard.release(Key.down)
                    print('Released button: {0}'.format(df['Button'][i]))