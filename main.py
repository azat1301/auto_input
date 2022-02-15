
import pandas as pd
from pywinauto import Application
from pywinauto import keyboard
import time
import win32gui
import tkinter
from tkinter import filedialog

#Filepath
tkinter.Tk().withdraw() # prevents an empty tkinter window from appearing
file_path = filedialog.askopenfilename()
df = pd.read_excel(file_path)

#Regex_for_specific_filtering
import re
def str_not_contains_UMUL(string):
    x = re.match('(^UM)|(^UL)', string)
    if x :
      return(False)
    return(True)

#Pick data needed and auto input it to system
for index, row in df.iterrows():
    if (((row['maker_nm'] == '切削品') | (row['maker_nm'] == 'その他（製作品等）'))) & (row['prov_typ'] == '3') & str_not_contains_UMUL(df.iloc[index]['itm_p.itm_num'] ):
        part_num = df.iloc[index]['itm_p.itm_num']
        prod_name = df.iloc[index]['itm_p.itm_nm']
        qty = df.iloc[index]['ord_prov_qty']
        sleep_duration = 0.3
        app = Application(backend="win32").connect(title="生産管理 - [生産管理システム]")
        dlg = app.window(title="受注入力")
        app['dlg']['control']
        dlg.set_keyboard_focus()
        keyboard.send_keys(part_num)
        time.sleep(sleep_duration)
        keyboard.send_keys('{RIGHT}')
        while win32gui.GetCursorInfo()[1] == 65541 | win32gui.GetCursorInfo()[1] == 65542:
            pass
        time.sleep(sleep_duration)
        while win32gui.GetCursorInfo()[1] == 65543:
            time.sleep(sleep_duration)
        time.sleep(sleep_duration)
        while win32gui.GetCursorInfo()[1] == 65543:
            time.sleep(sleep_duration)
        time.sleep(sleep_duration)
        while win32gui.GetCursorInfo()[1] == 65543:
            time.sleep(sleep_duration)
        time.sleep(sleep_duration)

        first_loop_flag = True
        for a_name in prod_name.split():
            if first_loop_flag:
                first_loop_flag = False
            else:
                keyboard.send_keys('{SPACE}')
            keyboard.send_keys(a_name)
        time.sleep(sleep_duration)
        time.sleep(sleep_duration)

        keyboard.send_keys('{RIGHT}')
        time.sleep(sleep_duration)
        print(qty)
        keyboard.send_keys(str(qty))
        time.sleep(sleep_duration)
        keyboard.send_keys('{RIGHT}')
        while win32gui.GetCursorInfo()[1] == 65541 | win32gui.GetCursorInfo()[1] == 65542:
            pass
        time.sleep(sleep_duration)
        while win32gui.GetCursorInfo()[1] == 65543:
            time.sleep(sleep_duration)
        time.sleep(sleep_duration)
        while win32gui.GetCursorInfo()[1] == 65543:
            time.sleep(sleep_duration)
        time.sleep(sleep_duration)
        while win32gui.GetCursorInfo()[1] == 65543:
            time.sleep(sleep_duration)
        time.sleep(sleep_duration)
        keyboard.send_keys('{DOWN}')
        time.sleep(sleep_duration)
        keyboard.send_keys('{LEFT}')
        time.sleep(sleep_duration)
        keyboard.send_keys('{LEFT}')
        time.sleep(sleep_duration)
        keyboard.send_keys('{LEFT}')
        time.sleep(sleep_duration)
