import pandas as pd
import pyautogui as pg
import pyperclip
import time
from xpinyin import Pinyin

def click(*para):
    """
    类型一：click(x,y,t)
    类型二：click([x1,y1,t1],[x2,y2,t2],...)
    其中，x代表横坐标，y代表纵坐标，t代表等待时间
    """
    if type(para[0]) == list:
        for li in para:
            click(li[0],li[1],li[2])
    if type(para[0]) == int and len(para) == 3:
        pg.click(para[0],para[1])
        time.sleep(para[2])
    
def cp_type(word):
    pyperclip.copy(word)
    pg.hotkey('ctrl','v')
    
def open_ie():
     click(1918,1049,0.5)
     pg.doubleClick(44,301)
     while not pg.pixelMatchesColor(1235,601,(253, 172, 29),tolerance=1):
         time.sleep(0.5)
     
def open_authen_system():
    click(1295,504,0.5)
    pg.typewrite("001333323")
    click(1295,550,0.5)
    pg.typewrite("Gylc=jzr")
    click(1295,594,2)#进入认证平台

def enter_and_split(prom):
    while True:
        info = pg.prompt(prom,'提示')
        if not info:
            return []
        try:
            n = int(info)
            splited_info = [info.strip()]
            break
        except:
            if ' ' in info:
                splited_info = info.split(' ')
            else:
                continue
            break
    return splited_info

def get_info(lst):
    num = enter_and_split('请输入要开户人员的统一认证号：')[0]
    filt = lst[lst['统一认证号'] == num]
    if len(filt) == 1:
        info = filt[['姓名','部室','统一认证号','邮箱']].iloc[0]
    else:
        splited_info = enter_and_split('自动导入失败！请按顺序输入开户人信息，用空格隔开：（例：张三 金融科技部 san.zhang@icbc.com）')
        splited_info.insert(2,num)
        info = pd.Series(splited_info,index=['姓名','部室','统一认证号','邮箱'])
    return info

def get_cha_info():
    cha_top = enter_and_split('请输入要开通顶层角色编号，用空格隔开，不开顶层请留空')
    cha_bot = enter_and_split('请输入要开通底层角色编号，用空格隔开，不开底层请留空')
    cha_sp = enter_and_split('请输入要开通审批角色编号，用空格隔开，不开审批角色请留空')
    return cha_top, cha_bot, cha_sp

def get_mail(cnname):
    pp=Pinyin()
    pin = pp.get_pinyin(cnname,',').split(',')
    if len(pin) == 3:
        pin = [pin[0],pin[1]+pin[2]]
    elif len(pin) == 4:
        pin = [pin[0],pin[1]+pin[2]+pin[3]]
    mail = pin[1]+'.'+pin[0]+'@wm.icbc.com.cn'
    return mail

def add_info(lst):
    change = False
    num = pg.prompt('请输入新增人员统一认证号：','提示')
    filt = lst[lst['统一认证号'] == num]
    if len(filt) == 1:
        info = filt[['姓名','部室','统一认证号','邮箱']].iloc[0]
        conf = pg.confirm(info['姓名']+'，'+info['部室']+'，'+info['邮箱']+'，以上信息是否需要修改？','提示',buttons=['是','否'])
        if conf == '是':
            subs = enter_and_split('请输入要修改的信息，用空格隔开。（1：姓名，2：部室，3：邮箱，例：“2 金融科技部”）')
            if subs[0] == '1':
                info['姓名'] = subs[1]
            if subs[0] == '2':
                info['部室'] = subs[1]
            if subs[0] == '3':
                info['邮箱'] = subs[1]
            change = True
    else:
        name_dep = enter_and_split('未收录！请输入姓名及部室，用空格隔开')
        m = get_mail(name_dep[0])
        f = pg.confirm('该用户邮箱是否为'+m+'？','提示',buttons=['是','否'])
        if f == '否':
            m = pg.prompt('请输入邮箱：','提示')
        info = pd.Series({'姓名':name_dep[0], '部室':name_dep[1], '统一认证号':num, '邮箱':m})
    return info



