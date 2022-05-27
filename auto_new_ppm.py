import numpy as np
import pandas as pd
import pyautogui as pg
import time
from auto_func import click, open_ie, open_authen_system, get_info, get_cha_info, cp_type
#from selenium.webdriver import Ie
#from selenium.webdriver.ie.options import Options


def enter_user_management(info,top_num,bot_num,sp_num,top=True,new=True):
    order, cha_top, cha_bot = pd.read_csv('order.csv',encoding='gbk'), pd.read_csv('cha_top.csv',encoding='gbk'), pd.read_csv('cha_bot.csv',encoding='gbk')
    top_list, bot_list = cha_top.loc[[int(x) for x in top_num]], cha_bot.loc[[int(x) for x in bot_num]]
    dic=dict(zip(order['部室'],order['顺序']))
    sp_list = []
    if '1' in sp_num:
        sp_list.append(1)
    if '2' in sp_num:
        sp_list.append(2)
    if '3' in sp_num:
        sp_list.append(7)
    if '4' in sp_num:
        sp_list.append(3)
    if '5' in sp_num:
        sp_list.append(4)
    if '6' in sp_num:
        sp_list.append(5)
    if '0' in top_num or '0' in bot_num:
        sp_list.append(8)
    if '1' in top_num or '1' in bot_num:
        sp_list.append(9)
    if '10' in top_num or '9' in bot_num:
        sp_list.append(10)
    if '11' in top_num or '10' in bot_num:
        sp_list.append(11)
    
    
    click(92,433,1) #用户管理

    if new:
        click([381,482,1],[1021,568,0.5]) #新增
        pg.typewrite(info['统一认证号'])
        click([1021,613,0.5],[1021,792,0.5],[1021,853,0.5]) #导入、点击、选择部门
        pg.press('down',presses=2,interval=0.2,_pause=0.5)
        pg.press('right',_pause=0.2)
        pg.press('down',presses=dic[info['部室']]+1,interval=0.1,_pause=0.5)
        pg.press('right',_pause=0.2)
        click(1600,900,1) #退出
        pg.press('down',presses=4,interval=0.2,_pause=0.5)
        click(1021,811,1)
        cp_type(info['邮箱'])
        time.sleep(0.5)
        click(1084,976,1)
        #---第一页完成---
        click(900,566,1) #定位
        cp_type(info['部室'])
        click(1393,565,1) #查询
        if top:
            click([938,700,0.3],[1125,965,0.3],[675,675,0.5]) #定位 
            pg.click()
            for i in range(len(top_list)):
                cp_type(top_list['名称'].iloc[i])
                click(947,669,0.5) #搜索
                if top_list['名称'].iloc[i] == '风险管理员':
                    click(518,781,0.3) #选择
                else:
                    click(518,750,0.3) #选择
                click([1080,812,0.3],[654,675,1]) #移动、定位
                pg.hotkey('ctrl','a') 
                pg.press('del')
            cp_type('工作台查询员')
            click([947,669,0.3],[518,750,0.3],[1080,812,0.3]) #搜索、选择、移动
            pg.press('down',presses=4,interval=0.1,_pause=0.5)
            click(1124,978,1)
            #---第三页完成--- 
            click(1178,704,1) #点选
            step = np.array(sp_list)-np.array([0]+sp_list[:-1])
            for j in step:
                pg.press('down',presses=j,interval=0.2,_pause=0.5)
                pg.press('enter',presses=1,_pause=0.5)
            
        else:
            click([938,780,0.3],[1125,965,0.3],[654,675,0.5])
            pg.click()
        #---第二页完成---
            for i in range(len(bot_list)):
                cp_type(bot_list['名称'].iloc[i])
                click([947,669,0.3],[518,750,0.3],[1080,812,0.3],[654,675,1]) #搜索、选择、移动
                pg.hotkey('ctrl','a') 
                pg.press('del')
            cp_type('工作台查询员')
            click([947,669,0.3],[518,750,0.3],[1080,812,0.3]) #搜索、选择、移动
            pg.press('down',presses=4,interval=0.1,_pause=0.5)
            click(1124,978,1)
            #---第三页完成--- 
            click(1178,704,1) #点选
            step = np.array(sp_list)-np.array([0]+sp_list[:-1])
            for j in step:
                pg.press('down',presses=j,interval=0.1,_pause=0.5)
                pg.press('enter',presses=1,_pause=0.5)
            click(1132,869,1)        
        click(1132,869,1)
        next_task = pg.confirm('请手工选择审批角色，完成后请选择下一步操作：','提示',buttons=['继续开底层','开新用户'])               
        return next_task
        #---
    else:
        click(1057,331,0.5) #定位
        pg.typewrite(info['统一认证号'])
        click(1057,422,0.5) #查询
        if top:
            click(1647,589,0.5) #顶层
        else:
            click(1647,626,0.5) #底层
        next_task = pg.confirm('请手工选择用户角色，完成后请选择下一步操作：','提示',buttons=['继续开底层','开新用户'])
        return next_task
        
        
        
def process():
    open_ie()
    open_authen_system()
    click(647,116,1)
    info_lst = pd.read_excel('p.xlsx',dtype=str)
    while True:
        info = get_info(info_lst)
        top_num,bottom_num,sp_num = get_cha_info()
        new_prompt = pg.confirm('该用户的开户形式是？','提示',buttons=['新增开户','修改信息'])
        if_new = True if new_prompt =='新增开户' else False
        if_top = True
        click(75,378,0.5) #系统管理
        while True:
            next_task = enter_user_management(info,top_num,bottom_num,sp_num,top=if_top,new=if_new)
            click(509,259,1)
            if next_task == '继续开底层':
                if_top = False
                continue
            else:
                click(75,378,0.5) #系统管理
                break

process()
