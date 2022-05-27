import pandas as pd
import pyautogui as pg
from auto_func import click, open_ie, open_authen_system, get_info, cp_type


def enter_user_management(info,top=True,new=True):
    if not top:
        click([1590,155,1],[1590,174,2]) #切换底层       
    click([1827,157,1],[1801,173,2]) #选择系统管理员
    if top:
        click([57,438,1],[68,524,1],[444,393,1],[1048,609,1]) #顶层用户管理
    else:
        click([57,358,1],[57,400,1],[444,393,1],[1048,432,1]) #底层用户管理  
    if new:
        if top:
            click(917,749,1) #新增
        else:
            click(961,749,1)
        click(863,498,1) #定位
        pg.typewrite(info['统一认证号'])
        click([1374,493,1],[882,415,0.5]) #查询部门
        cp_type(info['部室'])
        pg.press('enter',presses=2,interval=0.3,_pause=0.5)
        if top:
            click(860,534,0) #邮箱
            cp_type(info['邮箱'])
        next_task = pg.confirm('请手工选择用户角色，完成后请选择下一步操作：','提示',buttons=['继续开底层','开新用户'])
        return next_task
    
    if not new:
        if top:
            click(1094,749,1) #修改
        else:
            click(1142,749,1)
        click(841,406,1) #定位
        pg.typewrite(info['统一认证号'])
        click([1025,550,1],[311,407,1]) #查询选择
        if top:
            click(850,491,1) #修改
        else:
            click(898,490,1)
        next_task = pg.confirm('请手工选择用户角色，完成后请选择下一步操作：','提示',buttons=['继续开底层','开新用户'])
        return next_task
        
def process_ppm():
    open_ie()
    open_authen_system()
    click([124,482,1],[547,431,1]) #进入ppm
    info_lst = pd.read_excel('p.xlsx',dtype=str)
    while True:
        info = get_info(info_lst)
        new_prompt = pg.confirm('该用户的开户形式是？','提示',buttons=['新增开户','修改信息'])
        if_new = True if new_prompt =='新增开户' else False
        if_top = True
        bot_already = False
        while True:
            next_task = enter_user_management(info,top=if_top, new=if_new)
            if next_task == '继续开底层':
                if_top = False
                bot_already = True
                continue
            else:
                if not bot_already:
                    click([1857,155,1],[1857,136,0])
                else:
                    click([1590,155,1],[1590,136,0]) #切换底层  
                break

process_ppm()




