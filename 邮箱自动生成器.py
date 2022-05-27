from xpinyin import Pinyin

name = input('请输入姓名：')
p=Pinyin()
pin = p.get_pinyin(name,',').split(',')
if len(pin) == 3:
    pin = [pin[0],pin[1]+pin[2]]
elif len(pin) == 4:
    pin = [pin[0],pin[1]+pin[2]+pin[3]]
mail = pin[1]+'.'+pin[0]+'@wm.icbc.com.cn'
print('邮箱为：'+mail)
