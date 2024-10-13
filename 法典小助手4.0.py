#需要安装python-docx和pyperclip库
#python docx库可以用清华源：pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-docx
#pyperclip库没有清华源：pip install pyperclip

#路径命名规范 种类_内容和层级
#种类包括单个目录dir和目录列表dirlist
#内容包括文件夹folder和具体文件file
#层级包括0、1、2。内容为具体文件file时，不加层级

##########库调用、路径和其他前置准备##########

#向开发python及其标准库, python-docx, pyperclip库的前人致敬。
from tkinter import *
import os
import platform
import multiprocessing
from docx import *
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pyperclip

system=platform.system()

dir_folder0=os.getcwd()#法典小助手.py所在文件夹的目录。
dir_folder1= dir_folder0 + '/' + '法律' #"法律"文件夹的目录
dirlist_folder2=os.listdir(dir_folder1)#法律种类文件夹的目录的列表
dir_result=dir_folder0+'/'+'result.txt'
dir_clipboard=dir_folder0 + '/' + 'clipboard.txt'

##########数字转汉字函数##########
def numtrans(num):  # num为str型或int型

    length = len(num)
    chnnum = ('零', '一', '二', '三', '四', '五', '六', '七', '八', '九')

    def num2(num, control):
        if num.endswith('0'):  # 10,20,30,……,90
            if num.startswith('1'):  # 10
                if control:  # 后两位数在该数只有两位时不读一，多于两位时候要读一
                    return '十'
                else:
                    return chnnum[int(num[0])] + '十'
            else:  # 20,30,……,90
                return chnnum[int(num[0])] + '十'
        else:
            if num.startswith('1'):
                if control:
                    return '十' + chnnum[int(num[-1])]
                else:
                    return chnnum[int(num[0])] + '十' + chnnum[int(num[-1])]
            else:
                return chnnum[int(num[0])] + '十' + chnnum[int(num[-1])]

    def num3(num):  # 处理三位数的函数，可以在四位数中用到
        if num.endswith('00'):  # 100的倍数
            return chnnum[int(num[0])] + '百'
        elif int(num[-2:]) <= 9:  # 后二位小于10
            return chnnum[int(num[0])] + '百' + '零' + chnnum[int(num[-1])]
        else:
            return chnnum[int(num[0])] + '百' + num2(num[-2:], False)

    def num4(num):
        if num.endswith('000'):  # 100的倍数
            return chnnum[int(num[0])] + '千'
        elif int(num[-3:]) <= 9:  # 后三位小于10
            return chnnum[int(num[0])] + '千' + '零' + chnnum[int(num[-1])]
        elif int(num[-3:]) <= 99:  # 后三位大于等于10,小于100
            return chnnum[int(num[0])] + '千' + '零' + num2(num[-2:], False)
        else:
            return chnnum[int(num[0])] + '千' + num3(num[-3:])

    if length == 1:  # 0~9
        return chnnum[int(num)]
    elif length == 2:  # 10~99
        return num2(num, True)
    elif length == 3:  # 100~999
        return num3(num)
    else:  # 1000~9999
        return num4(num)

##########多进程法典转化##########

#convert函数
def convert(dir):
    file=Document(dir) #创建python-docx对象
    content=list()
    art_num=1 #条文号从1开始
    new_secion_exclude=False #很多文件每进入一个新节，先叙述宗旨，然后才是条文，需要这些跳过宗旨。具体而言，略过以条文号开头的paragraph前的所有Paragraph。
    for paragraph in file.paragraphs: #生成以条文为元素的列表content
        #通用排除
        paratext=paragraph.text
        if len(paratext)==0: #空行
            continue
        #有效内容前排除
        #判断新节只需一步就行：这一行要是没有句号、冒号、分号（注意不含括号、顿号），那它一定是章节标题。因为法条内容虽不一定有逗号，但一定以句号或冒号或分号之一结尾
        if '。' not in paratext and '；' not in paratext and '：'not in  paratext:
            new_secion_exclude=True
            continue
        #有效内容后排除
        if "法宝" in paratext or '*' in paratext: #北大法宝的文件末尾有广告和规范文件撰写说明，需要终止转化
            break
        #定性
        if "条之" not in paratext[:12]: #不属于之一、之二……
            #本条第1款
            if (('第' + numtrans(str(art_num)) + '条') in paratext[:12]) \
                or ((str(art_num)+'.') in paratext[:10]): #汉字数字型或阿拉伯数字型
                new_secion_exclude=False
                art_num+=1
                content.append(paratext)
            #在新节排除时，遇到的不是由条文号开头的paragraph必须全部排除
            elif new_secion_exclude==True:
                continue
            #本条非第1款
            else:
                content[-1]=content[-1]+'\n'+paratext
        else: #属于之一、之二……，此时应当保持article不变。
            #本条第1款
            if numtrans(str(art_num-1)) in paratext[:12]:#以阿拉伯数字标号的规范性文件不可能有之一、之二，首款判断无需有阿拉伯数字型
                content.append(paratext)
            #本条非第1款
            else:
                content[-1]=content[-1]+'\n'+paratext
    tempdict={dir:content}
    return tempdict

#多进程法典转化
if __name__ == '__main__':
    if system=='Darwin':#mac上多进程管理基于fork
        multiprocessing.set_start_method('fork')
    with multiprocessing.Pool(processes=os.cpu_count()) as pool: #进程池中的进程数量应等于本机cpu核心数
        #1,生成目录列表dirlist
        dirlist = list()
        for type in os.listdir(os.getcwd() + '/法律'):
            if '.DS' in type:
                continue
            else:
                for file in os.listdir(os.getcwd() + '/法律' + '/' + type):
                    if '.DS' in file:
                        continue
                    else:
                        dirlist.append(os.getcwd()+'/法律'+'/'+type+'/'+file)
        #2,把目录列表作为参数，提交给许多个convert。千万注意这一行代码不要放到前面的for里面，否则逻辑上只能在单进程里跑！
        lawdict = dict()
        rsl=pool.map(convert,dirlist) #rsl是一个列表，各元素为小字典
        pool.close()
        for item in rsl:#整合成result大字典，键为法典路径，值为转化好的内容
            lawdict.update(item)

##########框体部分##########
#框体
root=Tk()
root.title("法典小助手")
screenwidth=root.winfo_screenwidth() #本机横向分辨率
screenheight=root.winfo_screenheight() #本机纵向分辨率
size_geo='%dx%d+%d+%d' % (1000,618,(screenwidth-1000)/2, (screenheight-618)/2-50) #默认界面大小1000x618
root.geometry(size_geo)
root["background"]='#66CCFF'#天依蓝

##########键盘事件部分##########

#几个通用函数
pagelocation=0  #窗口状态,1为全屏，0为原始，-1为最小化
#全屏
def fullscreen():
    root.geometry('%dx%d'%(screenwidth,screenheight))
    sbar1.place(x=screenwidth-60,y=40,width=20,height=screenheight-105)
    text1.place(x=200,y=40,width=screenwidth-260,height=screenheight-105)
    lb1.place(x=40,y=150,width=144,height=screenheight-215)
    
#半屏
def halfscreen():
    root.geometry('%dx%d'%(screenwidth//2,screenheight))
    sbar1.place(x=screenwidth//2-60,y=40,width=20,height=screenheight-95)
    text1.place(x=200,y=40,width=screenwidth//2-260,height=screenheight-95)
    lb1.place(x=40,y=150,width=144,height=screenheight-205)

#还原
def default():
    root.geometry(size_geo)
    sbar1.place(x=940,y=40,width=20,height=538)
    text1.place(x=200,y=40,width=740,height=538)
    lb1.place(x=40,y=150,width=144,height=428)
    
#键盘事件函数
#回车
def event_search(event=None): 
    search()
#delete
def event_delete(event=None): 
    clear()
#win+up
def event_up(event=None): 
    global pagelocation
    if pagelocation==0: 
        pagelocation+=1
        fullscreen()
    elif pagelocation==-1: 
        pagelocation+=1
        default()
#win+left win+right
def event_lr(event=None): 
    global pagelocation
    halfscreen()
    pagelocation=0
#win+down
def event_down(event=None): 
    global pagelocation
    if pagelocation==1: 
        default()
        pagelocation-=1
    elif pagelocation==0:
        pass
        pagelocation-=1
#shift+enter
def event_deleteandsearch(event=None):
    clear()
    search()
#control+s或S
def event_copytoclipboard(event=None):
    result=open(dir_result, 'r', encoding='utf8')
    clipboard=open(dir_clipboard, 'w+', encoding='utf8')
    linelist=result.readlines()
    result.close()
    for line in linelist: #给每款添加数字（款）并删除其中的两个空格，项不添加
        if 'docx' in line or len(line)<5:continue #文件名行、调整观感的空行
        if line.startswith('第')and '条' in line[:12]:
            newarticle=True
        else:
            newarticle=False
        if newarticle==True: #新的一条首款
            paragraph=1#款号为1
            #国家法律数据库的文件，首款顶格，“第n条”和主文之间有一个空格。其他行则以两个空格开头。
            clipboard.write(line.replace('　','（%d）'%(paragraph)))
            paragraph+=1
        else: #不是首款
            if '（'in line[:5] and '）' in line[:5]: #前六位有括号，说明是项
                if line.startswith('　　'):#有的文件开头空两格
                    clipboard.write(line.replace('　　','\t'))#把开头两个空格换成制表符，与Onenote缩进搭配
                else:#有的不空格
                    clipboard.write('\t'+line)
                continue
            else: #不是项，则是款
                if line.startswith('　　'):#有的文件开头空两格
                    clipboard.write(line.replace('　　','（%d）'%(paragraph)))
                    paragraph+=1
                else:#有的不空格
                    clipboard.write('（%d）'%(paragraph)+line)
                    paragraph+=1
    else:
        clipboard.close()
    clipboard=open(dir_clipboard, 'r', encoding='utf8')
    pyperclip.copy(clipboard.read())
    clipboard.close()

#键盘事件绑定
root.bind('<Return>',event_search) #将键盘事件绑定到root以实现全局操控
root.bind('<Delete>',event_delete)
root.bind('<KeyRelease-Up>',event_up)
root.bind('<KeyRelease-Right>',event_lr)
root.bind('<KeyRelease-Left>',event_lr)
root.bind('<KeyRelease-Down>',event_down)
root.bind('<Shift-Return>',event_deleteandsearch)
root.bind('<Control-s>',event_copytoclipboard)#大写s和小写s不一样
root.bind('<Command-s>',event_copytoclipboard)
root.bind('<Command-S>',event_copytoclipboard)#大写s和小写s不一样
root.bind('<Control-S>',event_copytoclipboard)

##########组件部分##########
        
#输入框
entry1=Entry(root,bg='white',fg='black')
entry1.place(x=40,y=40,width=150,height=40)

#输出框、滚动条
#输出框
text1=Text(root,wrap='char',font=("黑体",14),fg='black',bg='white',spacing1=10,\
           spacing2=10,spacing3=10,bd=0)
text1.place(x=200,y=40,width=740,height=538)
#滚动条
sbar1=Scrollbar(command=text1.yview,bg='white')
sbar1.place(x=940,y=40,width=20,height=538)
text1.config(yscrollcommand=sbar1.set)

#清除按钮及相关函数
#输出框清除函数
def clear():
    text1.delete(0.0, 'end')
#清除按钮本体
clearbutton = Button(root, text="清除Delete", font=('宋体', 10), command=clear, bg='white')
clearbutton.place(x=115, y=90, width=75, height=30)

#检索标签
label1=Label(text='检索范围',font=('宋体',12),bg='white',fg='black')
label1.place(x=40,y=130,width=144,height=20)

#复选框
def classification(lawlistRV,typeoflaw): #文件分类函数,lawlistRV为该函数的局部变量，切勿与复选框的lawlist混淆
    statutelist=list()#法律
    interpretationlist=list()#解释
    otherslist=list()#其他
    for law in lawlistRV:
        if '.DS_Store' in law or 'DS' in law or law.startswith('.'): #苹果会对访达打开过的文件创建.DS_Store，需要略过
            continue
        if '法典' in law or law.endswith('法.docx'): #法律优先处理，其他后置
            statutelist.append(typeoflaw+'/'+law)
            continue
        elif '解释' in law or '规定' in law or '批复' in law or '决定' in law:
            interpretationlist.append(typeoflaw+'/'+law)
            continue
        else:
            otherslist.append(typeoflaw+'/'+law)
    else:
        del lawlistRV
        lawlistRV=statutelist+interpretationlist+otherslist
    return lawlistRV

lawlist=list() #创建所有法律文件的列表,其元素亦为列表
for folder in dirlist_folder2: #不同类型法律放在不同文件夹中 此处的这个临时变量folder尽量不要重构，很麻烦
    dir_folder2= dir_folder1 + '/' + folder #单个法律文件夹地址列表
    if '.DS_Store' in dir_folder2: #苹果会在访达打开过的文件夹创建.DS_Store文件，需要略过
        continue
    if os.path.isdir(dir_folder2)==False: #法律文件夹下不能直接放文件
        text1.insert(0.0,"请注意，“法律”文件夹下不能直接放docx文档。\n请将docx文档放入相应的法律类别文件夹中。") 
        break
    dirlist_filesofsingletype=os.listdir(dir_folder2)
    dirlist_filesofsingletype.sort()#先初排一遍，这样解释1、解释2、解释3就不会乱
    lawlist.append(classification(dirlist_filesofsingletype, folder))

if system=='Darwin':#创建listbox，mac和windows的宽度不一样。Darwin即macos
    lb1=Listbox(root,bg='white',fg='black',bd=0,selectmode='multiple',width='16')
if system=='Windows':
    lb1=Listbox(root,bg='white',fg='black',bd=0,selectmode='multiple',width='20')

lb1.place(x=40,y=150,height=428)
lb1.configure(exportselection=False)

lb1index=0
typeoflaw=0
bgcolorlist=['lightgrey','white']

for dirlist_filesofsingletype in lawlist:
    typeoflaw+=1
    for file in dirlist_filesofsingletype:
        lb1.insert('end',file[1+file.index('/'):file.index('.')])
        lb1.itemconfig(lb1index,bg=bgcolorlist[typeoflaw%2])#为同种法律设置同种背景色
        lb1index+=1

#检索输出及相关函数
#输出函数
def simple_output():
    result=open(dir_result, 'r', encoding='utf8')
    text1.insert(0.0, result.read().replace('　', '  '))  # 输出到文本框，苹果电脑上文件自带的空格会导致bug因此需要替换
    result.close()
#关键词高亮函数
def addtag(text,colorconfig):
    color=('#00FF7F','#00CED1','#FFC0CB','#F0E68C','#A9A9A9')
    end=str(1.0)
    count=0
    try:
        count=0
        while True:
            start=text1.search(text,end,stopindex="end")
            end="%s+%sc"%(start,str(len(text)))
            tagname=text+str(count)
            text1.tag_add(tagname,start,end)
            text1.tag_config(tagname, background=color[colorconfig], foreground="black")
    except:
        pass
#检索函数
def search():
    #获取关键词
    input=entry1.get()
    #创建输出文件，统一打开，模式为写入，各流程自行关闭
    result = open(dir_result, 'w', encoding='utf8')
    #排除特情
    if input=='': #未输入内容
        text1.insert(0.0,"您尚未输入检索内容\n") 
        return None #结束，防止未输入内容时卡bug
    if len(lb1.curselection())==0:#未选中任何法律
        text1.insert(0.0,"您尚未选中检索范围\n")
        return None #结束，防止未选择时卡bug

    #组建选中法律的列表selectedlist
    selectedindextuple=lb1.curselection()
    selections=lb1.get(0,lb1.size()-1)
    selectedlist=list()
    for index in selectedindextuple:
        selectedlist.append(selections[index])

    #输出从某条到某条检索
    if '-' in input:
        rangeofarticle=input.split('-')
        startofrange=int(rangeofarticle[0])
        endofrange=int(rangeofarticle[1])
        for singletypelawslist in lawlist:  
            for law in singletypelawslist:
                if law[1 + law.index('/'):law.index('.')] in selectedlist:  # 选中了该法律
                    articleslist = lawdict[os.getcwd() + '/法律/' + law]# 因为lawdict的键是完整路径，所以要凑一下
                    result.write('\n' + law + '\n') #在文件头部写入法律文件名称
                    for item in articleslist[startofrange-1:endofrange]: #对于不含之一、之二的法律文件而言，直接用列表索引就行
                        result.write(item+'\n')
            else:
                result.close()
        else:
            simple_output()
        return None #结束but1()函数

    #输出某一条
    if input.isdigit():
        for singletypelawslist in lawlist:
            for law in singletypelawslist:
                if law[1 + law.index('/'):law.index('.')] in selectedlist:
                    articleslist = lawdict[os.getcwd() + '/法律/' + law]
                    result.write('\n' + law + '\n') #在文件头部写入法律文件名称
                    result.write(articleslist[int(input)-1])
        else:
            result.close()
        simple_output()
        return None

    #常规检索
    #1,形成关键词列表。无论是否多关键词，一律使用关键词列表
    keywordlist=list()
    #多关键词和单关键词
    if ' ' in input: #多关键词，最后生成的keyword为list
        keywordlist=input.split(' ')
        if '' in keywordlist: #多余的空格、错误位置的空格会导致split结果中有空字符，出现bug
            text1.insert(0.0, "请检查输入中是否有多余的空格\n")
            return None
    else: #单关键词，最后生成的keyword为str
        keywordlist.append(input)
    #用.表示之一、之二
    for keyword in keywordlist:
        if '.' in keyword:  # 判断之一、之二……
            dotindex = keywordlist[index].index('.')
            keyword = '第' + numtrans(keyword[:dotindex]) + '条之' + numtrans(keyword[dotindex + 1:])
        if keyword.isdigit():
            text1.insert(0.0, "输入多关键词时不支持数字转条文\n")
            return None
    #2,进行匹配
    for singletypelawslist in lawlist: #lawlist是所有法律文件的大列表，singletypelawslist是某中法律文件的小列表
        for law in singletypelawslist:
            if law[1+law.index('/'):law.index('.')] in selectedlist: #选中了该法律
                #从lawdict中读取内容
                keyname=os.getcwd()+'/法律/'+law #因为lawdict的键是完整路径，所以要凑一下
                articleslist=lawdict[keyname]
                #在文件头部写入法律名
                result.write('\n'+law+'\n')
                #进行关键词匹配
                for article in articleslist:
                    for keyword in keywordlist:
                        if keyword not in article:
                            break #未能全部匹配，跳过
                    else: #全部匹配，方能写入
                        result.write(article+'\n')
    else:
        result.close()

    #输出流程
    simple_output()

    #关键词高亮流程
    colorcount=0  # 给不同的内容配不同的颜色
    for keyword in keywordlist:
        addtag(keyword,colorcount%5)
        colorcount+=1

#检索按钮本体
searchbutton=Button(root, text='检索Enter', font=('宋体', 10), command=search, bg='white')
searchbutton.place(x=40, y=90, width=75, height=30)

##########进入事件循环##########
root.mainloop()