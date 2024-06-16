#需要安装python-docx和pyperclip库
#python docx库可以用清华源：pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-docx
#pyperclip库没有清华源：pip install pyperclip

#路径命名规范 种类_内容和层级
#种类包括单个目录dir和目录列表dirlist
#内容包括文件夹folder和具体文件file
#层级包括0、1、2。内容为具体文件file时，不加层级

##########PART0 库调用、路径和其他前置准备##########

from tkinter import *
from docx import * 
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import pyperclip
import platform

system=platform.system()

dir_folder0=os.getcwd()#法典小助手.py所在文件夹的目录。
dir_folder1= dir_folder0 + '/' + '法律' #"法律"文件夹的目录
dirlist_folder2=os.listdir(dir_folder1)#法律种类文件夹的目录的列表
dir_result=dir_folder0+'/'+'result.txt'
dir_clipboard=dir_folder0 + '/' + 'clipboard.txt'

##########PART1 框体部分##########

#1框体
root=Tk()
root.title("法典小助手")
screenwidth=root.winfo_screenwidth() #本机横向分辨率
screenheight=root.winfo_screenheight() #本机纵向分辨率
size_geo='%dx%d+%d+%d' % (1000,618,(screenwidth-1000)/2, (screenheight-618)/2-50) #默认界面大小1000x618
root.geometry(size_geo)
root["background"]='#66CCFF'#天依蓝

##########PART2 键盘事件部分##########

#2.1几个通用函数
pagelocation=0  #窗口状态,1为全屏，0为原始，-1为最小化
#2.1.1全屏
def fullscreen():
    root.geometry('%dx%d'%(screenwidth,screenheight))
    sbar1.place(x=screenwidth-60,y=40,width=20,height=screenheight-105)
    text1.place(x=200,y=40,width=screenwidth-260,height=screenheight-105)
    lb1.place(x=40,y=150,width=144,height=screenheight-215)
    
#2.1.2半屏
def halfscreen():
    root.geometry('%dx%d'%(screenwidth//2,screenheight))
    sbar1.place(x=screenwidth//2-60,y=40,width=20,height=screenheight-95)
    text1.place(x=200,y=40,width=screenwidth//2-260,height=screenheight-95)
    lb1.place(x=40,y=150,width=144,height=screenheight-205)
    
    
#2.1.3还原
def default():
    root.geometry(size_geo)
    sbar1.place(x=940,y=40,width=20,height=538)
    text1.place(x=200,y=40,width=740,height=538)
    lb1.place(x=40,y=150,width=144,height=428)
    
#2.2键盘事件函数
#2.2.1回车
def event_search(event=None): 
    but1()
#2.2.2delete
def event_delete(event=None): 
    but2()
#2.2.3win+up
def event_up(event=None): 
    global pagelocation
    if pagelocation==0: 
        pagelocation+=1
        fullscreen()
    elif pagelocation==-1: 
        pagelocation+=1
        default()
#2.2.4win+left win+right
def event_lr(event=None): 
    global pagelocation
    halfscreen()
    pagelocation=0
#2.2.5win+down
def event_down(event=None): 
    global pagelocation
    if pagelocation==1: 
        default()
        pagelocation-=1
    elif pagelocation==0:
        pass
        pagelocation-=1
#2.2.6shift+enter
def event_deleteandsearch(event=None):
    but2()
    but1()
#2.2.7control+s或S
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

#2.3键盘事件绑定
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

##########PART 3组件部分##########
        
#3.1输入框
entry1=Entry(root,bg='white',fg='black')
entry1.place(x=40,y=40,width=150,height=40)

#3.2输出框、滚动条、关键词高亮函数

#3.2.1输出框
text1=Text(root,wrap='char',font=("黑体",14),fg='black',bg='white',spacing1=10,\
           spacing2=10,spacing3=10,bd=0)
text1.place(x=200,y=40,width=740,height=538)

#3.2.2滚动条
sbar1=Scrollbar(command=text1.yview,bg='white')
sbar1.place(x=940,y=40,width=20,height=538)
text1.config(yscrollcommand=sbar1.set)

#3.2.3关键词高亮函数
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

#3.3检索标签
label1=Label(text='检索范围',font=('宋体',12),bg='white',fg='black')
label1.place(x=40,y=130,width=144,height=20)

#3.4复选框

def classification(lawlistRV,typeoflaw): #分类函数,lawlistRV为该函数的局部变量，切勿与复选框的lawlist混淆
    statutelist=list()
    interpretationlist=list()
    otherslist=list()
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

lawlist=list() #创建所有法律文件的列表
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

#3.5清除按钮及相关函数
    
#3.5.1输出框清除函数
def but2():
    text1.delete(0.0,'end')
#3.5.2清除按钮本体     
button2=Button(root,text="清除Delete",font=('宋体',10),command=but2,bg='white')
button2.place(x=115,y=90,width=75,height=30)

#3.6检索按钮及相关函数
#3.6检索按钮及相关函数

#3.6.1数字转汉字函数
def numtrans(num): #num为str型或int型
    if type(num)==int:
        num=str(num) #转换为str型，便于调取位数
    length=len(num)
    chnnum=('零','一','二','三','四','五','六','七','八','九')
    
    def num2(num,control):
        if num.endswith('0'): #10,20,30,……,90
            if num.startswith('1'):#10
                if control: #后两位数在该数只有两位时不读一，多于两位时候要读一
                    return '十'
                else:
                    return chnnum[int(num[0])]+'十'
            else: #20,30,……,90
                return chnnum[int(num[0])]+'十'
        else:
            if num.startswith('1'):
                if control:
                    return '十'+chnnum[int(num[-1])]
                else:
                    return chnnum[int(num[0])]+'十'+chnnum[int(num[-1])]
            else:
                return chnnum[int(num[0])]+'十'+chnnum[int(num[-1])]
    
    def num3(num): #处理三位数的函数，可以在四位数中用到
        if num.endswith('00'): #100的倍数
            return chnnum[int(num[0])]+'百'
        elif int(num[-2:])<=9: #后二位小于10
            return chnnum[int(num[0])]+'百'+'零'+chnnum[int(num[-1])]
        else:
            return chnnum[int(num[0])]+'百'+num2(num[-2:],False)

    def num4(num):
        if num.endswith('000'): #100的倍数
            return chnnum[int(num[0])]+'千'
        elif int(num[-3:])<=9: #后三位小于10
            return chnnum[int(num[0])]+'千'+'零'+chnnum[int(num[-1])]
        elif int(num[-3:])<=99: #后三位大于等于10,小于100
            return chnnum[int(num[0])]+'千'+'零'+num2(num[-2:],False)
        else:
            return chnnum[int(num[0])]+'千'+num3(num[-3:])
                                                
    if length==1: #0~9
        return chnnum[int(num)]
    elif length==2: #10~99
        return num2(num,True)
    elif length==3: #100~999
        return num3(num)
    else: #1000~9999
        return num4(num)

#3.6.2法典转化函数

contentdict=dict() #创建一个字典来保存某部法典的转化结果，提高继续在同一部法典中检索的速度，亦有节省电量效果

def articlelist(dir):
    if dir in contentdict: #该法典之前检索过
        return contentdict[dir]
    else: #该法典之前未检索过
        file=Document(dir) #创建python-docx对象
        content=list()
        article=1 #条文号从1开始
        first_article=False
        for paragraph in file.paragraphs: #生成以条文为元素的列表content
            if paragraph.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER: #居中的都是章节名，需要跳过
                continue
            if len(paragraph.text)==0: #空行，需要跳过
                continue
            if '、' in paragraph.text[:4]: #一些司法解释会以一、二、为节标题，需要跳过
                continue
            if "法宝" in paragraph.text: #北大法宝的文件末尾有广告，需要跳过
                break
            try:
                
                if "条之"  not in paragraph.text[:12]: #不属于之一、之二……
                    if numtrans(article) in paragraph.text[:12] and '条'in paragraph.text[:12]: #第一条以及本条首款
                        new_article=True;first_article=True
                        article+=1
                    else:
                        new_article=False
                else: #属于之一、之二……，此时应当保持article不变
                    if numtrans(article-1) in paragraph.text[:12]: #本条首款
                        new_article=True;first_article=True
                    else:
                        new_article=False
                        
                if first_article: #从第一条以后开始操作
                    if new_article:
                        content.append(paragraph.text[paragraph.text.index('第'):])#去掉条文号前面的空格             
                    else:
                        content[-1]=content[-1]+'\n'+paragraph.text
            except:
                continue
        contentdict[dir]=content #将转化结果加入字典
        return contentdict[dir]
    
#3.6.3检索按钮函数

def but1():
    #获取关键词
    keyword=entry1.get()
    keyword_str=keyword

    #创建输出文件
    result = open(dir_result, 'w', encoding='utf8')

    #排除特情
    if keyword=='': #未输入内容
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
    if '-' in keyword:
        rangeofarticle=keyword.split('-')
        startofrange=int(rangeofarticle[0])
        endofrange=int(rangeofarticle[1])
        for singletypelawslist in lawlist:  
            for law in singletypelawslist:
                if law[1 + law.index('/'):law.index('.')] in selectedlist:  # 选中了该法律
                    articleslist = articlelist(dir_folder1 + '/' + law)  # 对选中的法律进行法典转化
                    result.write('\n' + law + '\n') #在文件头部写入法律文件名称
                    for item in articleslist[startofrange-1:endofrange]: #对于不含之一、之二的法律文件而言，直接用列表索引就行
                        result.write(item+'\n')
        else:
            result=open(dir_result, 'r', encoding='utf8')  # 输出流程
            text1.insert(0.0, result.read().replace('　', '  '))  # 输出到文本框，苹果电脑上文件自带的空格会导致bug因此需要替换
            result.close()
        return None #结束but1()函数

    #常规检索
    #生成关键词列表
    if ' ' in keyword: #多关键词，最后生成的keyword为list
        multikeyword=True
        keyword=keyword.split(' ')
        for index in range(len(keyword)):
            if '.' in keyword[index]: #判断之一、之二……
                dotindex=keyword[index].index('.')
                keyword[index]='第'+numtrans(keyword[index][:dotindex])+'条'+'之'+numtrans(keyword[index][dotindex+1:])
            if keyword[index].isdigit(): 
                keyword[index]='第'+numtrans(keyword[index])+'条'

    else: #单关键词，最后生成的keyword为str
        multikeyword=False
        if '.' in keyword:#判断之一、之二……
            dotindex=keyword.index('.')
            keyword='第'+numtrans(keyword[:dotindex])+'条'+'之'+numtrans(keyword[dotindex+1:])
        elif keyword.isdigit(): 
            keyword='第'+numtrans(keyword)+'条'

    #根据关键词在选中的法典中进行检索，
    for singletypelawslist in lawlist: #lawlist是所有法律文件的大列表，singletypelawslist是某中法律文件的小列表
        for law in singletypelawslist:
            if law[1+law.index('/'):law.index('.')] in selectedlist: #选中了该法律
                articleslist=articlelist(dir_folder1 + '/' + law) #对选中的法律进行法典转化
                result.write('\n'+law+'\n') #在文件头部写入法律文件名称
                if multikeyword==True: #多个关键词，keyword为list
                    for article in articleslist:
                        for item in keyword:
                            if item not in article:
                                break
                        else: #跑完for循环而没有发生break，意味着这个article包含全部item
                            result.write(article+'\n')
                else: #单个关键词，keyword为str
                    for article in articleslist:
                        if keyword in article:
                            result.write(article+'\n')

    else:
        result.close()

    #输出流程
    result=open(dir_result, 'r', encoding='utf8')
    text1.insert(0.0,result.read().replace('　','  ')) #输出到文本框，苹果电脑上文件自带的空格会导致bug因此需要替换
    result.close()

    #关键词高亮流程
    if multikeyword:
        count=0 #给不同的内容配不同的颜色
        for item in keyword:
            addtag(item,count%5) #只有5种颜色，防止溢出导致错误
            count+=1
    else:
        addtag(keyword,0)
        
#3.6.4检索按钮本体
button1=Button(root,text='检索Enter',font=('宋体',10),command=but1,bg='white')
button1.place(x=40,y=90,width=75,height=30)

##########PART4 进入事件循环##########
root.mainloop()
