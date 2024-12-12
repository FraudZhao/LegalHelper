#第一部分：导入包
from tkinter import *
import os
import platform
import multiprocessing
from docx import Document
import pyperclip

#第二部分：文件路径与文件列表
dir_pyw=os.getcwd()#法典小助手.pyw所在文件夹的路径
dir_folder1= dir_pyw + '/' + '法律' #"法律"文件夹的路径
dirlist_folder2=os.listdir(dir_folder1)#法律种类文件夹路径的列表
dir_result=dir_pyw+'/'+'result.txt'#检索结果文件路径
dir_clipboard=dir_pyw+'/'+'clipboard.txt'#剪贴板路径

#生成各法律文件直接路径列表
dirlist_files=list()
for type in dirlist_folder2:
    if '.DS' in type:
        continue
    else:
        for file in os.listdir(dir_folder1 + '/' + type):
            if '.DS' in file:
                continue
            else:
                dirlist_files.append(dir_folder1 + '/' + type + '/' + file)

#生成分类后各种类法律文件路径列表之列表
def classification(lawlistRV, type):
    statutelist = list()
    interpretationlist = list()
    otherslist = list()
    for law in lawlistRV:
        if '.DS_Store' in law or 'DS' in law or law.startswith('.'):
            continue
        if '法典' in law or law.endswith('法.docx'):
            statutelist.append(type + '/' + law)
        elif '解释' in law or '规定' in law or '批复' in law or '决定' in law:
            interpretationlist.append(type + '/' + law)
        else:
            otherslist.append(type + '/' + law)
    lawlistRV = statutelist + interpretationlist + otherslist
    return lawlistRV

lawlist = list()
for type in dirlist_folder2:
    dir_folder2 = dir_folder1 + '/' + type
    if '.DS_Store' in dir_folder2:
        continue
    if os.path.isdir(dir_folder2) == False:
        print("请注意，“法律”文件夹下不能直接放docx文档。\n请将docx文档放入相应的法律类别文件夹中。\n")
        break
    dirlist_filesofsingletype = os.listdir(dir_folder2)
    dirlist_filesofsingletype.sort()
    lawlist.append(classification(dirlist_filesofsingletype,type))

#第三部分：多进程法典转化
def numtrans(num):
    length = len(num)
    chnnum = ('零', '一', '二', '三', '四', '五', '六', '七', '八', '九')

    def num2(num, control):
        if num.endswith('0'):
            if num.startswith('1'):
                if control:
                    return '十'
                else:
                    return chnnum[int(num[0])] + '十'
            else:
                return chnnum[int(num[0])] + '十'
        else:
            if num.startswith('1'):
                if control:
                    return '十' + chnnum[int(num[-1])]
                else:
                    return chnnum[int(num[0])] + '十' + chnnum[int(num[-1])]
            else:
                return chnnum[int(num[0])] + '十' + chnnum[int(num[-1])]

    def num3(num):
        if num.endswith('00'):
            return chnnum[int(num[0])] + '百'
        elif int(num[-2:]) <= 9:
            return chnnum[int(num[0])] + '百' + '零' + chnnum[int(num[-1])]
        else:
            return chnnum[int(num[0])] + '百' + num2(num[-2:], False)

    def num4(num):
        if num.endswith('000'):
            return chnnum[int(num[0])] + '千'
        elif int(num[-3:]) <= 9:
            return chnnum[int(num[0])] + '千' + '零' + chnnum[int(num[-1])]
        elif int(num[-3:]) <= 99:
            return chnnum[int(num[0])] + '千' + '零' + num2(num[-2:], False)
        else:
            return chnnum[int(num[0])] + '千' + num3(num[-3:])

    if length == 1:
        return chnnum[int(num)]
    elif length == 2:
        return num2(num, True)
    elif length == 3:
        return num3(num)
    else:
        return num4(num)

def convert(dir):
    file = Document(dir)
    content = list()
    art_num = 1
    new_secion_exclude = False
    for paragraph in file.paragraphs:
        paratext = paragraph.text
        if len(paratext) == 0:
            continue
        if '。' not in paratext and '；' not in paratext and '：' not in paratext:
            new_secion_exclude = True
            continue
        if "法宝" in paratext or '*' in paratext:
            break
        if "条之" not in paratext[:12]:
            if (('第' + numtrans(str(art_num)) + '条') in paratext[:12]) or ((str(art_num) + '.') in paratext[:10]):
                new_secion_exclude = False
                art_num += 1
                content.append(paratext)
            elif new_secion_exclude == True:
                continue
            else:
                content[-1] = content[-1] + '\n' + paratext
        else:
            if numtrans(str(art_num - 1)) in paratext[:12]:
                content.append(paratext)
            else:
                content[-1] = content[-1] + '\n' + paratext
    tempdict = {dir: content}
    return tempdict

def start_multiprocessing(dirlist_files):
    with multiprocessing.Pool(processes=os.cpu_count()) as pool:
        rsl = pool.map(convert, dirlist_files)
        pool.close()

        lawdict = dict()
        for item in rsl:
            lawdict.update(item)

        return lawdict

#第四部分：GUI与其他功能
def start_gui(lawlist, lawdict):
    #基础框体
    root = Tk()
    root.title("法典小助手")
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size_geo = '%dx%d+%d+%d' % (1000, 618, (screenwidth - 1000) / 2, (screenheight - 618) / 2 - 50)
    root.geometry(size_geo)
    root["background"] = '#66CCFF'
    #各元素对应的功能
    def simple_output():
        result = open(dir_result, 'r', encoding='utf8')
        text1.insert(0.0, result.read().replace('　', '  '))
        result.close()

    def clear():
        text1.delete(0.0, 'end')

    def addtag(text, colorconfig):
        color = ('#00FF7F', '#00CED1', '#FFC0CB', '#F0E68C', '#A9A9A9')
        end = str(1.0)
        count = 0
        try:
            while True:
                start = text1.search(text, end, stopindex="end")
                end = "%s+%sc" % (start, str(len(text)))
                tagname = text + str(count)
                text1.tag_add(tagname, start, end)
                text1.tag_config(tagname, background=color[colorconfig], foreground="black")
                count += 1
        except:
            pass

    def search():
        input = entry1.get()
        result = open(dir_result, 'w', encoding='utf8')
        if input == '':
            text1.insert(0.0, "您尚未输入检索内容\n")
            return None
        if len(listbox1.curselection()) == 0:
            text1.insert(0.0, "您尚未选中检索范围\n")
            return None

        selectedindextuple = listbox1.curselection()
        selections = listbox1.get(0, listbox1.size() - 1)
        selectedlist = list()
        for index in selectedindextuple:
            selectedlist.append(selections[index])

        if '-' in input:
            rangeofarticle = input.split('-')
            startofrange = int(rangeofarticle[0])
            endofrange = int(rangeofarticle[1])
            for singletypelawslist in lawlist:
                for law in singletypelawslist:
                    if law[1 + law.index('/'):law.index('.')] in selectedlist:
                        articleslist = lawdict[os.getcwd() + '/法律/' + law]
                        result.write('\n' + law + '\n')
                        for item in articleslist[startofrange - 1:endofrange]:
                            result.write(item + '\n')
            else:
                result.close()
                simple_output()
            return None

        if input.isdigit():
            for singletypelawslist in lawlist:
                for law in singletypelawslist:
                    if law[1 + law.index('/'):law.index('.')] in selectedlist:
                        articleslist = lawdict[os.getcwd() + '/法律/' + law]
                        result.write('\n' + law + '\n')
                        result.write(articleslist[int(input) - 1])
            else:
                result.close()
                simple_output()
            return None

        keywordlist = list()
        if ' ' in input:
            keywordlist = input.split(' ')
            if '' in keywordlist:
                text1.insert(0.0, "请检查输入中是否有多余的空格\n")
                return None
        else:
            keywordlist.append(input)
        for index, keyword in enumerate(keywordlist):
            if '.' in keyword:
                dotindex = keyword.index('.')
                keywordlist[index] = '第' + numtrans(keyword[:dotindex]) + '条之' + numtrans(keyword[dotindex + 1:])
            if keyword.isdigit():
                text1.insert(0.0, "输入多关键词时不支持数字转条文\n")
                return None
        for singletypelawslist in lawlist:
            for law in singletypelawslist:
                if law[1 + law.index('/'):law.index('.')] in selectedlist:
                    keyname = os.getcwd() + '/法律/' + law
                    articleslist = lawdict[keyname]
                    result.write('\n' + law + '\n')
                    for article in articleslist:
                        for keyword in keywordlist:
                            if keyword not in article:
                                break
                        else:
                            result.write(article + '\n')
        else:
            result.close()
            simple_output()
            colorcount = 0
            for keyword in keywordlist:
                addtag(keyword, colorcount % 5)
                colorcount += 1


    #各元素本体
    searchbutton = Button(root, text='检索Enter', font=('宋体', 10), command=search, bg='white')
    searchbutton.place(x=40, y=90, width=75, height=30)

    entry1 = Entry(root, bg='white', fg='black')
    entry1.place(x=40, y=40, width=150, height=40)

    text1 = Text(root, wrap='char', font=("黑体", 14), fg='black', bg='white', spacing1=10,spacing2=10, spacing3=10, bd=0)
    text1.place(x=200, y=40, width=740, height=538)

    sbar1 = Scrollbar(command=text1.yview, bg='white')
    sbar1.place(x=940, y=40, width=20, height=538)
    text1.config(yscrollcommand=sbar1.set)

    clearbutton = Button(root, text="清除Delete", font=('宋体', 10), command=clear, bg='white')
    clearbutton.place(x=115, y=90, width=75, height=30)

    label1 = Label(text='检索范围', font=('宋体', 12), bg='white', fg='black')
    label1.place(x=40, y=130, width=144, height=20)

    if platform.system() == 'Darwin':
        listbox1 = Listbox(root, bg='white', fg='black', bd=0, selectmode='multiple', width='16')
    elif platform.system() == 'Windows':
        listbox1 = Listbox(root, bg='white', fg='black', bd=0, selectmode='multiple', width='20')
    listbox1.place(x=40, y=150, height=428)
    listbox1.configure(exportselection=False)
    lb1index = 0
    typeoflaw = 0
    bgcolorlist = ['lightgrey', 'white']
    for dirlist_filesofsingletype in lawlist:
        typeoflaw += 1
        for file in dirlist_filesofsingletype:
            listbox1.insert('end', file[1 + file.index('/'):file.index('.')])
            listbox1.itemconfig(lb1index, bg=bgcolorlist[typeoflaw % 2])
            lb1index += 1
    #键盘事件
    def fullscreen():
        root.geometry('%dx%d' % (screenwidth, screenheight))
        sbar1.place(x=screenwidth - 60, y=40, width=20, height=screenheight - 105)
        text1.place(x=200, y=40, width=screenwidth - 260, height=screenheight - 105)
        listbox1.place(x=40, y=150, width=144, height=screenheight - 215)

    def halfscreen():
        root.geometry('%dx%d' % (screenwidth // 2, screenheight))
        sbar1.place(x=screenwidth // 2 - 60, y=40, width=20, height=screenheight - 95)
        text1.place(x=200, y=40, width=screenwidth // 2 - 260, height=screenheight - 95)
        listbox1.place(x=40, y=150, width=144, height=screenheight - 205)

    def default():
        root.geometry(size_geo)
        sbar1.place(x=940, y=40, width=20, height=538)
        text1.place(x=200, y=40, width=740, height=538)
        listbox1.place(x=40, y=150, width=144, height=428)

    def event_search(event=None):
        search()

    def event_delete(event=None):
        clear()

    def event_up(event=None):
        global pagelocation
        if pagelocation == 0:
            pagelocation += 1
            fullscreen()
        elif pagelocation == -1:
            pagelocation += 1
            default()

    def event_lr(event=None):
        global pagelocation
        halfscreen()
        pagelocation = 0

    def event_down(event=None):
        global pagelocation
        if pagelocation == 1:
            default()
            pagelocation -= 1
        elif pagelocation == 0:
            pagelocation -= 1

    def event_deleteandsearch(event=None):
        clear()
        search()

    def event_copytoclipboard(event=None):
        result = open(dir_result, 'r', encoding='utf8')
        clipboard=open(dir_clipboard,'w',encoding='utf8')
        for line in result.readlines():
            if '.docx' in line:
                continue
            else:
                clipboard.write(line)
        else:
            result.close()
            clipboard.close()
        clipboard=open(dir_clipboard,'r',encoding='utf8')
        pyperclip.copy(clipboard.read())
        clipboard.close()

    root.bind('<Return>', event_search)
    root.bind('<Delete>', event_delete)
    root.bind('<KeyRelease-Up>', event_up)
    root.bind('<KeyRelease-Right>', event_lr)
    root.bind('<KeyRelease-Left>', event_lr)
    root.bind('<KeyRelease-Down>', event_down)
    root.bind('<Shift-Return>', event_deleteandsearch)
    root.bind('<Control-s>', event_copytoclipboard)
    root.bind('<Command-s>', event_copytoclipboard)
    root.bind('<Command-S>', event_copytoclipboard)
    root.bind('<Control-S>', event_copytoclipboard)
    #进入事件循环
    root.mainloop()

#第五部分：核心思路在于分离多进程的法典转化和单进程的其他业务，以防止windows上弹出多个窗口的bug
if __name__ == '__main__':
    lawdict = start_multiprocessing(dirlist_files)
    start_gui(lawlist, lawdict)