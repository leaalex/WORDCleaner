#-*- coding: utf-8 -*-
import re
import win32com
import win32clipboard
from tkinter import *
#from tkFileDialog import *
import win32com.client
from bs4 import BeautifulSoup, Comment

import HtmlClipboard

##################  Тестирование функций  ##################
def create_ol(text):
    def create_li(m):
        text = m.group(0)
        text = re.sub('(?s)<p>(.*?)</p>', '<li>\g<1></li>', text)
        text = '<ol type="none">\n'+text+'</ol>\n'
        return text

    text = re.sub(r'(?s)((<p>\s*(\d+|\w+)[\.\)].*?</p>\n)+)', create_li, text)
    text = re.sub(r'\n\n\Z', '', text)
    return text
############################################################

##################  Работа с таблицами  ##################
def create_tab(text):
    def create_tr(m):
        def create_td(m):
            text = m.group(1)
            text = re.sub('(?s)<td>(.*?)</td>', '<td>\n\g<1>\n</td>\n', text)
            text = '<tr>\n'+text+'</tr>\n'
            return text

        text = m.group(1)
        text = re.sub('(?s)<tr>(.*?)</tr>', create_td, text)
        text = '<table>\n'+text+'</table>\n'
        return text

    text = re.sub(r'(?s)<table>(.*?)</table>\n', create_tr, text)
    text = re.sub(r'\n\n\Z', '', text)
    return text
##########################################################


def specificReplace(data):
    #------- Страный пробел ----------#
    data = re.sub(r' ', ' ', data)
    data = re.sub(r'&nbsp;', ' ', data)
    data = re.sub(r'^\s*', '', data)
    return data

def delHtmlTegs(data):
    data = re.sub(re.escape('<o:p>'), '', data)
    data = re.sub(re.escape('</o:p>'), '', data)
    data = re.sub(re.escape('<span>'), '', data)
    data = re.sub(re.escape('</span>'), '', data)
    return data

def replaceHtmlTegs(data):
    data = re.sub(re.escape('<i>'), '<em>', data)
    data = re.sub(re.escape('</i>'), '</em>', data)
    data = re.sub(re.escape('<b>'), '<strong>', data)
    data = re.sub(re.escape('</b>'), '</strong>', data)
    return data

def formula(text):
    text = re.sub(re.escape('\['), '\(', text)
    text = re.sub(re.escape('\]'), '\)', text)
    text = re.sub(re.escape('\('), ' \(', text)
    text = re.sub(re.escape('\)'), '\) ', text)
    text = re.sub(re.escape('\) ,'), '\),', text)
    text = re.sub(re.escape('\) .'), '\).', text)
    text = re.sub(re.escape('\) ;'), '\);', text)
    text = re.sub(re.escape('\) )'), '\))', text)
    text = re.sub(re.escape('> \('), '>\(', text)
    text = re.sub(re.escape('> \['), '>\[', text)
    return text

def gethtml(*args):
    if HtmlClipboard.GetHtml() is None:
        buttonClean['text'] = 'Нет ничего'
    else:
        buttonClean['text'] = 'Чистить'
        data = HtmlClipboard.GetHtml()
        data = data.decode()
        data = specificReplace(data)

        data = re.sub(r'(?s)\s+', '  ', data)

        data = re.sub(r'(?s)<(\w*)\s.*?>', '<\g<1>>', data)
        data = re.sub(r'>\s+<', '><', data)


        data = delHtmlTegs(data)
        data = replaceHtmlTegs(data)
        data = formula(data)

        #data = BeautifulSoup(data)
        #data = data.prettify()

        data = re.sub(r'(?s)\s+', ' ', data)
        #data = re.sub(r'(?s)(<(\w+)>\.*?<\2>)', '<\g<1>>\n', data)
        data = re.sub(r'(?s)(<(\w+)>.*?</\2>)', '\g<1>\n', data)
        #data = re.sub(r'(?s)(<(\w+)></\2>)', '', data)
        #data = data.decode('utf-8')
##################  Дополнительные изменения текста  ##################

        data = re.sub(r'<p>\s*</p>\n','', data)
        data = re.sub(r'<\!\[if \!supportLists\]>', '', data)
        data = re.sub(r'<\!\[endif\]>', '', data)
        data = create_ol(data)
        data = create_tab(data)

        print("-----------in clipboard -----------")
        print(data)
        win32clipboard.OpenClipboard()
        win32clipboard.SetClipboardText(data, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()








######################################################
#####                 ОФОРМЛЕНИЕ                 #####
######################################################

root = Tk()
root.title("Чистить")
frame = Frame(root)
frametop = Frame(frame)

frame.pack()
frametop.pack(side='top')

#buttonFeature = Button(frametop, text="q", padx='35', pady='5')
#buttonFeature.pack(side='left')

#buttonFeature = Button(frametop, text="q", padx='35', pady='5')
#buttonFeature.pack(side='left')

buttonClean = Button(frame, text='Чистить', width=10, padx='90', pady='90', font='Colibri 14')
buttonClean.pack()
buttonClean.bind("<Button-1>", gethtml)

root.resizable(0, 0)
root.mainloop()