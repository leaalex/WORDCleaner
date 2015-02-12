#-*- coding: utf-8 -*-
import re
import win32com
import win32clipboard
from tkinter import *
#from tkFileDialog import *
import win32com.client
from bs4 import BeautifulSoup, Comment

import HtmlClipboard
import html


# функция перевода &#XXXX; в символы
#start
def replace_unicode(text):
    print("working...")
    def fixup(m):
        text = m.group(0)
        if text[:1] == "<":
            return ""
        if text[:2] == "&#":
            try:
                if text[:3] == "&#x":
                    return chr(int(text[3:-1], 16))
                else:
                    return chr(int(text[2:-1]))
            except ValueError:
                pass
        elif text[:1] == "&":
            import htmlentitydefs
            entity = htmlentitydefs.entitydefs.get(text[1:-1])
            if entity:
                if entity[:2] == "&#":
                    try:
                        return chr(int(entity[2:-1]))
                    except ValueError:
                        pass
                else:
                    return entity.encode("iso-8859-1")
        return text
    text = re.sub("(?s)<*>|&#?\w+;", fixup, text)
    return text

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

##################  Тестирование функций  ##################
def create_il(text):
    def create_li(m):
        text = m.group(0)
        text = re.sub('(?s)<p>·(.*?)</p>', '<li>\g<1></li>', text)
        text = '<il>\n'+text+'</il>\n'
        return text

    text = re.sub(r'(?s)((<p>\s*·.*?</p>\n)+)', create_li, text)
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
    data = re.sub(re.escape('<o:p>'), ' ', data)
    data = re.sub(re.escape('</o:p>'), ' ', data)
    data = re.sub(re.escape('<span>'), ' ', data)
    data = re.sub(re.escape('</span>'), ' ', data)
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

def formula_v2(text):
    text = re.sub(re.escape('\['), '\(', text)
    text = re.sub(re.escape('\]'), '\)', text)
    text = re.sub(re.escape('\('), ' \(', text)
    text = re.sub(r'\\\)\s*', '\) ', text)
    text = re.sub(r'(\\[)(])\s+([.,;)])', '\g<1>\g<2>', text)
    text = re.sub(r'([.,;])\s*(\\\))', '\g<2>\g<1>', text)
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
        data = html.unescape(data)
        data = replace_unicode(data)
        print(data)

        data = formula_v2(data)


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
        #data = create_ol(data)
        #data = create_il(data) Тестирование замены ненумерованого списка
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