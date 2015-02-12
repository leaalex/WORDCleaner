#-*- coding: utf-8 -*-
import re
import win32com
import win32clipboard
from tkinter import *
#from tkFileDialog import *
import win32com.client
from bs4 import BeautifulSoup, Comment

import HtmlClipboard

arrayTemp=None
arrayResult=None
arrayNumber=0

def setArrayTemp(data):
    global arrayTemp
    arrayTemp = data

def getArrayTemp(*args):
    global arrayTemp
    return arrayTemp

def setArrayResult(data):
    global arrayResult
    arrayResult = data

def getArrayResult(*args):
    global arrayResult
    return arrayResult

def setArrayNumber(data):
    global arrayNumber
    arrayNumber = int(data)

def getArrayNumber(*args):
    global arrayNumber
    return int(arrayNumber)


def specificReplace(data):
    #------- Страный пробел ----------#
    data = re.sub(r' ', ' ', data)
    data = re.sub(r'&nbsp;', ' ', data)
    data = re.sub(r'^\s*', '', data)
    return data








def gethtml(*args):

    try:
        data = HtmlClipboard.GetHtml()
        #data = data.replace(b'\r\n', b' ')
        data = data.decode()

        data = specificReplace(data)
        data = data.split('\r\n\r\n')

        setArrayTemp(data)
        setArrayNumber("1")
        labelStatusGetHtml['text'] = 'Данные получены'
        labelStatusAllElements['text'] = len(data)-1

        print("-----------in clipboard -----------")
        print(data)


        #win32clipboard.OpenClipboard()
        #win32clipboard.SetClipboardText(data, win32clipboard.CF_UNICODETEXT)
        #win32clipboard.CloseClipboard()

    except:
        labelStatusGetHtml['text'] = 'Данные отсутстуют'



def getElementbyNamber(*args):
    lenArray = len(getArrayTemp())
    n = getArrayNumber()
    print(n)

    if n < lenArray:
        textArea.delete('1.0', 'end')
        data = getArrayTemp()
        textArea.insert('1.0', data[int(n-1)])
        labelStatusWorks['text'] = n
        setArrayNumber(n+1)

def plusElementbyNamber(*args):
    lenArray = len(getArrayTemp())
    n = getArrayNumber()
    print(n)

    if n < lenArray:
        data = getArrayTemp()
        textArea.insert('end', data[int(n-1)])
        labelStatusWorks['text'] = n
        setArrayNumber(n+1)


######################################################
#####                 ОФОРМЛЕНИЕ                 #####
######################################################

root = Tk()
root.title("Чистить")
frame = Frame(root)

frameTop = Frame(frame)
frame.pack()
frameTop.pack(side=TOP, fill=X)

frameContent = Frame(frame)
frame.pack()
frameContent.pack(side=TOP, fill=X)

frameBottom = Frame(frame)
frame.pack()
frameContent.pack(side=BOTTOM, fill=X)

#-------------------------------#
#------- Верхнее поле ----------#
getContent = Button(frameTop, text='Получить данные', command=gethtml)
getContent.pack(side=LEFT, fill=X)

labelStatusGetHtml = Label(frameTop, text='')
labelStatusGetHtml.pack(side=LEFT, fill=X)

labelStatusText1 = Label(frameTop, text='Обработанно:')
labelStatusText1.pack(side=LEFT, fill=X)

labelStatusWorks = Label(frameTop, text='0')
labelStatusWorks.pack(side=LEFT, fill=X)

labelStatusText2 = Label(frameTop, text='из')
labelStatusText2.pack(side=LEFT, fill=X)

labelStatusAllElements = Label(frameTop, text='0')
labelStatusAllElements.pack(side=LEFT, fill=X)

refrashStatus = Button(frameTop, text='Обновить статус')
refrashStatus.pack(side=RIGHT)
#------- Верхнее поле ----------#
#-------------------------------#


#------------------------------------#
#------- Центрральное поле ----------#
leftMenu = Frame(frameContent)
frame.pack()
leftMenu.pack(side=LEFT, fill=Y)

getContent = Button(leftMenu, text='H1')
getContent.pack(side=TOP, fill=X)

getContent = Button(leftMenu, text='H2')
getContent.pack(side=TOP, fill=X)

getContent = Button(leftMenu, text='H3')
getContent.pack(side=TOP, fill=X)

getContent = Button(leftMenu, text='H4')
getContent.pack(side=TOP, fill=X)

getContent = Button(leftMenu, text='p')
getContent.pack(side=TOP, fill=X)

getContent = Button(leftMenu, text='ul')
getContent.pack(side=TOP, fill=X)

getContent = Button(leftMenu, text='ol')
getContent.pack(side=TOP, fill=X)

centerContent = Frame(frameContent)
frame.pack()
centerContent.pack(side=LEFT)

textArea = Text(centerContent)
textArea.pack(side=LEFT)
sbar = Scrollbar(centerContent)
sbar.config(command=textArea.yview)
textArea.config(yscrollcommand=sbar.set)
sbar.pack(side=RIGHT, fill=Y)


rightMenu = Frame(frameContent)
frame.pack()
rightMenu.pack(side=RIGHT, fill=Y)

getContent = Button(rightMenu, text='Получить элемент', command=getElementbyNamber)
getContent.pack(side=TOP, fill=X)

getContent = Button(rightMenu, text='+ Добавить', command=plusElementbyNamber)
getContent.pack(side=TOP, fill=X)

getContent = Button(rightMenu, text='- Удалить')
getContent.pack(side=TOP, fill=X)

getContent = Button(rightMenu, text='Сохранить')
getContent.pack(side=BOTTOM, fill=X)


#------- Центрральное поле ----------#
#------------------------------------#



#buttonFeature = Button(frametop, text="q", padx='35', pady='5')
#buttonFeature.pack(side='left')



root.resizable(0, 0)
root.mainloop()