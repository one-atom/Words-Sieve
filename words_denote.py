# -*- coding:utf8 -*-
import sys
import Tkinter
reload(sys)
import string

sys.setdefaultencoding('utf-8')

import xlrd

vocabulary = 'TOEFL'


def change(row_data_):
    row_data = row_data_.lower()
    change_form = []

    if row_data == '':
        return []
    change_form.append(row_data)
    if row_data[len(row_data) - 1] == 'y':
        change_form.append(row_data[0:len(row_data) - 1] + 'ie')
        change_form.append(row_data[0:len(row_data) - 1] + 'ies')
    if row_data[len(row_data) - 1] == 'f':
        change_form.append(row_data[0:len(row_data) - 1] + 'ves')
    if row_data[(len(row_data) - 2): (len(row_data) - 1)] == 'on' or row_data[
                                                                     (len(row_data) - 2): (len(row_data) - 1)] == 'un':
        change_form.append(row_data[0:len(row_data) - 2] + 'a')
    if row_data[(len(row_data) - 2): (len(row_data) - 1)] == 'us':
        change_form.append(row_data[0:len(row_data) - 2] + 'i')
    if row_data[len(row_data) - 1] == 'e':
        change_form.append(row_data[0:len(row_data) - 1] + 'ing')


def change_to_ori(row_data_):
    row_data = row_data_.lower()
    l = len(row_data)
    if row_data == '':
        return []
    ori = []
    ori.append(row_data)

    if l >= 3 and row_data[l - 3:l] == 'ies':
        ori.append(row_data[0:l - 3] + 'y')
    if l >= 2 and row_data[l - 2:l] == 'ie':
        ori.append(row_data[0:l - 2] + 'y')
    if l >= 3 and row_data[l - 3:l] == 'ves':
        ori.append(row_data[0:l - 3] + 'f')
    if row_data[-1] == 'a':
        ori.append(row_data[0:l - 1] + 'on')
        ori.append(row_data[0:l - 1] + 'um')
        ori.append(row_data[0:l - 1] + 'un')
    if row_data[-1] == 'i':
        ori.append(row_data[0:l - 1] + 'us')
    if l >= 3 and row_data[l - 3:l] == 'ing':
        ori.append(row_data[0:l - 3])
        ori.append(row_data[0:l - 3] + 'e')
    return ori


def binarysearch(array, low, high, target):

    if low > high:
        return -1

    mid = (low + high)/2
    if array[mid] > target:
        return binarysearch(array, low, mid - 1, target)
    if array[mid] < target:
        return binarysearch(array, mid + 1, high, target)
    if array[mid] == target:
        return mid

# read vocabulary data

fname = vocabulary + '.xls'
data = xlrd.open_workbook(fname)
vocabulary_sheet = data.sheet_by_name("Sheet1")
rows = vocabulary_sheet.nrows

word_list_ = []

for i in range(0, rows):
    word_list_.append(vocabulary_sheet.cell(i, 0).value)

word_list = []
for word in word_list_:
    aword = ''
    for s in word:
        if s!= u'\xa0' and s != ' ':
            aword += s
    if aword != '':
        word_list.append(aword)


def format_print(word, meaning):
    return word + ': ' + meaning


# search in toefl vocabulary

def denote_line(line, word_list):
    word = ''
    new_line = ''
    for alphabet in line:
        if ('a' <= alphabet <= 'z') or ('A' <= alphabet <= 'Z'):
            word += alphabet
        else:
            if word != '':
                change_list = change_to_ori(word)
                index = []
                for form in change_list:
                    po = binarysearch(word_list, 0, rows-1, form)
                    index.append(po)

                word_meaning = []
                for num in index:
                    if num != -1:
                        word_meaning.append((vocabulary_sheet.cell(num, 0).value, vocabulary_sheet.cell(num, 1).value))
                new_line = new_line + word
                if len(word_meaning) > 0:
                    new_line += '('
                    for wm in word_meaning:
                        new_line = new_line + format_print(wm[0], wm[1]) + '; '
                    new_line += ')'
            word = ''
            try:
                new_line += alphabet
            except:
                pass
    return new_line

from Tkinter import *

root = Tk()

frame1 = Frame(root)
frame1.pack(side=TOP)

content1 = StringVar()

t1 = Text(frame1, wrap=WORD,height=40)
t1.pack(side=LEFT)
scrollbar1 = Scrollbar(frame1, command=t1.yview)
t1.config(yscrollcommand=scrollbar1.set)
scrollbar1.pack(side=LEFT, fill=Y )

content2 = StringVar()
content2.set("这里将显示标注")
t2 = Text(frame1, wrap=WORD,height=40)
t2.pack(side=LEFT)
scrollbar2 = Scrollbar(frame1, command=t2.yview)
t2.config(yscrollcommand=scrollbar2.set)
scrollbar2.pack(side=LEFT, fill=Y)

frame2 = Frame(root)
frame2.pack(side=TOP)


# radiobutton action

def sel():
    global vocabulary
    global word_list
    global vocabulary_sheet
    global rows

    vocabulary = var.get()
    fname = vocabulary + '.xls'
    data = xlrd.open_workbook(fname)
    vocabulary_sheet = data.sheet_by_name("Sheet1")
    rows = vocabulary_sheet.nrows

    word_list_ = []

    for i in range(0, rows):
        word_list_.append(vocabulary_sheet.cell(i, 0).value)

    word_list = []
    for word in word_list_:
        aword = ''
        for s in word:
            if s!= u'\xa0' and s != ' ':
                aword += s
        if aword != '':
            word_list.append(aword)
    print 'Change to %s' % vocabulary

var = StringVar()
radiobutton1 = Radiobutton(frame2, text='TOEFL', variable=var, value='TOEFL', command=sel)
radiobutton1.pack(side=LEFT)
radiobuttion2 = Radiobutton(frame2, text='GRE', variable=var, value='GRE', command=sel)
radiobuttion2.pack(side=LEFT)


def button():
    passage = t1.get("1.0",END)
    new_passage = denote_line(passage, word_list)
    t2.delete("1.0",END)
    t2.insert("1.0",new_passage)
def clear():
    t1.delete("1.0",END)
    t2.delete("1.0",END)

button = Button(frame2, text='批注',command=button)
button2 = Button(frame2, text='清除',command=clear)
button.pack(side=LEFT)
button2.pack(side=LEFT)
root.mainloop()