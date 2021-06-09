from re import search
from PySimpleGUI.PySimpleGUI import Multiline
from pymongo import MongoClient
import pymongo
from sys import platform as show_platform
import os
import PySimpleGUI as sg
import pandas as pnd
import numpy as np
from tkinter.filedialog import askopenfilename
from tkinter import Tk
from datetime import datetime, timedelta
import sys


#to do: разделить на модули https://pythonworld.ru/osnovy/rabota-s-modulyami-sozdanie-podklyuchenie-instrukciyami-import-i-from.html
#
# Создание клиента
client = MongoClient('localhost', 27017)

# Подключение к БД (если БД не существует - она будет создана)
db = client['FMOPDB']

# Выбор коллекции
stud_collection = db['students']
duty_collection = db['duties']
duties_deleted_collection = db['duties_deleted']
stud_deleted_collection = db['students_deleted']

# удаление пробелов вначале и в конце всех импортируемых ячеек
def deletespaces(df):
    for i in range (len(df.columns)-1):
        for n in range (len(df[df.columns[i]])-1):
            if type(df[df.columns[i]][n])==str:
                df[df.columns[i]][n]=df[df.columns[i]][n].strip(' ')


# импорт студентов и их данных из excel
def studs_import(collection, file, change):
    duplicate_list=list()
    delstud_list=list()
    try:
        df2=pnd.read_excel(file)
        deletespaces(df2)
        slist = list(df2.iloc[:,0])
    except Exception as exception: 
        errordesc = sys.exc_info()
        return errordesc[1]
    if (len(list(set(slist)))!=len(slist)):
        return ('Номера личных дел дублируются, проверьте файл!')
    for i in range(len(slist)):
        stud=collection.find_one(
                {'_id': slist[i]}
                )
        if stud:
            duplicate_list.append(slist[i])
        delstud=stud_deleted_collection.find_one(
                {'_id': slist[i]}
                )
        if delstud:
            delstud_list.append(slist[i])
    if(duplicate_list):
        slist=list(set(slist) - set(duplicate_list))
    if(delstud_list):
        slist=list(set(slist) - set(delstud_list))
    if (change):
        for i in range(len(duplicate_list)): 
            gr=df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[1]]].values[0][0]
            collection.update_one(
                {
                '_id': duplicate_list[i]
                },
                {"$set": 
                {
                'last_name': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[2]]].values[0][0],
                'first_name': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[3]]].values[0][0],
                'full_name': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[2]]].values[0][0] + ' ' + df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[3]]].values[0][0],
                'sex': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[4]]].values[0][0],
                'citizenship': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[5]]].values[0][0],
                'phone': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[6]]].values[0][0],
                'e-mail': df2.loc[df2[df2.columns[0]] == duplicate_list[i]][[df2.columns[7]]].values[0][0],
                'faculty': gr.split('-')[0],
                'semester': gr.split('-')[1][0],
                'group': gr.split('-')[1][1:len(gr.split('-')[1])-1],
                'program':gr[len(gr)-1]}}
                )
    for i in range(len(slist)): 
            gr=df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[1]]].values[0][0]
            collection.insert_one(
                {'_id': slist[i],
                'last_name': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[2]]].values[0][0],
                'first_name': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[3]]].values[0][0],
                'full_name': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[2]]].values[0][0] + ' ' + df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[3]]].values[0][0],
                'sex': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[4]]].values[0][0],
                'citizenship': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[5]]].values[0][0],
                'phone': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[6]]].values[0][0],
                'e-mail': df2.loc[df2[df2.columns[0]] == slist[i]][[df2.columns[7]]].values[0][0],
                'faculty': gr.split('-')[0],
                'semester': gr.split('-')[1][0],
                'group': gr.split('-')[1][1:len(gr.split('-')[1])-1],
                'program':gr[len(gr)-1]}
                )
    if(duplicate_list and not change):
        return ('Студенты с номерами личных дел: ', duplicate_list, ' уже присутствуют в базе, остальные студенты добавлены')
    else: return ('Список студентов успешно импортирован')

def marklists_import(marklistcollection, studcollection, deldutiescollection, file):
    try:
        df=pnd.read_excel(file)
        deletespaces(df)
        slist = list(set(df.iloc[:,0]))
    except Exception as exception: 
        errordesc = sys.exc_info()
        return errordesc[1]
    unknownstudlist=list()
    for student in slist:
        existstud=studcollection.find_one({'full_name': student})
        if not existstud:
            unknownstudlist.append(student)
            df.drop(df.index[df[df.columns[0]]==student], axis = 0, inplace = True)
        else:
            df.loc[df[df.columns[0]] ==student, df.columns[0]] = existstud.get('_id')
    if (df.iloc[:,0].empty):
        return ('Карточки всех студентов в списке отсутствуют, сначала импортируйте их!')
    for i in range(len(df.iloc[:,0])):
        group=marklistcollection.find_one(
                {'students': df[df.columns[0]][i],
                'exam_group': df[df.columns[1]][i],
                'discipline':df[df.columns[3]][i],
                }
                )
        if group:
            continue
        group=deldutiescollection.find_one(
            {'students': df[df.columns[0]][i],
            'exam_group': df[df.columns[1]][i],
            'discipline':df[df.columns[3]][i],
                }
                )
        if group:
            continue
        marklistcollection.insert_one(
            {'students': df[df.columns[0]][i],
            'exam_group': df[df.columns[1]][i],
            'exam_faculty': df[df.columns[2]][i],
            'exam_semester': df[df.columns[1]][i].split('-')[1][0],
            'discipline':df[df.columns[3]][i],
            'exam_type':df[df.columns[4]][i],
            'mark':df[df.columns[5]][i]
            }
            ) 
    if (unknownstudlist):
        return ('Студенты: ', unknownstudlist, ' неизвестны, остальные оценки импортированы')
    return ('Листы с оценками успешно импортированы')

def getlist (collection, key, **searchcondition): 
    if (searchcondition):
        docslist=collection.find(searchcondition)
    else:
        docslist=collection.find()
    returnlist=list()
    if any(char=='+' for char in key):
        for doc in docslist:
            string=doc.get(key.split('+')[0])
            for i in range(1,len(key.split('+'))):
                string=string+ ' - ' + doc.get(key.split('+')[i])
            returnlist.append(string)
        if len(returnlist)>0:
            returnlist.sort()
        return (list(returnlist))
    if key=='fullgroup':
        for doc in docslist:
            try:
                appendix=doc.get('faculty')+'-'+doc.get('semester')+doc.get('group')+doc.get('program')
                returnlist.append(appendix)
            except Exception as exception: 
                errordesc = sys.exc_info()
                return errordesc[1]
    else:
        for doc in docslist:
            try:
                returnlist.append(doc.get(key))
            except Exception as exception: 
                errordesc = sys.exc_info()
                return errordesc[1]
    if len(returnlist)>0:
        returnlist.sort()
    return (list(returnlist))

def keytranslator(key):
    if key=='_id':
        return('Личное дело')
    if key=='last_name':
        return('Фамилия')
    if key=='first_name':
        return('Имя')
    if key=='full_name':
        return('Полное имя')
    if key=='sex':
        return('Пол')
    if key=='citizenship':
        return('Гражданство')
    if key=='phone':
        return('Телефон')
    if key=='students':
        return('Номер личного дела')
    if key=='exam_group':
        return('Экзаменационная группа')
    if key=='exam_faculty':
        return('Кафедра экзамена')
    if key=='exam_semester':
        return('Семестр экзамена')
    if key=='discipline':
        return('Дисциплина')
    if key=='exam_type':
        return('Тип')        
    if key=='mark':
        return('Оценка')  
    if key=='pass_date':
        return('Дата сдачи')
    if key=='ispassed':
        return('Долг сдан')   
    if key=='exclusion_date':
        return('Дата отчисления')      
    if key=='isgraduated':
        return('Закончил')  
    else:
        return (key)

def findstudent(collection, searchcondition):
    return(collection.find_one(searchcondition))

def getitem (collection, searchcondition):
    student=findstudent(collection, searchcondition)
    if not student:
        return
    messageform=False
    fullgroup=str()
    keys=list(student.keys())
    if any(field=='telegram token' for field in keys):
        keys.remove('telegram token')
        messageform=True
    if any(field=='full_name' for field in keys):
        keys.remove('full_name')
    if any(field=='faculty' for field in keys) and any(field=='semester' for field in keys) and any(field=='group' for field in keys) and any(field=='program' for field in keys):
        fullgroup=student.get('faculty')+'-'+student.get('semester')+student.get('group')+student.get('program')
        keys.remove('faculty')
        keys.remove('semester')
        keys.remove('group')
        keys.remove('program')
    if any(field=='program' for field in keys):
        keys.remove('program')
    returnstring=keytranslator(keys[0])+' : '+str(student.get(keys[0]))
    if (fullgroup):
        returnstring=returnstring+'\nНомер группы : '+fullgroup
    for i in range(1,len(keys)):
        returnstring=returnstring+'\n'+keytranslator(keys[i])+' : '+str(student.get(keys[i]))
    if collection==stud_collection:
        returnstring=returnstring+'\n'+'Долгов : '+str(len(list(duty_collection.find({'students':student.get('_id')}))))
    return([returnstring, messageform])

def movedoc(fromcollection, tocollection, document, fields, method):
        try: 
            fromcollection.delete_one(document)
        except Exception as exception: 
            errordesc = sys.exc_info()
            return errordesc[1]
        if method=='add':
            try:
                document.update(fields)
            except Exception as exception: 
                errordesc = sys.exc_info()
                return errordesc[1]
        if method=='del':
            try:
                for field in fields:
                    document.pop(field)
            except Exception as exception: 
                errordesc = sys.exc_info()
                return errordesc[1]        
        try: 
            tocollection.insert_one(document)
        except Exception as exception: 
            errordesc = sys.exc_info()
            return errordesc[1]

def pass_duty(searchcondition, date, ispassed):
    documents=list(duty_collection.find(searchcondition))
    if len(documents)>0:
        newfields={'pass_date': date, 'ispassed': ispassed}
        for document in documents:
            res=movedoc(duty_collection, duties_deleted_collection, document, newfields, 'add')
        return ('Выбранные долги перемещены в "сданные"')
    return "Ошибка, долг не найден!"

def excludestudent(searchcondition, date, isgraduated):
    documents=list(stud_collection.find(searchcondition))
    if len(documents)>0:
        for document in documents:
            if isgraduated:
                duties=duty_collection.find({'students':document.get('_id')})
                if duties.count()>0:
                    return ('Студент не может быть выпущен, у него есть долги!')
            newfields={'exclusion_date': date, 'isgraduated': isgraduated}
            pass_duty({'students':document.get('_id')}, date, isgraduated)
            movedoc(stud_collection, stud_deleted_collection, document, newfields, 'add')
        return ('Выбранные студенты перемещены в "отчисленные"')
    return ('Ошибка, студенты не найдены!')

def return_duty(searchcondition):
    documents=list(duties_deleted_collection.find(searchcondition))
    passed=list()
    if (len(searchcondition.items()))!=3:
        if len(documents)>0:
            delfields=['pass_date', 'ispassed']
            for document in documents:
                if document.get('ispassed')==False:
                    res=movedoc(duties_deleted_collection, duty_collection, document, delfields, 'del')
                else:
                    passed.append(document)
            if len(passed)>0:
                return ('Долг был сдан!')
            return ('Выбранные долги восстановлены!')
        return "Ошибка, долг не найден!"
    if (len(searchcondition.items()))==3:
        if len(documents)>0:
            delfields=['pass_date', 'ispassed']
            for document in documents:
                res=movedoc(duties_deleted_collection, duty_collection, document, delfields, 'del')
            return ('Выбранные долги восстановлены!')
        return "Ошибка, долг не найден!"

def returnstudent(searchcondition):
    documents=list(stud_deleted_collection.find(searchcondition))
    if len(documents)>0:
        for document in documents:
            delfields=['exclusion_date', 'isgraduated']
            return_duty({'students':document.get('_id')})
            movedoc(stud_deleted_collection, stud_collection, document, delfields, 'del')
        return ('Выбранные студенты восстановлены!')
    return ('Ошибка, студенты не найдены!')


def errorexplain (errortext):
    if str(errortext)=='single positional indexer is out-of-bounds':
        return 'Похоже, список пуст!'
    if str(errortext)=='list index out of range':
        return 'Похоже, столбцы расположены неверно!'
    if str(errortext)=='list index out of range':
        return 'Похоже, столбцы расположены неверно!'
    if str(errortext)[:14]=='No sheet named':
        return 'Листа с таким именем не существует!'
    else:
        return errortext

def getrelativechange(result1, result2):
    try:
        q=int(result1)/int(result2)*100-100
    except Exception as exception:
        return 'NaN'
    if q>0:
        return('+'+str(q))
    return str(q)

def statistics():
    week_ago=datetime.now().date()-timedelta(days=7)
    two_weeks_ago=datetime.now().date()-timedelta(days=14)
    month_ago=datetime.now().date()-timedelta(days=30)
    two_months_ago=datetime.now().date()-timedelta(days=60)
    stat=(getlist(duty_collection, 'students'))
    statlist=list()
    statlist.append(str(len(stat))) #0 - Общее количество долгов
    statlist.append(str(len(set(stat)))) #1 - Общее количество должников
    stat2=(getlist(duties_deleted_collection, 'students', pass_date= { '$gt': str(week_ago) } ))
    statlist.append(str(len(stat2)))  #2 - Уменьшение количества долгов за последнюю неделю
    stat3=set(stat2).difference(set(stat))
    statlist.append(str(len(stat3))) #3 - Уменьшение количества должников за последнюю неделю
    stat4=(getlist(duties_deleted_collection, 'students', pass_date= { '$gt': str(month_ago) } ))
    statlist.append(str(len(stat4))) #4 - Уменьшение количества долгов за последний месяц
    stat5=set(stat4).difference(set(stat))
    statlist.append(str(len(stat5))) #5 - Уменьшение количества должников за последний месяц
    stat6=(getlist(duties_deleted_collection, 'students', pass_date= { '$gt': str(two_weeks_ago) } ))
    statlist.append(str(len(stat6)-len(stat2))) #6 - Уменьшение количества долгов за предыдущую неделю
    stat7=(getlist(duties_deleted_collection, 'students', pass_date= { '$gt': str(two_months_ago) } ))
    statlist.append(str(len(stat7)-len(stat4))) #7 - Уменьшение количества долгов за предыдущий месяц
    statlist.append(getrelativechange(statlist[2], statlist[6])) #8 - Динамика за последние две недели, %
    statlist.append(getrelativechange(statlist[4], statlist[7])) #9 - Динамика за последние два месяца, %

    return statlist

def create_statistics_report():
    return ('Должников на факультете: '+statistics()[1]+
     '\nОбщее количество долгов: '+statistics()[0]+
     '\n\nИсправлено за предыдущую неделю: '+statistics()[6]+
     '\nИсправлено за последнюю неделю: '+statistics()[2]+
     '\nДинамика по отношению к предыдущей неделе: '+statistics()[8]+'%'+
     '\nДолжников за последнюю неделю стало меньше на: '+statistics()[3]+
     '\n\nИсправлено за предыдущий месяц: '+statistics()[7]+
     '\nИсправлено за последний месяц: '+statistics()[4]+
     '\nДинамика по отношению к предыдущему месяцу: '+statistics()[9]+'%'+
     '\nДолжников за последний месяц стало меньше на: '+statistics()[5])

def create_full_report():
    file = open('report'+str(datetime.now().date())+'.txt', 'w')
    file.write(create_statistics_report()+'\n\nДолжники:')
    studlist=set(getlist(duty_collection, 'students'))
    for stud in studlist:
        file.write('\n\n'+getitem(stud_collection, {'_id':stud})[0])
        dutylist=getlist(duty_collection, '_id', students=stud)
        for duty in dutylist:
            dutycard=getitem(duty_collection, {'_id':duty})[0].replace('\n', '\n\t')
            dutycard=dutycard.replace(dutycard.split('Экзаменационная группа :')[0], '')
            file.write('\n\n\t'+dutycard)
        



# Построение графического интерфейса
tab1_layout = [
    [sg.Button("Импортировать excel с данными студентов")],
    [sg.Button("Импортировать excel с оценками студентов")],
    [sg.Output(size=(106, 16), key = "Out", tooltip='Out')]]

col =   [
    [sg.Multiline(key='Studtext', size=[32, 14])],
    [sg.Text(key='Studtext2', size=[32,2])],
    [sg.Text("Сообщение:", key='Сообщение:', visible = False)],
    [sg.InputText(key="messagetext", size=(32, 1), visible = False)],
    [sg.Button("Отправить сообщение в Telegram", key='tgbutton', size=(33, 1), visible = False)],
    [sg.Button("Отчислить студента", size=(33, 1))]
    ]

facultylist=sorted(set(getlist(stud_collection, 'faculty')))
tab2_layout = [
 	[sg.Listbox(values=facultylist, key='flbox', size=(10, 20), enable_events =  True), 
     sg.Listbox(values=list(), key='glbox', size=(10, 20), enable_events =  True),
     sg.Listbox(values=list(), key='slbox', size=(40, 20), enable_events =  True),
     #sg.Frame('Студент:', frame_layout, size=(30, 20), font='Any 12', title_color='blue')],
     sg.Column(col, key='studcard', visible = False)],
     [sg.Text(str(len(facultylist))+' каф.', key='faccount', size=[12, 1]),
     sg.Text(key='groupcount', size=[12, 1]),
     sg.Text(key='studcount', size=[20, 1])
     ]
    ]

col2 =   [
    [sg.Multiline(key='dutytext', size=[30, 12])],
    [sg.Button("Перейти к карточке студента", size=(31, 1))],
    [sg.Text("Дата закрытия долга (гггг-мм-дд):", size=(30, 1))],
    [sg.InputText(key='passdate', size=(32, 1))],
    [sg.Button("Долг закрыт", size=(31, 1))]
    ]

def smlboxlist():
    smlboxlist=list()
    for stud in getlist(duty_collection, 'students'):
        smlboxlist=smlboxlist+getlist(stud_collection, 'full_name+_id', _id=stud)
    return (smlboxlist)


def fullrefresh(window, stud_collection, stud_deleted_collection):
    window.find_element('dutycard').Update(visible = False)
    window.find_element('dutydelcard').Update(visible = False)    
    window.find_element('studcard').Update(visible = False)
    window.find_element('studdelcard').Update(visible = False)
    window.FindElement('flbox').Update(sorted(set(getlist(stud_collection, 'faculty'))))
    window.FindElement('fdlbox').Update(sorted(set(getlist(stud_deleted_collection, 'faculty'))))
    window.FindElement('smlbox').Update(list(set(smlboxlist())))
    window.FindElement('smdlbox').Update(list(set(smdellboxlist())))
    window.find_element('dutytext').Update('')
    window.find_element('dutycard').Update('')
    window.find_element('dutydeltext').Update('')
    window.find_element('dutydelcard').Update('')
    window.find_element('studcard').Update('')
    window.find_element('studdelcard').Update('')
    window.find_element('qlbox').Update('')
    window.find_element('qdlbox').Update('')
    window.find_element('glbox').Update('')
    window.find_element('gdlbox').Update('')
    window.find_element('slbox').Update('')
    window.find_element('sdlbox').Update('')
    window.find_element('statistics').Update(create_statistics_report())

tab3_layout = [
    [sg.Listbox(values=list(set(smlboxlist())), key='smlbox', size=(30, 20), enable_events =  True), 
     sg.Listbox(values=list(), key='qlbox', size=(36, 20), enable_events =  True),
     sg.Column(col2, key='dutycard', visible=False)],
     [sg.Text('Должников: '+str(len(set(smlboxlist()))), key='smlcount', size=[31, 1]),
     sg.Text(key='qlcount', size=[31, 1]),
     sg.Text(key='dcount', size=[20, 1])
     ]
    ]


tab4_layout = [
     [sg.Multiline(create_statistics_report(), key='statistics', size=[106, 21])
     ],
     [sg.Button('Выгрузить отчет', size=(33,1))]
    ]

coldel =   [
    [sg.Multiline(key='Studdeltext', size=[32, 14])],
    [sg.Text(key='Studdeltext2', size=[32,2])],
    [sg.Text("Сообщение:", key='Сообщение:del', visible = False)],
    [sg.InputText(key="messagedeltext", size=(32, 1), visible = False)],
    [sg.Button("Отправить сообщение в Telegram", key='tgdelbutton', size=(33, 1), visible = False)],
    [sg.Button("Восстановить студента", size=(33, 1))]
    ]

facultydellist=sorted(set(getlist(stud_deleted_collection, 'faculty')))
tab5_layout = [
 	[sg.Listbox(values=facultydellist, key='fdlbox', size=(10, 20), enable_events =  True), 
     sg.Listbox(values=list(), key='gdlbox', size=(10, 20), enable_events =  True),
     sg.Listbox(values=list(), key='sdlbox', size=(40, 20), enable_events =  True),
     #sg.Frame('Студент:', frame_layout, size=(30, 20), font='Any 12', title_color='blue')],
     sg.Column(coldel, key='studdelcard', visible = False)],
     [sg.Text(str(len(facultydellist))+' каф.', key='facdelcount', size=[12, 1]),
     sg.Text(key='groupdelcount', size=[12, 1]),
     sg.Text(key='studdelcount', size=[20, 1])
     ]
    ]

def smdellboxlist():
    smdellboxlist=list()
    for stud in getlist(duties_deleted_collection, 'students'):
        smdellboxlist=smdellboxlist+getlist(stud_collection, 'full_name+_id', _id=stud)
        smdellboxlist=smdellboxlist+getlist(stud_deleted_collection, 'full_name+_id', _id=stud)
    return (smdellboxlist)

coldelduty =   [
    [sg.Multiline(key='dutydeltext', size=[30, 16])],
    [sg.Button("Восстановить долг", size=(29, 1))]
    ]

tab6_layout = [
    [sg.Listbox(values=list(set(smdellboxlist())), key='smdlbox', size=(30, 20), enable_events =  True), 
     sg.Listbox(values=list(), key='qdlbox', size=(36, 20), enable_events =  True),
     sg.Column(coldelduty, key='dutydelcard', visible=False)],
     [sg.Text('Должников: '+str(len(set(smdellboxlist()))), key='smdlcount', size=[31, 1]),
     sg.Text(key='qdlcount', size=[31, 1]),
     sg.Text(key='ddcount', size=[20, 1])
     ]
    ]

layout = [[sg.TabGroup([
    [sg.Tab('Импорт данных', tab1_layout, key='tab1'),
    sg.Tab('Просмотр контингента', tab2_layout, key='tab2'), 
    sg.Tab('Просмотр долгов', tab3_layout, key='tab3'),
    sg.Tab('Статистика на факультете', tab4_layout, key='tab4'),
    sg.Tab('Отчисленные студенты', tab5_layout, key='tab5'),
    sg.Tab('Закрытые долги', tab6_layout, key='tab6')]])]] 
window = sg.Window("Факультет ФМОП", layout)

# Привязка логики к интерфейсу
while True:
    event, values = window.read()
    root = Tk()
    root.withdraw()
    root.update()
    if event == "Импортировать excel с данными студентов":
        filename=askopenfilename()
        if (filename):
            window.FindElement('Out').Update('')
            print(errorexplain(studs_import(stud_collection, filename, True)))
            fullrefresh(window, stud_collection, stud_deleted_collection)
        else: 
            window.FindElement('Out').Update('')
    if event == "Импортировать excel с оценками студентов":
        filename=askopenfilename()
        if (filename):
            window.FindElement('Out').Update('')
            print(errorexplain(marklists_import(duty_collection, stud_collection, duties_deleted_collection, filename)))
            window.FindElement('flbox').Update(sorted(set(getlist(stud_collection, 'faculty'))))
            window.FindElement('smlbox').Update(list(set(smlboxlist())))
            fullrefresh(window, stud_collection, stud_deleted_collection) 
        else: 
            window.FindElement('Out').Update('')
    if event == "flbox":
        window.FindElement('glbox').Update(values=sorted(set(getlist(stud_collection, 'fullgroup', faculty=values['flbox'][0]))))
        window.FindElement('slbox').Update(values={})
        window.find_element('groupcount').Update(str(len(window.FindElement('glbox').Values))+' гр.')
        window.find_element('studcount').Update('')
        window.find_element('studcard').Update(visible = False)
    if event == "glbox":
        if (values['glbox']):
            window.FindElement('slbox').Update(values=getlist(stud_collection,
                                    'full_name+_id', 
                                    faculty=values['glbox'][0].split('-')[0], 
                                    semester= values['glbox'][0].split('-')[1][0],
                                    group=values['glbox'][0].split('-')[1][1:len(values['glbox'][0].split('-')[1])-1],
                                    program=values['glbox'][0][len(values['glbox'][0])-1]))
            window.find_element('studcount').Update(str(len(window.FindElement('slbox').Values))+' студ.')
            window.find_element('studcard').Update(visible = False)
    if event == "slbox":
        if (values['slbox']):
            window.find_element('studcard').Update(visible = True)
            slboxvalue=getitem(stud_collection, {'_id':(values['slbox'][0].split('- ')[1])})
            window.find_element('Studtext').Update(slboxvalue[0])
            if slboxvalue[1]==True:
                window.find_element('tgbutton').Update(visible = True)
                window.find_element('messagetext').Update(visible = True)
                window.find_element('Сообщение:').Update(visible = True)
            else: 
                window.find_element('tgbutton').Update(visible = False)
                window.find_element('messagetext').Update(visible = False)
                window.find_element('Сообщение:').Update(visible = False)
    if event == 'smlbox':
        if (values['smlbox']):
            smlboxvalue=getlist(duty_collection,'exam_semester+discipline', students=values['smlbox'][0].split('- ')[1])
            for i in range(len(smlboxvalue)):
                smlboxvalue[i]=smlboxvalue[i].split(' - ')[0]+' сем. '+smlboxvalue[i].split(' - ')[1]
            window.find_element('qlbox').Update(values=smlboxvalue)
            window.find_element('qlcount').Update('Долгов: '+str(len(smlboxvalue)))
            window.find_element('dutycard').Update(visible = False)
    if event == 'qlbox':            
        if values['qlbox']:
            smlboxvalue=getitem(duty_collection, {'students':values['smlbox'][0].split('- ')[1], 
                                                'exam_semester':values['qlbox'][0][0],
                                                'discipline':values['qlbox'][0].split('сем. ')[1]})
            window.find_element('dutytext').Update(smlboxvalue[0].split('\n',1)[1])
            window.find_element('dutycard').Update(visible = True)
    if event == 'Перейти к карточке студента':
        window.Element('tab2').select()
        window.find_element('studcard').Update(visible = True)
        slboxvalue=getitem(stud_collection, {'_id':values['smlbox'][0].split('- ')[1]})
        window.find_element('Studtext').Update(slboxvalue[0])
    if event == 'Долг закрыт':
        searchcondition={'students':values['smlbox'][0].split('- ')[1], 
                        'exam_semester':values['qlbox'][0][0],
                        'discipline':values['qlbox'][0].split('сем. ')[1]}
        result=pass_duty(searchcondition, values['passdate'], True)
        window.find_element('dutytext').Update(result)
        fullrefresh(window, stud_collection, stud_deleted_collection)
    if event == 'Отчислить студента':
        result=excludestudent({'_id': values['Studtext'].split('Личное дело : ')[1].split('\n')[0]}, str(datetime.now().date()), False)
        window.find_element('Studtext').Update(result)
        fullrefresh(window, stud_collection, stud_deleted_collection)
    if event == "fdlbox":
        if values['fdlbox']:
            window.FindElement('gdlbox').Update(values=sorted(set(getlist(stud_deleted_collection, 'fullgroup', faculty=values['fdlbox'][0]))))
            window.FindElement('sdlbox').Update(values={})
            window.find_element('groupdelcount').Update(str(len(window.FindElement('gdlbox').Values))+' гр.')
            window.find_element('studdelcount').Update('')
            window.find_element('studdelcard').Update(visible = False)
    if event == "gdlbox":
        if (values['gdlbox']):
            window.FindElement('sdlbox').Update(values=getlist(stud_deleted_collection,
                                    'full_name+_id', 
                                    faculty=values['gdlbox'][0].split('-')[0], 
                                    semester= values['gdlbox'][0].split('-')[1][0],
                                    group=values['gdlbox'][0].split('-')[1][1:len(values['gdlbox'][0].split('-')[1])-1],
                                    program=values['gdlbox'][0][len(values['gdlbox'][0])-1]))
            window.find_element('studdelcount').Update(str(len(window.FindElement('sdlbox').Values))+' студ.')
            window.find_element('studdelcard').Update(visible = False)
    if event == "sdlbox":
        if (values['sdlbox']):
            window.find_element('studdelcard').Update(visible = True)
            sdlboxvalue=getitem(stud_deleted_collection, {'_id':(values['sdlbox'][0].split('- ')[1])})
            window.find_element('Studdeltext').Update(sdlboxvalue[0])
            if sdlboxvalue[1]==True:
                window.find_element('tgdelbutton').Update(visible = True)
                window.find_element('messagedeltext').Update(visible = True)
                window.find_element('Сообщение:del').Update(visible = True)
            else: 
                window.find_element('tgdelbutton').Update(visible = False)
                window.find_element('messagedeltext').Update(visible = False)
                window.find_element('Сообщение:del').Update(visible = False)
    if event == 'Восстановить студента':
        result=returnstudent({'_id': values['Studdeltext'].split('Личное дело : ')[1].split('\n')[0]})
        window.find_element('Studtext').Update(result)
        fullrefresh(window, stud_collection, stud_deleted_collection)
    if event == 'smdlbox':
        if (values['smdlbox']):
            smdlboxvalue=getlist(duties_deleted_collection,'exam_semester+discipline', students=values['smdlbox'][0].split('- ')[1])
            for i in range(len(smdlboxvalue)):
                smdlboxvalue[i]=smdlboxvalue[i].split(' - ')[0]+' сем. '+smdlboxvalue[i].split(' - ')[1]
            window.find_element('qdlbox').Update(values=smdlboxvalue)
            window.find_element('qdlcount').Update('Долгов: '+str(len(smdlboxvalue)))
            window.find_element('dutydelcard').Update(visible = False)
    if event == 'qdlbox': 
        if values['qdlbox']:
            smlboxvalue=getitem(duties_deleted_collection, {'students':values['smdlbox'][0].split('- ')[1], 
                                                'exam_semester':values['qdlbox'][0][0],
                                                'discipline':values['qdlbox'][0].split('сем. ')[1]})
            window.find_element('dutydeltext').Update(smlboxvalue[0].split('\n',1)[1])
            window.find_element('dutydelcard').Update(visible = True)
    if event == 'Восстановить долг':
        searchcondition={'students':values['smdlbox'][0].split('- ')[1], 
                            'exam_semester':values['qdlbox'][0][0],
                            'discipline':values['qdlbox'][0].split('сем. ')[1]}
        result=return_duty(searchcondition)
        window.find_element('dutydeltext').Update(result)
        fullrefresh(window, stud_collection, stud_deleted_collection)
    if event == 'Выгрузить отчет':
        create_full_report()
    if event in (None, "Exit",  "Cancel"):
        break