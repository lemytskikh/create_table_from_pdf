import pandas as pd
import numpy as np
import os
import pdfplumber
import PyPDF2
import json

# Ввод пути к файлам, названия файлов, подготовительный этап

direct = r'{}'.format(input('Введите путь к папке с файлами: '))
new_old = input('Добавить в старый файл? (Отвечать да или нет): ')
if new_old == 'да':
    old_file = r'{}'.format(input('Укажите путь к старому файлу: '))
    old_file_name = input('Укажите название файла: ')
    
    if 'xlsx' in old_file_name:
        old_file = old_file + '\{}'.format(old_file_name)
    else:
        old_file = old_file + '\{}.xlsx'.format(old_file_name)
else:
    new_file = r'{}'.format(input('Укажите куда сохарнить файл: '))
    new_file_name = input('Укажите название файла: ')
    
    if 'xlsx' in new_file_name:
        new_file = new_file + '\{}'.format(new_file_name)
    else:
        new_file = new_file + '\{}.xlsx'.format(new_file_name)
    
# direct = r'C:\Users\monot\Favorites\Jupyter\Эксперименты\Тюрин\2612\Резюме'
#r'D:\LEM\Documents\Jupyter\Эксперименты\Тюрин\2612\файлы'

# Цикл создания пути к файлам

nf = os.listdir(direct)

way = []
for i in range(len(nf)):
    try:
        nfile = []
        nfile = os.listdir(direct + '\{}'.format(nf[i]))
        for p in nfile:
            way_z = direct + '\{}\{}'.format(nf[i], p)
            way.append(way_z)
    except:
        way.append(direct + f'\{nf[i]}')
    
print('Количество файлов для обработки: ', len(way))


# Цикл сбора файлов в таблицу

final = pd.DataFrame()
#zp = pd.DataFrame()
last_column = []
x = 0
err = []

for wz in way:
    try:
        pdf_file = open(wz, 'rb')
        read_pdf = PyPDF2.PdfReader(pdf_file)
        number_of_pages = read_pdf.pages
        page_all = []
        for nop in range(len(number_of_pages)):
            page = read_pdf.pages[nop]
            page_all.append(page.extract_text())

        page_all1 = ' '.join(page_all)

        data = json.dumps([page_all1])
        formatj = json.loads(data)

        #df_zp1 = pd.DataFrame()
        df_1 = pd.DataFrame()
        df_1['nh'] = formatj
        df = df_1['nh'].str.split('\n', expand=True, n=9)
        df[10] = wz
        final = pd.concat([final, df])
        #df_zp1['nh'] = formatj
        #df_zp = df_zp1['nh'].str.split('\n', expand=True, n=15)
        #zp = pd.concat([zp, df_zp])

        last_column.append(page_all[0].split('\n')[-1])
    except:
        err.append(wz)
    
    print(x)
    x += 1
    
# Обработка данных из после цикла, переименование столбцов удаление не нужных данных
   
final['обнавлено'] = last_column
#final['резюме'] = way
final = final.reset_index(drop=True)

final = final.rename(columns={0: 'ФИО', 1: 'пол_возраст_др', 2: 'телефон', 3: 'email', 4: 'место_проживания',
                      5: 'граждаство', 6: 'вид_занятости', 7: 'drop', 8: 'должность', 9: 'опыт_работы', 
                             10: 'откуда_файл'})

final['место_проживания'] = final['место_проживания'].str.replace('Проживает: ', '', regex=True)
final['граждаство'] = final['граждаство'].str.replace('Гражданство: ', '', regex=True)
final['опыт_работы'] = final['опыт_работы'].str.replace('\n', '  ', regex=True)
# display(final)

# final.to_excel(r'D:\LEM\Documents\Jupyter\Эксперименты\Тюрин\2612\сбор.xlsx', index=False)

# Выделим название папки откуда файл и название файла

final[['del1', 'del2', 'del3', 'del4', 'del5', 'del6', 
       'del7', 'del8', 'del9', 'del10']] = final['откуда_файл'].str.split('\\', expand = True, n = 9)
final = final.drop(columns = ['del1', 'del2', 'del3', 'del4', 'del5', 'del6', 'del7', 'del8', 'откуда_файл'])
final = final.rename(columns = {'del9': 'название_папки(откуда_файл)', 'del10': 'название_файла'})
# final

# Сохраним полученные данные

if new_old == 'да':
    df = pd.read_excel(old_file)
    df = pd.concat([df, final])
    df.to_excel(old_file, index = False)
else:
    final.to_excel(new_file, index = False)
    
if new_old == 'да':   
    print('Программа закончила работу. Проверьте наличие файла по адресу:')
    print(old_file)
else:
    print('Программа закончила работу. Проверьте наличие файла по адресу:')
    print(new_file)


input('Для закрытия нажми "ENTER"')





