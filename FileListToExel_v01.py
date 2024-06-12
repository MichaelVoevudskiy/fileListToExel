import os
import time
import xlsxwriter
import colorama
from colorama import Fore, Back
colorama.init()
os.system('cls')

selected_language = ''

print(Fore.LIGHTYELLOW_EX, '''
Copy all Filenames in a Folder to Excel
.---------------..---------------..---------------..---------------. 
|.-------------.||.-------------.||.-------------.||.-------------.|
||  _________  ||||   _____     ||||  _________  ||||  _________  ||
|| |_   ___  | ||||  |_   _|    |||| |  _   _  | |||| |_   ___  | ||
||   | |_  |_| ||||    | |      |||| |_| | | |_| ||||   | |_  |_| ||
||   |  _|     ||||    | |   _  ||||     | |     ||||   |  _|  _  ||
||  _| |_      ||||   _| |__/ | ||||    _| |_    ||||  _| |___/ | ||
|| |_____|     ||||  |________| ||||   |_____|   |||| |_________| ||
||             ||||             ||||             ||||             ||
|'-------------'||'-------------'||'-------------'||'-------------'|
'---------------''---------------''---------------''---------------' 
                                https://github.com/MichaelVoevudskiy
''')
time.sleep(0.5)


ua = {
    "annotMs" : '''Що я вмію робити:
    Я створюю перелік файлів із вказаної папки та завантажую її в Exel файл
    1. Вкажіть шлях до директорії з файлами
    2. Вкажіть назву файлу Exel''',
    'dirMs':'''Вкажіть шлях до директорії: ''',
    'dirExMs':'''(приклад:  D:\doc\main  )''',
    'dirDisMs':'''   Ти не вказав шлях до директорії.
        Створюю з директорії в якій я знаходжусь:'''
}

en = {
    "annotMs" : '''What I can do:
    I create a list of files from the specified folder and upload it to an 
    Excel file
    1. Specify the path to the directory with files
    2. Specify the name of the Excel file''',
    'dirMs':'''Send folder path: ''',
    'dirExMs':'''(Example:  D:\doc\main  )''',
    'dirDisMs':'''   You didn’t send the path to the folder.
        I take the path I'm in:'''
}


def choose_language():
    print(Fore.WHITE, "Choose language: ""[1] - Eng "" [2] - Ua", end=":  ")
    getleng = input()
    mess = check_language(getleng)

def check_language(getleng):
    global selected_language
    if (getleng=='2'):
        selected_language = ua
    elif (getleng=='1'):
        selected_language = en
    else:
        choose_language()


choose_language()
print(Fore.YELLOW, selected_language['annotMs'] )



def create_xlsx_file(file_path: str, headers: dict, items: list):
    # Створюємо новий файл XLSX
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    # Записуємо заголовки в перший рядок
    worksheet.write_row(row=0, col=0, data=headers.values())

    # Записуємо імена файлів у наступні рядки
    for index, item in enumerate(items):
        row = index + 1
        worksheet.write(row, 0, item)

    # Закриваємо файл
    workbook.close()

try:
    # Вкажіть шлях до папки, для якої потрібно створити список файлів
    print(Fore.WHITE, selected_language['dirMs'] ,Fore.YELLOW, selected_language['dirExMs'])
    print(Fore.WHITE)
    folder_path = input()

    if (folder_path == ''):
        folder_path = "./"
        scriptpatch = os.path.abspath(os.curdir)
        print(Fore.RED, selected_language['dirDisMs'],Fore.YELLOW,scriptpatch)
    time.sleep(0.5)

    print(Fore.WHITE,'''
    Як назвемо?:''')
    output_file_path = input()

    if (output_file_path == ''):
        output_file_path = "перелік_файлів.xlsx"
        print(Fore.RED,'''   Ти не вказав назву.''',Fore.WHITE,
        '''Назвемо файл''',Fore.YELLOW, output_file_path)
        
    else:
        output_file_path = folder_path + "/" + output_file_path +".xlsx"



    # Отримуємо список файлів у папці
    file_list = os.listdir(folder_path)

    # Заголовки для стовпців
    headers = {"Имена файлов": "Файли"}

    # Створюємо файл XLSX
    create_xlsx_file(output_file_path, headers, file_list)

    time.sleep(0.5)
    print(Fore.GREEN)

    print('''
    ____________________________________________________________
        Файл {} успішно створено
        Перелік файлів з папки {}
    ____________________________________________________________
        
        Гарного трудового дня!!!
        '''.format(output_file_path, folder_path))

except KeyboardInterrupt:
    print(Fore.RED,'Ви скасували операцію.',Fore.WHITE)

input()
