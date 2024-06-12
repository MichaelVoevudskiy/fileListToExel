import os
import time
import xlsxwriter
import colorama
from colorama import Fore, Back
colorama.init()
os.system('cls')

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
    "annotation" : '''Що я вмію робити:
    Я створюю перелік файлів із вказаної папки та завантажую її в Exel файл
    1. Вкажіть шлях до директорії з файлами
    2. Вкажіть назву файлу Exel'''
}

en = {
    "annotation" : '''What I can do:
    I create a list of files from the specified folder and upload it to an Excel file
    1. Specify the path to the directory with files
    2. Specify the name of the Excel file'''
}


def choose_language():
    print(Fore.WHITE, "Choose language: ""[1] - Eng "" [2] - Ua", end=":  ")
    leng = input()
    mess = check_language(leng)
    return mess

def check_language(leng):
    if (leng=='2'):
        mess = ua["annotation"]
    elif (leng=='1'):
        mess = en["annotation"]
    else:
        choose_language()
    return mess


print(Fore.YELLOW, choose_language())



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
    print(Fore.WHITE, '''Вкажіть шлях до директорії: ''' ,Fore.YELLOW,"(приклад:  D:\doc\main  )" )
    print(Fore.WHITE)
    folder_path = input()

    if (folder_path == ''):
        folder_path = "./"
        scriptpatch = os.path.abspath(os.curdir)
        print(Fore.RED,'''   Ти не вказав шлях до директорії.
        Створюю з директорії в якій я знаходжусь:''',Fore.YELLOW,scriptpatch)
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
