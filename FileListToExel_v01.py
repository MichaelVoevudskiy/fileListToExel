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

print(Fore.YELLOW,'''
+------------------------------------------------------------------+   
|    Що я вмію робити:                                             |
+------------------------------------------------------------------+''', Fore.LIGHTYELLOW_EX,'''
|    Я створюю Exel файл із переліком файлів.                      |
|    Використовуй мене для прорахунку замовлень                    |
|    або для створення комерційних пропозицій.                     |
+------------------------------------------------------------------+    
    ''')



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



    # Получаем список файлов в папке
    file_list = os.listdir(folder_path)

    # Заголовки для столбцов
    headers = {"Имена файлов": "Файли"}

    # Создаем файл XLSX
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
