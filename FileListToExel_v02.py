import os
import time
import xlsxwriter
import colorama
from colorama import Fore, Back
colorama.init()
os.system('cls')

selected_language = ''

ua = {
    "annotMs" : '''
    Що я вмію робити:
    Я створюю перелік файлів із вказаної папки та завантажую її в 
    Exel файл
    
    1. Вкажіть шлях до директорії з файлами
    2. Вкажіть назву файлу Exel
    
    Вихід - [ctr + C]
    ''',
    'dirMs':'''Вкажіть шлях до директорії: ''',
    'dirExMs':'''(приклад:  D:\doc\main  )''',
    'dirDisMs':'''    Ти не вказав шлях до директорії.
     Створюю з директорії в якій я знаходжусь:''',
    'fileNameMs': "перелік_файлів.xlsx", #
    'UserfileNameMs':'''   Ти не вказав назву.''',
    'DefFileNameMs':'''Назвемо файл''',
    'sendname':"Як назвемо?: "
}

en = {
    "annotMs" : '''
    What I can do:
    I create a list of files from the specified folder and upload 
    it to an Excel file
    
    1. Specify the path to the directory with files
    2. Specify the name of the Excel file
    
    Exit - [ctr + C]

    PS: I would be grateful for financial support
    BTC - bc1q84fp5ws486s73xmymdfena0yw05lqrgx87efjd
    ''',
    'dirMs':'''Send folder path: ''',
    'dirExMs':'''(Example:  D:\doc\main  )''',
    'dirDisMs':'''    You didn’t send the path to the folder.
     I take the path I'm in:''',
    'fileNameMs': "file_list.xlsx", #
    'UserfileNameMs':'''   You didn't specify a name.''',
    'DefFileNameMs':'''Let's name the file''',
    'sendname':'''Let's call it: '''
}

print(Fore.LIGHTYELLOW_EX, '''
Copy all Filenames in a Folder to Excel
.---------------..---------------..---------------..---------------. 
|.-------------.||.-------------.||.-------------.||.-------------.|
||  _________  ||||   _____     ||||  _________  ||||  _________  ||
|| |_   ___  | ||||  |_   _|    |||| |  _   _  | |||| |_   ___  | ||
||   | |_  |_| ||||    | |      |||| |_| | | |_| ||||   | |_  |_| ||
||   |  _|     ||||    | |   _  ||||     | |     ||||   |  _|  _  ||
||  _| |_      ||||   _| |__| | ||||    _| |_    ||||  _| |___| | ||
|| |_____|     ||||  |________| ||||   |_____|   |||| |_________| ||
||             ||||             ||||             ||||             ||
|'-------------'||'-------------'||'-------------'||'-------------'|
'---------------''---------------''---------------''---------------' 
                                https://github.com/MichaelVoevudskiy
''')
time.sleep(0.5)



def choose_language():
    print(Fore.WHITE, "Choose language: ",Fore.YELLOW,"[1] - Eng "" [2] - Ua", end=":  ")
    getleng = input()
    check_language(getleng)

def check_language(getleng):
    global selected_language
    if (getleng=='2'):
        selected_language = ua
    elif (getleng=='1'):
        selected_language = en
    else:
        choose_language()







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
    choose_language()
    print(Fore.YELLOW, selected_language['annotMs'] )
    # Вкажіть шлях до папки, для якої потрібно створити список файлів
    print(Fore.WHITE, selected_language['dirMs'] ,Fore.YELLOW, selected_language['dirExMs'])
    print(Fore.WHITE)
    folder_path = input()

    if (folder_path == ''):
        folder_path = "./"
        scriptpatch = os.path.abspath(os.curdir)
        print(Fore.RED, selected_language['dirDisMs'],Fore.YELLOW,scriptpatch)
    time.sleep(0.5)

    print(Fore.YELLOW)
    extension = "." + input("extension(exe, txt, png):")

    print(Fore.WHITE)
    output_file_path = input(selected_language['sendname'])

    if (output_file_path == ''):
        output_file_path = folder_path + "/" +selected_language['fileNameMs']
        print(Fore.RED, selected_language['UserfileNameMs'], Fore.WHITE)
        print(selected_language['DefFileNameMs'], Fore.YELLOW, output_file_path)
        
    else:
        output_file_path = folder_path + "/" + output_file_path +".xlsx"



    # Отримуємо список файлів у папці
    file_list = os.listdir(folder_path)



    file_list2 = []

    for file in file_list:
        if file.endswith(extension):
            file_list2.append(file)
            file_list = file_list2
   

    # Заголовки для стовпців
    headers = {"Ім'я файлів": "Файли"}


    # Створюємо файл XLSX
    create_xlsx_file(output_file_path, headers, file_list)

    time.sleep(0.5)
    print(Fore.GREEN)

    print(f'''
    ____________________________________________________________
        {output_file_path} successfully created
        File List from the folder {folder_path}
    ____________________________________________________________
        Have a nice day
        ''')

except KeyboardInterrupt:
    print(Fore.RED,'''
____________________________________________________________
                      have nice day
                       goodbye!!!
____________________________________________________________
        ''',Fore.WHITE)

input()
