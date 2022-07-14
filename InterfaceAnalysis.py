from ExcelAnalysis_1 import load_sheet_names, excel_global
from ExcelAnalysis_2 import excel_who_is_next
import PySimpleGUI as sg

sg.ChangeLookAndFeel('GreenMono')

extensions = [("Excel (*.xlsx)", "*.xlsx"), ("Excel 2003 (*.xls)", "*.xls"), ("All files (*.*)", "*.*")]
sheet_name = ['Cyber_poker_1', 'Cyber_poker_2', 'Cyber_poker_3', 'Cyber_poker_4', 'Cyber_poker_5']
outcomes = ['Старшая карта', 'Одна пара', 'Две пары', 'Сет', 'Стрит', 'Флеш', 'Фул хаус', 'Каре', 'Стрит флеш', 'Флеш рояль', 'all']
length_selection = ['max', 'all', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17',
        '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30']
exodus, long, file, list_name = '', '', '', ''

layout = [
    [sg.Text('Исход'), sg.InputCombo((outcomes), default_value="Старшая карта", size=(14,1), key='-INPUT-'), sg.Text('Длина последовательности'),
     sg.InputCombo((length_selection), default_value="max", key='-LONG-', size=(4, 4), tooltip=('Подсказка')), sg.Button('Применить')],
    [sg.Text('Файл', size=(5, 1)), sg.InputText('Файл для анализа', size=(21, 4), key="-FILE-"),
     sg.FileBrowse(button_text = 'Выбрать', file_types=extensions), sg.InputCombo((sheet_name), default_value="Cyber_poker_1", size=(15,1), key='-LIST-'),
     sg.Button('Загрузить')], [sg.T('                       '),
     sg.Button("Старт анализа 1", button_color=("white", "green"), size=(12, 2)),
     sg.Button("Старт анализа 2", button_color=("white", "green"), size=(12, 2)),
    sg.Button("Выход", button_color=("white", "red"), size=(8, 2))],
    [sg.Output(size=(70, 20))]
]

window = sg.Window('CyberPokerAnalysis', layout)
while True:
    event, values = window.read()
    if event is None or event == 'Выход':
        break
    if event == 'Применить':
        exodus = values['-INPUT-']
        long = values['-LONG-']
        if exodus != '':
            if long != '':                
                sg.popup('Комбинация: ' + exodus + '\n' + 'Длина последовательности: ' + long)
            else:
                sg.popup('Не введена длина последовательности!')
        else:
            sg.popup('Не введена комбинация!')
    if event == 'Старт анализа 1':
        if (file != '') and (list_name != '') and (exodus != '') and (long != ''):
            excel_global(file, list_name, exodus, long)
        else:
            sg.popup('Введите данные для запуска!')
    elif event == 'Старт анализа 2':
        if (file != '') and (list_name != ''):
            excel_who_is_next(file, list_name)
        else:
            sg.popup('Введите данные для запуска!')
    elif event == 'Загрузить':
        file = ''
        file_path = values["-FILE-"]
        for i in file_path[::-1]:
            if i == "/":
                break
            else:
                file += i
        file = file[::-1]
        list_name = values["-LIST-"]
        if (file != "Файл для анализа данных") and (file != ''):
            sg.popup('Файл: ' + file + '\n' + 'Лист: ' + list_name)
        else:
            file = ''
            sg.popup("Выберите файл!")
window.close()
