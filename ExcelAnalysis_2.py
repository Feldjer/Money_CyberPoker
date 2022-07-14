import pandas as pd
import xlrd, xlwt
import openpyxl

# A program for analyzing the loss of the following combination

def excel_who_is_next(file, sheet_name):
    print('-' * 30 + '\nФайл: ' + file + '\nЛист: ' + sheet_name)
    determinant_exodus = {'П1': 0, 'П2': 1, 'Ничья': 2}
    mas_exodus = [['П1', [0]*3], ['П2', [0]*3], ['Ничья', [0]*3]]

    determinant_win_combo = {'Старшая карта': 0, 'Одна пара': 1, 'Две пары': 2, 'Сет': 3, 'Стрит': 4, 'Флеш': 5, 'Фул хаус': 6, 'Каре': 7, 'Стрит флеш': 8, 'Флеш рояль': 9}
    determinant_win_combo_reverse = ({values:keys for keys, values in determinant_win_combo.items()})
    mas_combo = [['Старшая карта', [0]*10], ['Одна пара', [0]*10], ['Две пары', [0]*10], ['Сет', [0]*10], ['Стрит', [0]*10],
             ['Флеш', [0]*10], ['Фул хаус', [0]*10], ['Каре', [0]*10], ['Стрит флеш', [0]*10], ['Флеш рояль', [0]*10]]

    cols_1, cols_2 = [2], [3]
    table_exodus = pd.read_excel(file, engine='openpyxl', usecols=cols_1, sheet_name=sheet_name)
    table_win_combo = pd.read_excel(file, engine='openpyxl', usecols=cols_2, sheet_name=sheet_name)

    for string in range(0, 150):
        exodus = str(table_exodus[string:string+1]).split()[-1]
        win_combo = str(table_win_combo[string:string+1]).split()[-1].replace('_', ' ')
        if string != 149:
            exodus_2 = str(table_exodus[string+1:string+2]).split()[-1]
            win_combo_2 = str(table_win_combo[string+1:string+2]).split()[-1].replace('_', ' ')
            for place in range(0, 3):
                if exodus == mas_exodus[place][0]:
                    mas_exodus[place][1][determinant_exodus.get(exodus_2)] = int(mas_exodus[place][1][determinant_exodus.get(exodus_2)]) + 1
                    break
            for place in range(0, 10):
                if win_combo == mas_combo[place][0]:
                    mas_combo[place][1][determinant_win_combo.get(win_combo_2)] = int(mas_combo[place][1][determinant_win_combo.get(win_combo_2)]) + 1
                    break

    print('-' * 30)
    for exodus in mas_exodus:
        print('[' + exodus[0] + '] П1: ' + str(exodus[1][0]) + ', П2: ' + str(exodus[1][1]) + ', Ничья: ' + str(exodus[1][2])) 
    print('-' * 30)
    for win_combo in mas_combo:
        count_position = 0
        string = ''
        string += '[' + win_combo[0] + '] '
        for combo in win_combo[1]:
            if combo != 0:
                string += str(determinant_win_combo_reverse.get(count_position)) + ': ' + str(combo) + ', '
            count_position += 1
        if ', ' in string:
            print(string[:-2])

    sentenceFirst = 'Предлагаю ставку на '
    sentenceSecond = ' раздачу '

    # This code is hidden for privacy and copyright reasons

if __name__ == "__main__":
    file = 'Cyber_Poker.xlsx'
    sheet_name = 'Cyber_poker_' + input("Какой лист нужен?\n1. Cyber_poker_1\n2. Cyber_poker_2\n3. Cyber_poker_3\n4. Cyber_poker_4\n5. Cyber_poker_5\n")
    excel_who_is_next(file, sheet_name)
