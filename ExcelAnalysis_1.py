from collections import Counter as cn
import pandas as pd
import xlrd, xlwt
import openpyxl

# A program for analyzing the calculation of the loss of the desired combination

def load_sheet_names(file):
    sheet = pd.ExcelFile(file)
    return sheet.sheet_names

def excel_global(file, sheet_name, combination, long):
    sequence_lengths = []
    max_print = []
    cols = [3]
    table_exodus = pd.read_excel(file, engine='openpyxl', usecols=cols, sheet_name=sheet_name)

    print('-' * 30 + '\nФайл: ' + file + '\nЛист: ' + sheet_name + '\nКомбинация: ' + combination + '\nДлина: ' + long + '\n' + '-' * 30)

    if long == 'max':
        if combination != 'all':
            long = 100
        else:
            long = -100
    elif long == 'all':
        if combination != 'all':
            long = 888
        else:
            long = -888
    else:
        long = int(long)

    max = -999
    maximum = []
    all_exodus = []
    count_long = 0
    string_pr = ''
    search_long = 0
    all_combination = []
    combination_default = ''
    for data in range(0, 151):
        exodus = str(table_exodus[data:data+1])
        exodus = exodus.split()[-1].replace('_', ' ')
        if combination == 'all':
            if combination_default == '':
                combination_default = exodus
                count_long += 1
            elif exodus == combination_default:
                count_long += 1
            else:
                all_combination.append([combination_default, count_long])
                combination_default = exodus
                count_long = 1
        else:
            if exodus == combination:
                count_long += 1
            else:
                if count_long > 0:
                    if count_long == long:
                        search_long += 1
                    elif (long == 888) or (long == 100):
                        sequence_lengths.append(count_long)
                    count_long = 0
    if long == 888:
        c = cn(sequence_lengths)
        for k in sorted(c.keys()):
            print('Комбинация "' + combination + '" длиной ' + str(k) + ' повторяется ' + str(c[k]))
    elif long == 100:
        c = cn(sequence_lengths)
        for k in sorted(c.keys()):
            if max == -999:
                max = c[k]
                maximum.append([k, c[k]])
            else:
                if c[k] == max:
                    maximum.append([k, c[k]])
        for i in maximum:
            print('Максимальная комбинация "' + combination + '" длиной ' + str(i[0]) + ' повторяется ' + str(i[1]))
    elif combination == 'all':
        all_combination = sorted(all_combination)
        combination_default = ''
        count_repeat = 0
        for element in all_combination:
            if combination_default == '':
                combination_default = element
                count_repeat += 1
            elif combination_default == element:
                count_repeat += 1
            else:
                string_pr = 'Комбинация "' +  combination_default[0] + '" длиной ' + str(combination_default[1]) + ' повторяется ' + str(count_repeat)
                if long == -888:
                    print(string_pr)
                elif long == -100:
                    max_print.append(string_pr)
                combination_default = element
                count_repeat = 1
        check = 'Комбинация "' +  combination_default[0] + '" длиной ' + str(combination_default[1]) + ' повторяется ' + str(count_repeat)
        if (check != string_pr) and (long == -888):
            print(check)
        if (check != string_pr) and (long == -100):
            max_print.append(string_pr)
        if (combination == 'all') and (long == -100):
            k = 0
            mas_strings = []
            for string in max_print:
                string_found = string.split()
                if int(string_found[-1]) > max:
                    mas_strings = []
                    max = int(string_found[-1])
                    mas_strings.append(k)
                elif max == int(string_found[-1]):
                    mas_strings.append(k)
                k += 1
            for pr in mas_strings:
                print(max_print[pr])
    else:
        print('Комбинация "' + combination + '" длиной ' + str(long) + ' повторяется ' + str(search_long))

    sentenceFirst = 'Предлагаю ставку на '
    sentenceSecond = ' раздачу '

    # This code is hidden for privacy and copyright reasons

if __name__ == "__main__":
    string = ''
    sheets_name = load_sheet_names('Cyber_Poker.xlsx')
    print('Какой нужен лист?')
    for i in sheets_name:
        string += ', ' + i
    sheet_names = input('(' + string[2:] + ')' + ': ')
    combination = input('Какая выигрышная комбинация нужна?\n(Старшая карта, Одна пара, Две пары, Сет, Стрит, Флеш, Фул хаус, Каре, all): ')
    if combination != 'all':
        long = input('Какая длина последовательности нужна?\n(number/max/all): ')
    else:
        long = input('Какая длина последовательности нужна?\n(max/all): ')
    excel_global('Cyber_Poker.xlsx', sheet_names, combination, long)
