from report import InputConnect

name = 'вакансии'
STAT = 'статистика'
INPUTS = (name, STAT)
TEXT_IN = 'Выводим вакансии или статистику?'
TEXT_ER = 'Некорректное значение ввода, попробуйте, ещё раз.'
NAME_FILE = '../vacancies_by_year.csv'
NAME_VAC = 'Программист'

if __name__ == '__main__':
    while True:
        res = str(input(TEXT_IN)).lower()
        if res.lower() in INPUTS:
            break
        print(TEXT_ER)

    inputconnect = InputConnect(NAME_FILE, NAME_VAC)

    if res == STAT:
        inputconnect.gen_stats(True)
    elif res == name:
        statist1, statist2, statist3, statist4, statist5, statist6 = inputconnect.gen_stats()
        inputconnect.gen_vac(statist1, statist2, statist3, statist4, statist5, statist6)
