import csv
import main
from datetime import datetime

now = datetime.now()
date_time = now.strftime("%d-%m-%Y %H:%M")

def create_futbool_csv():
    id_list, t1_list, t2_list, name_list = main.get_data()
    with open(f'./Футбол/Футбол{date_time}.csv', 'w') as file:
        writer = csv.writer(file)
        writer.writerow(
            ('НОМЕР ИГРЫ', 'ID ИГРЫ', 'ТЕКУЩЕЕ ВРЕМЯ', 'ВРЕМЯ ИГРЫ', 'ЛИГА', 'КОМАНДА 1', 'КОМАНДА 2')
        )

        id = 0
        for i in range(len(name_list)):
            if 'Футбол' in name_list[i] and not 'FIFA' in name_list[i]:
                writer.writerow(
                    (id, id_list[i], date_time , ' ',  name_list[i], t1_list[i], t2_list[i])
                )
                id += 1

def create_sport_csv(name):
    id_list, t1_list, t2_list, name_list = main.get_data()
    with open(f'./{name}/{name}{date_time}.csv', 'w') as file:
        writer = csv.writer(file)
        writer.writerow(
            ('НОМЕР ИГРЫ', 'ID ИГРЫ', 'ТЕКУЩЕЕ ВРЕМЯ', 'ВРЕМЯ ИГРЫ', 'ЛИГА', 'КОМАНДА 1', 'КОМАНДА 2')
        )

        id = 0
        for i in range(len(name_list)):
            if f'{name}' in name_list[i]:
                writer.writerow(
                    (id, id_list[i], date_time , ' ',  name_list[i], t1_list[i], t2_list[i])
                )
                id += 1

def games_create_csv():

    create_futbool_csv()
    create_sport_csv('Теннис')
    create_sport_csv('Гандбол')
    create_sport_csv('Баскетбол')
    create_sport_csv('Настольный теннис')
    create_sport_csv('Хоккей')
    create_sport_csv('FIFA')

#games_create_csv()



