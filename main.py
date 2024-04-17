################################################
# Автор: Сомов Фёдор Михайлович
# Контакты: телефон/whatsapp = +79313429628
#			email = f.m.somov@ya.ru
#
# Лучше писать на WhatsApp, там отвечаю быстрее
#################################################
#
# Для работы необходимо положить файлы с соответствующими данным
# в папки. Наименование файлов неважно, расширение JSON для файлов вакансий
# Файл с вакансиями в папку: vacancies
#
# Ссылки, откуда были взяты последние файлы:
# Вакансии: https://opendata.trudvsem.ru/json/vacancy_5.json
#################################################

######################################
# Не забыть про pip!
# Импорт модулей:
# json - для работы с json файлами
# time - для отображения таймеров
# glob - для работы с папками
# pandas - для перевода словаря в дата фрейм и последующей записи в excel
# os - для работы с системными данными
######################################

import json
import time
import pandas as pd
import glob
import os
import openpyxl

# Предупреждаем пользователя, что бы не трогал без надобности.
print('------------------------------------------------',
      'Скрипт запущен. Пожалуйста, ожидайте сообщения о том, что "выполнение скрипта полностью завершено".',
      '------------------------------------------------', sep='\n')
# Задаем начало отсчета выполнения скрипта

tic = time.perf_counter()


def check_files():
    try:
        files = glob.glob(r'vacancies/*.json')
        files.sort(key=os.path.getmtime, reverse=True)
        latest_file = files[0]
        print('------------------------------------------------',
              'Файл в формате JSON найден.',
              '------------------------------------------------', sep='\n')
    except:
        print('------------------------------------------------',
              'Файлы в формате JSON не найдены. ПРОВЕРЬТЕ ФАЙЛ',
              '------------------------------------------------', sep='\n')
        raise SystemExit

    last_modified = os.path.getmtime(latest_file)
    rus_date = time.strftime('%d.%m.%Y %H:%M:%S', time.localtime(last_modified))
    print('-------------ДАТЫ ОБНОВЛЕНИЯ ФАЙЛОВ-------------')
    print('Самый свежий файл: ', latest_file)
    print('Дата последней модификации самого свежего файла: ', rus_date)

    return latest_file


def parser(file_path=None):
    """ Задаем список, в который будем записывать словари из json вакансий"""
    vac_data_list = []
    nodata = ''
    """ Для смены региона - поменять переменную region_code на код региона из справочника """
    region_code = "3200000000000"
    print('------------------------------------------------',
          'Начинаем разбор файла',
          '------------------------------------------------', sep='\n')
    with open(file_path, encoding='utf-8') as f:
        data_vac = json.load(f)

    for vac in data_vac['vacancies']:
        if region_code not in vac['state_region_code']:
            continue
        vac_dict = {'vac_id': vac.get('id', nodata), 'a_socialprotected': vac.get('social_protected_ids', nodata),
                    'a_jobname': vac.get('vacancy_name', nodata),
                    'a_specialization': vac.get('professionalSphereName', nodata),
                    'company_code': int(vac['company'].get('companycode', nodata)) if str(
                        vac['company'].get('companycode', nodata)).isdigit() else vac['company'].get('companycode',
                                                                                                     nodata),
                    'a_inn': int(vac['company'].get('inn', nodata)) if str(
                        vac['company'].get('inn', nodata)).isdigit() else vac['company'].get('inn', nodata),
                    'a_ogrn': int(vac['company'].get('ogrn', nodata)) if str(
                        vac['company'].get('ogrn', nodata)).isdigit() else vac['company'].get('ogrn', nodata),
                    'a_kpp': int(vac['company'].get('kpp', nodata)) if str(
                        vac['company'].get('kpp', nodata)).isdigit() else vac['company'].get('kpp', nodata),
                    'a_cname': vac['company'].get('name', nodata), 'a_education': vac.get('education', nodata),
                    'a_experience': int(vac.get('required_experience', nodata)) if str(
                        vac.get('required_experience', nodata)).isdigit() else vac.get('required_experience', nodata),
                    'a_shedule': vac.get('schedule_type', nodata), 'a_employment': vac.get('busy_type', nodata),
                    'a_salarymin': int(vac.get('salary_min', nodata)) if str(
                        vac.get('salary_min', nodata)).isdigit() else vac.get('salary_min', nodata),
                    'a_salarymax': int(vac.get('salary_max', nodata)) if str(
                        vac.get('salary_max', nodata)).isdigit() else vac.get('salary_max', nodata),
                    'a_vac_url': vac.get('vac_url', nodata), 'a_createdate': vac.get('date_create', nodata),
                    'code_profession': int(vac.get('code_profession', nodata)) if str(
                        vac.get('code_profession', nodata)).isdigit() else vac.get('code_profession', nodata),
                    'state_region_code': int(vac['state_region_code']) if str(
                        vac.get('state_region_code', nodata)).isdigit() else vac.get('state_region_code', nodata),
                    'vacancy_address': vac.get('vacancy_address', nodata),
                    'work_places': int(vac.get('work_places', nodata)) if str(
                        vac.get('work_places', nodata)).isdigit() else vac.get('work_places', nodata),
                    'c_date': vac.get('date_modify', nodata), 'is_quoted': str(vac.get('is_quoted', nodata)),
                    'original_source_type': (vac.get('original_source_type', nodata))}

        vac_data_list.append(vac_dict)
    print('------------------------------------------------',
          'Разбор файла закончен',
          'Переносим полученные данные в DataFramePandas',
          'И записываем данные в Excel',
          '------------------------------------------------', sep='\n')
    pd.DataFrame(vac_data_list).to_excel('output.xlsx', index=False)
    print('------------------------------------------------',
          'Данные записаны в Excel',
          '------------------------------------------------', sep='\n')

file_path = check_files()
parser(file_path)
toc = time.perf_counter()
print('------------------------------------------------',
      f'Скрипт выполнен за: {toc - tic:0.1f} секунд',
      'Файл output.xlsx в корневой папке скрипта',
      'выполнение скрипта полностью завершено',
      '------------------------------------------------', sep='\n')