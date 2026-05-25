import os

BASE_PATH_TEMPLATES = './Templates/'
BASE_PATH_DATA = './Data/'
BASE_PATH_REZULTS = './Results/'

SECRETARY_DEFAULT = "ст.преподаватель ОЦ8 Плотникова Н.О."  # для ФИИТ, ПМИ и др.
SECRETARY_PM = "вед.спец. по УМР ОЦ8 Юшкова И.О."          # для направления ПМ

DASH = "---"

GRADE_MAP = {'5': 'отлично', '4': 'хорошо', '3': 'удовлетворительно'}

MONTH_NAMES = {
    '01': 'января', '02': 'февраля', '03': 'марта', '04': 'апреля',
    '05': 'мая', '06': 'июня', '07': 'июля', '08': 'августа',
    '09': 'сентября', '10': 'октября', '11': 'ноября', '12': 'декабря'
}

# Пути к файлам
template_path_defense_4 = os.path.join(BASE_PATH_TEMPLATES, 'ЗащитыТаблицы4.docx')
template_path_defense_5 = os.path.join(BASE_PATH_TEMPLATES, 'ЗащитыТаблицы5.docx')
template_path_defense_6 = os.path.join(BASE_PATH_TEMPLATES, 'ЗащитыТаблицы6.docx')
template_path_defense_7 = os.path.join(BASE_PATH_TEMPLATES, 'ЗащитыТаблицы7.docx')
template_path_pred_bachelor = os.path.join(BASE_PATH_TEMPLATES, 'бакалавры_Протокол предзащиты ГЭК.docx')
template_path_pred_master   = os.path.join(BASE_PATH_TEMPLATES, 'магистры_Протокол предзащиты ГЭК.docx')
template_path_commission = os.path.join(BASE_PATH_TEMPLATES, 'Протокол комиссии.docx')

excel_dates = os.path.join(BASE_PATH_DATA, 'Даты.xlsx')
excel_programs = os.path.join(BASE_PATH_DATA, 'Программы.xlsx')
excel_students = os.path.join(BASE_PATH_DATA, 'Студенты.xlsx')
gek_docx = os.path.join(BASE_PATH_DATA, 'ГЭК_состав.docx')

output_dir_defense = os.path.join(BASE_PATH_REZULTS, 'DataZ/')
output_dir_pred = os.path.join(BASE_PATH_REZULTS, 'dataP/')
output_dir_protocols = os.path.join(BASE_PATH_REZULTS, 'protocols/')