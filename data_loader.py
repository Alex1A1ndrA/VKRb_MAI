import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from docx import Document

from constants import (
    excel_programs, excel_commissions, excel_dates, excel_students, gek_docx,
    template_path_defense_5, template_path_defense_4,
    template_path_pred_bachelor, template_path_pred_master, template_path_commission,
    output_dir_defense, output_dir_pred, output_dir_protocols
)
from utils import extract_group, safe_date_conversion

@st.cache_data
def get_programs_data():
    try:
        df = pd.read_excel(excel_programs, sheet_name='Лист1')
    except Exception as e:
        st.error(f"Ошибка загрузки файла программ: {e}")
        return {}
    df.columns = ['Код', 'Направление', 'Профиль', 'Квалификация',
                  'Уровень', 'Год_поступления', 'Группы']
    df = df.dropna(subset=['Группы'])
    df['Год_поступления'] = pd.to_numeric(df['Год_поступления'], errors='coerce')
    df = df.dropna(subset=['Год_поступления'])
    programs = {}
    for _, row in df.iterrows():
        code = str(row['Код']).strip()
        direction = str(row['Направление']).strip()
        profile = str(row['Профиль']).strip()
        level = str(row['Уровень']).strip()
        qualification = str(row['Квалификация']).strip()
        year = int(row['Год_поступления'])
        groups_str = str(row['Группы']).strip()
        if not groups_str:
            continue
        for part in groups_str.replace(' ', '').split(','):
            group = part.strip()
            if group:
                programs.setdefault(group, []).append({
                    'code': code,
                    'direction': direction,
                    'profile': profile,
                    'level': level,
                    'qualification': qualification,
                    'year': year
                })
    for group in programs:
        programs[group].sort(key=lambda x: x['year'], reverse=True)
    return programs

def get_program_info(group_num, year=None):
    programs_data = get_programs_data()
    if group_num not in programs_data:
        return None
    records = programs_data[group_num]
    if year is not None:
        for rec in records:
            if rec['year'] == year:
                return rec
    return records[0]

def get_program_description(level: str) -> str:
    level = level.strip()
    if level == 'СВО':
        return 'программа специализированного высшего образования — магистратуры'
    elif level == 'БВО':
        return 'программа базового высшего образования — бакалавриата'
    elif level == 'Магистратура':
        return 'магистерская программа'
    elif level == 'Бакалавриат':
        return 'профиль'
    else:
        return ''

def qualification_genitive(qual: str) -> str:
    qual_lower = qual.lower().strip()
    if qual_lower == 'бакалавр':
        return 'БАКАЛАВРА'
    elif qual_lower == 'магистр':
        return 'МАГИСТРА'
    elif qual_lower == 'инженер':
        return 'ИНЖЕНЕРА'
    elif qual_lower == 'инженер-исследователь':
        return 'ИНЖЕНЕРА-ИССЛЕДОВАТЕЛЯ'
    else:
        return qual.upper()

def get_group_full_info(group_num, year=None):
    programs_data = get_programs_data()
    if group_num not in programs_data:
        return group_num, ''
    info = get_program_info(group_num, year)
    if info is None:
        info = get_program_info(group_num)
    if info is None:
        return group_num, ''
    level = info['level']
    year_used = info['year']
    year_str = str(year_used)[-2:]
    if level == 'БВО':
        suffix = 'БВ'
    elif level == 'СВО':
        suffix = 'СВ'
    elif level == 'Магистратура':
        suffix = 'М'
    elif level == 'Бакалавриат':
        suffix = 'Б'
    else:
        suffix = ''
    full_group_name = f"М8О-{group_num}{suffix}-{year_str}"
    first_digit = int(str(group_num)[0])
    graduation_year = year_used + first_digit
    return full_group_name, str(graduation_year)

@st.cache_data
def get_commission_data():
    try:
        df = pd.read_excel(excel_commissions, sheet_name='Лист1')
    except Exception as e:
        st.error(f"Ошибка загрузки комиссии: {e}")
        return {}
    commission = {}
    for _, row in df.iterrows():
        fam = str(row.get('Фамилия', '')).strip()
        if not fam:
            continue
        role = str(row.get('Должность', '')).strip().lower()
        degree = str(row.get('Степень', '')).strip().lower()
        title = str(row.get('Звание', '')).strip().lower()
        can_be_chair = False
        if any(word in role for word in ['профессор', 'зав', 'заведующий']):
            can_be_chair = True
        if 'доктор' in degree:
            can_be_chair = True
        if title == 'профессор':
            can_be_chair = True

        is_docent = 'доцент' in role

        im = str(row.get('Имя', '')).strip()
        ot = str(row.get('Отчество', '')).strip()
        initials = ''
        if im:
            initials += im[0].upper() + '.'
        if ot:
            initials += ot[0].upper() + '.'

        if role:
            full_str = f"{role} каф. 806 {fam} {initials}".strip()
        else:
            full_str = f"{fam} {initials}".strip()

        commission[fam] = (full_str, can_be_chair, is_docent)
    return commission

def get_commission_data_norm():
    commission_data = get_commission_data()
    return {fam.lower().rstrip('.'): value for fam, value in commission_data.items()}

@st.cache_data
def get_gek_members():
    if not os.path.exists(gek_docx):
        st.warning(f"Файл {gek_docx} не найден. Состав ГЭК не загружен.")
        return {}

    def normalize_profile(text):
        if not isinstance(text, str):
            text = str(text)
        text = re.sub(r'[«»"*]', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text.lower()

    def make_short(full):
        parts = full.split()
        if len(parts) >= 3:
            return f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
        elif len(parts) == 2:
            return f"{parts[0]} {parts[1][0]}."
        else:
            return parts[0]

    def extract_commission(lines, start):
        idx = start
        while idx < len(lines) and 'Экзаменационная комиссия' not in lines[idx]:
            if ('Профиль:' in lines[idx] or 'Магистерская программа' in lines[idx] or 
                ('Направление' in lines[idx] and re.search(r'\d{2}\.\d{2}\.\d{2}', lines[idx]))):
                return [], None, idx
            idx += 1
        if idx >= len(lines):
            return [], None, idx
        idx += 1
        members = []
        while idx < len(lines):
            cur_line = lines[idx]
            if ('Профиль:' in cur_line or 'Магистерская программа' in cur_line or
                ('Направление' in cur_line and re.search(r'\d{2}\.\d{2}\.\d{2}', cur_line))):
                break
            if 'Секретарь:' in cur_line:
                idx += 1
                break
            if ('–' in cur_line or '-' in cur_line) and not any(x in cur_line for x in ['Экзаменационная комиссия', 'Секретарь:', 'Магистерская программа:', 'Профиль:']):
                fio_part = re.split(r'[–-]', cur_line)[0].strip()
                short = make_short(fio_part)
                members.append(short)
            idx += 1
        return members, None, idx

    doc = Document(gek_docx)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    gek_members = {}
    i = 0
    while i < len(lines):
        line = lines[i]
        if 'Направление' in line and re.search(r'\d{2}\.\d{2}\.\d{2}', line):
            code_match = re.search(r'(\d{2}\.\d{2}\.\d{2})', line)
            if not code_match:
                i += 1
                continue
            current_code = code_match.group(1)

            prof_match = re.search(r'(?:Профиль:|Магистерская программа:?)\s*«?([^»]+)»?', line, re.IGNORECASE)
            if prof_match:
                profile = prof_match.group(1).strip()
                members, _, next_i = extract_commission(lines, i + 1)
                if members:
                    norm_profile = normalize_profile(profile)
                    key = (current_code, norm_profile)
                    gek_members[key] = members
                i = next_i
                continue
            else:
                i += 1
                while i < len(lines):
                    next_line = lines[i]
                    if 'Направление' in next_line and re.search(r'\d{2}\.\d{2}\.\d{2}', next_line):
                        break
                    prof_match = re.search(r'(?:Профиль:|Магистерская программа:?)\s*«?([^»]+)»?', next_line, re.IGNORECASE)
                    if prof_match:
                        profile = prof_match.group(1).strip()
                        members, _, next_i = extract_commission(lines, i + 1)
                        if members:
                            norm_profile = normalize_profile(profile)
                            key = (current_code, norm_profile)
                            gek_members[key] = members
                        i = next_i
                        continue
                    i += 1
                continue
        i += 1
    return gek_members

@st.cache_data
def get_dates_defense():
    try:
        df = pd.read_excel(excel_dates, sheet_name='Даты защит')
    except Exception as e:
        st.error(f"Ошибка загрузки дат защит: {e}")
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    expected_cols = ['Направление', 'Группа_полная', 'Форма_обучения',
                     'Даты_основные', 'Время', 'Аудитория', 'Председатель']
    if len(df.columns) == 7:
        df.columns = expected_cols
    else:
        if len(df.columns) > 7:
            df = df.iloc[:, :7]
            df.columns = expected_cols
        else:
            st.error(f"Неожиданное число колонок в Датах защит: {len(df.columns)}")
            return pd.DataFrame()
    df['Группа_номер'] = df['Группа_полная'].apply(lambda x: extract_group(x) if pd.notna(x) else None)
    df['Год'] = df['Группа_полная'].apply(lambda x: x.split('-')[-1] if pd.notna(x) and '-' in x else None)
    df['Время'] = df['Время'].astype(str).str.strip()
    df['Код_направления'] = df['Направление'].apply(lambda x: x.split()[0] if pd.notna(x) else None)
    return df

@st.cache_data
def get_dates_predefense():
    try:
        df = pd.read_excel(excel_dates, sheet_name='Даты предзащит')
    except Exception as e:
        st.error(f"Ошибка загрузки дат предзащит: {e}")
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    expected_cols = ['Группа_короткая', 'День_недели', 'Дата_предзащиты', 'Комиссия', 'Аудитория']
    if len(df.columns) == 5:
        df.columns = expected_cols
    else:
        if len(df.columns) > 5:
            df = df.iloc[:, :5]
            df.columns = expected_cols
        else:
            st.error(f"Неожиданное число колонок в Датах предзащит: {len(df.columns)}")
            return pd.DataFrame()
    df['Группа_номер'] = df['Группа_короткая'].apply(lambda x: extract_group(x) if pd.notna(x) else None)
    return df

@st.cache_data
def get_students():
    try:
        df = pd.read_excel(excel_students, sheet_name='итог')
    except Exception as e:
        st.error(f"Ошибка загрузки студентов: {e}")
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    expected_cols = ['ФИО', 'Группа', 'Тема', 'Руководитель', 'Консультант', 'Рецензент',
                     'Присутствие_пред', 'Оценка_пред', 'Новая_тема', 'Новый_рук',
                     'Присутствие_защ', 'Листов', 'Оценка_защ']
    if len(df.columns) == 13:
        df.columns = expected_cols
    else:
        st.error(f"Неожиданное число колонок в листе итог: {len(df.columns)}")
        return pd.DataFrame()
    df['Группа'] = df['Группа'].astype(str).str.strip()
    return df

@st.cache_data
def get_questions_zash():
    try:
        df = pd.read_excel(excel_students, sheet_name='защ')
    except:
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    if 'ФИО' in df.columns and 'Вопросы' in df.columns:
        return df[['ФИО', 'Вопросы']]
    return pd.DataFrame()

@st.cache_data
def get_questions_pred():
    try:
        df = pd.read_excel(excel_students, sheet_name='пред')
    except:
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    if 'ФИО' in df.columns and 'Вопросы' in df.columns:
        return df[['ФИО', 'Вопросы']]
    return pd.DataFrame()