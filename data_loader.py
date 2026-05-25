import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from docx import Document

from constants import (
    excel_programs, excel_dates, excel_students, gek_docx,
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

def normalize_profile_text(profile: str) -> str:
    if not isinstance(profile, str):
        profile = str(profile)
    profile = re.sub(r'[«»"*]', '', profile)
    profile = re.sub(r'\s+', ' ', profile).strip()
    return profile.lower()

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

def _get_suffix(level: str) -> str:
    if level == 'БВО':
        return 'БВ'
    elif level == 'СВО':
        return 'СВ'
    elif level == 'Магистратура':
        return 'М'
    elif level == 'Бакалавриат':
        return 'Б'
    else:
        return ''
    
def get_group_full_info(group_num, year=None):
    programs_data = get_programs_data()
    group_num_str = str(group_num)

    for full_name, records in programs_data.items():
        match = re.search(r'М8О-(\d{3})', full_name)
        if match and match.group(1) == group_num_str:
            if year is not None:
                for rec in records:
                    if rec['year'] == year:
                        level = rec['level']
                        year_used = rec['year']
                        year_str = str(year_used)[-2:]
                        suffix = _get_suffix(level)
                        full_group_name = f"М8О-{group_num}{suffix}-{year_str}"
                        graduation_year = year_used + int(group_num_str[0])
                        return full_group_name, str(graduation_year)
            if records:
                rec = records[0]
                level = rec['level']
                year_used = rec['year']
                year_str = str(year_used)[-2:]
                suffix = _get_suffix(level)
                full_group_name = f"М8О-{group_num}{suffix}-{year_str}"
                graduation_year = year_used + int(group_num_str[0])
                return full_group_name, str(graduation_year)

    return str(group_num), ''

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
        while idx < len(lines) and 'экзаменационная комиссия' not in lines[idx].lower():
            if ('профиль:' in lines[idx].lower() or 
                'магистерская программа' in lines[idx].lower() or 
                'программа спецво:' in lines[idx].lower() or
                ('направление' in lines[idx].lower() and re.search(r'\d{2}\.\d{2}\.\d{2}', lines[idx]))):
                return [], None, idx
            idx += 1
        if idx >= len(lines):
            return [], None, idx
        idx += 1
        members = []
        while idx < len(lines):
            cur_line = lines[idx]
            if ('профиль:' in cur_line.lower() or 
                'магистерская программа' in cur_line.lower() or 
                'программа спецво:' in cur_line.lower() or
                ('направление' in cur_line.lower() and re.search(r'\d{2}\.\d{2}\.\d{2}', cur_line))):
                break
            if 'секретарь:' in cur_line.lower():
                idx += 1
                break
            if ('–' in cur_line or '-' in cur_line) and not any(x in cur_line.lower() for x in 
                ['экзаменационная комиссия', 'секретарь:', 'магистерская программа:', 'профиль:', 'программа спецво:']):
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
        code_match = re.search(r'(\d{2}\.\d{2}\.\d{2})', line)
        if code_match:
            current_code = code_match.group(1)
            j = i
            while j < len(lines):
                next_code_match = re.search(r'(\d{2}\.\d{2}\.\d{2})', lines[j]) if j > i else None
                if next_code_match and next_code_match.group(1) != current_code:
                    break
                prof_match = re.search(r'(?:профиль:|магистерская программа:?|программа спецво:?)\s*«?([^»]+)»?', lines[j], re.IGNORECASE)
                if prof_match:
                    profile = prof_match.group(1).strip()
                    start_commission = j
                    found = False
                    for k in range(j, min(j+20, len(lines))):
                        if 'экзаменационная комиссия' in lines[k].lower():
                            start_commission = k
                            found = True
                            break
                    if found:
                        members, _, next_idx = extract_commission(lines, start_commission)
                        if members:
                            norm_profile = normalize_profile(profile)
                            key = (current_code, norm_profile)
                            gek_members[key] = members
                            j = next_idx if next_idx > j else j + 1
                            continue
                    j += 1
                else:
                    j += 1
            i = j if j > i else i + 1
        else:
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

    group_col = None
    date_col = None
    chair_col = None
    time_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if 'групп' in col_lower:
            group_col = col
        if 'дата' in col_lower and ('защит' in col_lower or 'защиты' in col_lower):
            date_col = col
        if 'председатель' in col_lower:
            chair_col = col
        if 'время' in col_lower:
            time_col = col

    if group_col is None or date_col is None:
        st.error(f"Не найдены колонки с группой или датой в листе 'Даты защит'. Найдены: {list(df.columns)}")
        return pd.DataFrame()

    df = df.rename(columns={group_col: 'Группа_полная', date_col: 'Даты_основные'})
    if chair_col:
        df = df.rename(columns={chair_col: 'Председатель'})
    if time_col:
        df = df.rename(columns={time_col: 'Время'})
    else:
        df['Время'] = ''

    df['Группа_полная'] = df['Группа_полная'].astype(str).str.strip()
    df['Группа_номер'] = df['Группа_полная'].apply(lambda x: extract_group(x) if pd.notna(x) and x != 'nan' else None)
    df['Год'] = df['Группа_полная'].apply(lambda x: x.split('-')[-1] if '-' in x else None)
    df['Время'] = df['Время'].astype(str).str.strip()
    df['Код_направления'] = None

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
            st.error(f"Неожиданное число колонок в Датах предзащит: {len(df.columns)}. Ожидается 5 колонок.")
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