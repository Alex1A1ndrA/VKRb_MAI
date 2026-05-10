import os
import re
import random
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
from docx import Document

from constants import (
    GRADE_MAP, SECRETARY_DEFAULT, DASH, MONTH_NAMES,
    template_path_defense_5, template_path_defense_4,
    template_path_pred_bachelor, template_path_pred_master,
    template_path_commission,
    output_dir_defense, output_dir_pred, output_dir_protocols
)
from utils import (
    value_or_dash, short_fio, fio_in_case, role_fio_case, clean_text,
    parse_questions, split_students_by_dates, parse_defense_dates,
    parse_predefense_dates, remove_empty_paragraphs_around_tables,
    insert_empty_paragraphs_after_text, set_landscape, ensure_section,
    detect_gender_by_patronymic
)
from data_loader import (
    get_students, get_dates_defense, get_dates_predefense, get_questions_zash,
    get_questions_pred, get_gek_members, get_program_info, get_group_full_info,
    get_program_description, qualification_genitive, get_commission_data,
    get_commission_data_norm
)

def process_defenses(selected_pairs=None):
    df_students = get_students()
    if df_students.empty:
        st.error("Нет данных о студентах.")
        return 0
    df_dates = get_dates_defense()
    if df_dates.empty:
        st.error("Нет данных о датах защит.")
        return 0
    df_questions = get_questions_zash()

    df_students = df_students[~df_students['Присутствие_защ'].astype(str).str.strip().isin(['Академ', 'полное отсутствие'])]
    if df_students.empty:
        st.warning("Нет студентов для защиты.")
        return 0

    saved_count = 0
    protocol_number = 1
    daily_slots = {}

    selected_set = set(selected_pairs) if selected_pairs is not None else None

    for group_num, group_students in df_students.groupby('Группа'):
        date_row = df_dates[df_dates['Группа_номер'] == group_num]
        if date_row.empty:
            st.warning(f"⚠ Для группы {group_num} не найдено расписание защит.")
            continue
        date_row = date_row.iloc[0]
        full_group_name = date_row['Группа_полная']
        group_year = date_row['Год']
        try:
            year_int = 2000 + int(group_year) if group_year else None
        except:
            year_int = None

        dates_list = parse_defense_dates(date_row['Даты_основные'])
        if not dates_list:
            st.warning(f"⚠ Для группы {group_num} нет дат защит.")
            continue
        start_time_str = date_row['Время']
        try:
            start_h, start_m = map(int, start_time_str.split(':')[:2])
        except:
            start_h, start_m = 10, 0

        info = get_program_info(group_num, year_int)
        if info:
            code = info['code']
            profile = info['profile']
        else:
            code = date_row['Код_направления']
            profile = None
            st.warning(f"⚠ Не удалось определить программу для группы {group_num}")

        commission_members = []
        gek_members = get_gek_members()
        if profile and gek_members:
            norm_profile = re.sub(r'[«»"*]', '', profile).lower().strip()
            key = (code, norm_profile)
            if key in gek_members:
                commission_members = gek_members[key]
            else:
                st.warning(f"⚠ Для группы {group_num} не найден состав комиссии (код={code}, профиль={profile})")
                commission_members = [date_row['Председатель']] if pd.notna(date_row['Председатель']) else []
        else:
            commission_members = [date_row['Председатель']] if pd.notna(date_row['Председатель']) else []

        if not commission_members:
            continue

        predsedatel = commission_members[0]
        other_members = commission_members[1:]

        has_fifth = len(other_members) >= 5
        template_path = template_path_defense_5 if has_fifth else template_path_defense_4

        context_base = {
            'Председатель': value_or_dash(short_fio(predsedatel)),
        }
        for i in range(1, 6):
            if i <= len(other_members):
                context_base[f'Член_комиссии_{i}'] = value_or_dash(short_fio(other_members[i-1]))
            else:
                context_base[f'Член_комиссии_{i}'] = DASH

        indices_groups = split_students_by_dates(group_students, dates_list)

        for date_idx, indices in enumerate(indices_groups):
            if not indices:
                continue
            date_str = dates_list[date_idx]

            if selected_set is not None and (date_str, group_num) not in selected_set:
                continue

            try:
                base_day = datetime.strptime(date_str, '%d.%m.%y')
            except:
                st.error(f"Ошибка парсинга даты {date_str}")
                continue

            for pos_in_day, idx in enumerate(indices):
                student = group_students.iloc[idx]
                fio = student['ФИО']
                if pd.isna(fio):
                    continue

                tema = student['Тема']
                ruk = student['Руководитель']
                new_tema = student.get('Новая_тема')
                new_ruk = student.get('Новый_рук')
                if pd.notna(new_tema) and new_tema.strip() and new_tema.strip() != str(tema).strip():
                    tema = new_tema
                if pd.notna(new_ruk) and new_ruk.strip() and new_ruk.strip() != str(ruk).strip():
                    ruk = new_ruk

                konsultant = student.get('Консультант', '')
                recenzent = student.get('Рецензент', '')
                pages = student.get('Листов', '')
                if pd.notna(pages) and pages != '':
                    try:
                        pages = str(int(float(pages)))
                    except:
                        pages = str(pages).strip()
                else:
                    pages = ''

                ocenka = student.get('Оценка_защ', '')
                ocenka_str = str(ocenka).strip()
                if ocenka_str.replace('.', '', 1).isdigit():
                    ocenka_str = str(int(float(ocenka_str)))
                grade_word = GRADE_MAP.get(ocenka_str, '')

                questions_raw = ''
                if not df_questions.empty:
                    qrow = df_questions[df_questions['ФИО'] == fio]
                    if not qrow.empty:
                        questions_raw = qrow.iloc[0].get('Вопросы', '')
                all_members = commission_members
                questions = parse_questions(questions_raw, all_members)

                gender_student = detect_gender_by_patronymic(fio)

                info_prog = info
                if info_prog:
                    direction = f"{info_prog['code']} {info_prog['direction']}"
                    program = info_prog['profile']
                    level = info_prog['level']
                    qualification_orig = info_prog['qualification']
                    program_description = get_program_description(level)
                    qualification_gen = qualification_genitive(qualification_orig)
                    diploma = f"{qualification_genitive(qualification_orig).lower()} без отличия"
                    qualification = qualification_orig
                else:
                    direction = ""
                    program = ""
                    program_description = ""
                    qualification_gen = ""
                    qualification = ""
                    diploma = ""

                full_group_name_prog, graduation_year = get_group_full_info(group_num, year=year_int)

                date_key = date_str
                if date_key not in daily_slots:
                    daily_slots[date_key] = 0
                slot = daily_slots[date_key]
                daily_slots[date_key] += 1

                start = base_day.replace(hour=start_h, minute=start_m) + timedelta(minutes=45 * slot)
                end = start + timedelta(minutes=45)

                context = {
                    'Номер': str(protocol_number),
                    'Дата': base_day.strftime('%d.%m.%Y'),
                    'Начало_ч': str(start.hour).zfill(2),
                    'Начало_м': str(start.minute).zfill(2),
                    'Конец_ч': str(end.hour).zfill(2),
                    'Конец_м': str(end.minute).zfill(2),

                    'Студент': fio_in_case(fio, 'nomn', gender_student),
                    'СтудентР': fio_in_case(fio, 'gent', gender_student),
                    'СтудентД': fio_in_case(fio, 'datv', gender_student),
                    'СтудентТ': fio_in_case(fio, 'ablt', gender_student),

                    'Группа': full_group_name,
                    'ГруппаПолная': full_group_name_prog,
                    'ГодВыпуска': graduation_year,

                    'Тема': clean_text(tema),
                    'Направление': value_or_dash(direction),
                    'Программа': value_or_dash(program),

                    'программаО': value_or_dash(program_description),
                    'КвалификацияР': value_or_dash(qualification_gen),
                    'квалификация': value_or_dash(qualification),
                    'диплом': value_or_dash(diploma),
                    'Секретарь': SECRETARY_DEFAULT,

                    'РуководительР': value_or_dash(role_fio_case(ruk, 'gent')),
                    'КонсультантР': value_or_dash(role_fio_case(konsultant, 'gent')),
                    'РецензентР': value_or_dash(role_fio_case(recenzent, 'gent')),

                    'Руководитель': value_or_dash(role_fio_case(ruk, 'gent')),
                    'Консультант': value_or_dash(role_fio_case(konsultant, 'gent')),
                    'Рецензент': value_or_dash(role_fio_case(recenzent, 'gent')),

                    'Листов': pages,
                    'Выступление': '15',

                    'Вопрос_1': questions[0],
                    'Вопрос_2': questions[1],
                    'Вопрос_3': questions[2],
                    'Вопрос_4': questions[3],

                    'ОценкаП': grade_word,
                }
                context.update(context_base)

                doc = DocxTemplate(template_path)
                doc.render(context)

                safe_fio = short_fio(fio).replace('.', '').replace(' ', '_')
                filename = f"{full_group_name}_{safe_fio}_{protocol_number}.docx"
                filepath = os.path.join(output_dir_defense, filename)
                doc.save(filepath)

                st.write(f"✔ Сохранён: {filename}")
                protocol_number += 1
                saved_count += 1

    return saved_count

def process_pre_defenses(selected_pairs=None):
    df_students = get_students()
    if df_students.empty:
        return 0
    df_dates_pred = get_dates_predefense()
    if df_dates_pred.empty:
        return 0
    df_dates_def = get_dates_defense()
    df_questions = get_questions_pred()

    df_students = df_students[~df_students['Присутствие_пред'].astype(str).str.strip().isin(['Академ'])]
    if df_students.empty:
        st.warning("Нет студентов для предзащиты.")
        return 0

    saved_count = 0
    protocol_number = 1

    selected_set = set(selected_pairs) if selected_pairs is not None else None

    for group_num, group_students in df_students.groupby('Группа'):
        date_rows = df_dates_pred[df_dates_pred['Группа_номер'] == group_num]
        if date_rows.empty:
            st.warning(f"⚠ Для группы {group_num} не найдено расписание предзащит.")
            continue
        date_row = date_rows.iloc[0]
        dates_list, time_str = parse_predefense_dates(date_row['Дата_предзащиты'])
        if not dates_list:
            st.warning(f"⚠ Не удалось извлечь даты предзащиты для группы {group_num}.")
            continue
        auditorium = date_row['Аудитория']
        comm_raw = date_row['Комиссия']

        comm_fams_raw = []
        if pd.notna(comm_raw):
            comm_str = str(comm_raw)
            comm_str = re.sub(r'[,;]|<br>|\n', '|', comm_str)
            comm_fams_raw = [f.strip() for f in comm_str.split('|') if f.strip()]

        comm_fams = []
        for full in comm_fams_raw:
            fam_part = full.split()[0] if full else ''
            fam_part = fam_part.rstrip('.')
            comm_fams.append(fam_part)

        commission_data = get_commission_data()
        commission_data_norm = get_commission_data_norm()
        
        commission_with_roles = []
        for fam in comm_fams:
            norm_fam = fam.lower().rstrip('.')
            if norm_fam in commission_data_norm:
                full_str, can_be_chair, is_docent = commission_data_norm[norm_fam]
                commission_with_roles.append((can_be_chair, is_docent, full_str))
            else:
                commission_with_roles.append((False, False, fam))

        commission_with_roles.sort(key=lambda x: (not x[0], not x[1]))
        chair_name = commission_with_roles[0][2] if commission_with_roles else "____________"
        members_final = [x[2] for x in commission_with_roles[1:]]
        while len(members_final) < 4:
            members_final.append("____________")

        def_row = df_dates_def[df_dates_def['Группа_номер'] == group_num]
        year_int = None
        full_group_name_from_def = None
        if not def_row.empty:
            full_def = def_row.iloc[0]['Группа_полная']
            year_part = full_def.split('-')[-1] if '-' in full_def else None
            try:
                year_int = 2000 + int(year_part) if year_part else None
            except:
                pass
            full_group_name_from_def = full_def

        full_group_name, _ = get_group_full_info(group_num, year=year_int)
        if not full_group_name or full_group_name == group_num:
            full_group_name = full_group_name_from_def if full_group_name_from_def else group_num

        defense_date_str = ''
        if not def_row.empty:
            defense_dates_raw = def_row.iloc[0]['Даты_основные']
            if pd.notna(defense_dates_raw):
                main_part = defense_dates_raw.split(',')[0].strip()
                defense_date_str = main_part.replace('\\', ',')
            else:
                defense_date_str = ''
        else:
            defense_date_str = ''

        info = get_program_info(group_num, year_int) if year_int else get_program_info(group_num)
        if info:
            level = info['level']
        else:
            level = 'Бакалавриат'
        is_master = level in ['Магистратура', 'СВО']
        template_path = template_path_pred_master if is_master else template_path_pred_bachelor

        indices_groups = split_students_by_dates(group_students, dates_list)

        for date_idx, indices in enumerate(indices_groups):
            if not indices:
                continue
            date_str = dates_list[date_idx]

            if selected_set is not None and (date_str, group_num) not in selected_set:
                continue

            try:
                dt = datetime.strptime(date_str, '%d.%m.%y')
                day_str = dt.strftime('%d')
                month_str = MONTH_NAMES[dt.strftime('%m')]
            except:
                day_str = date_str[:2]
                month_str = date_str[3:5]

            students_data = []
            missed_lines = []
            missed_index = 2

            sub_students = group_students.iloc[indices]
            for idx_in_sub, (_, student) in enumerate(sub_students.iterrows()):
                fio = student['ФИО']
                group_name = student['Группа']
                was_present = str(student.get('Присутствие_пред', '')).strip().lower() != 'нет'

                tema = student.get('Тема', '')
                ruk = student.get('Руководитель', '')
                new_tema = student.get('Новая_тема', '')
                new_ruk = student.get('Новый_рук', '')
                reviewer = student.get('Рецензент', '')

                if is_master:
                    if was_present:
                        students_data.append({
                            'num': idx_in_sub + 1,
                            'ФИО': fio,
                            'Группа': full_group_name,
                            'Рецензент': "" if pd.isnull(reviewer) else str(reviewer),
                        })
                    else:
                        missed_lines.append(f"{missed_index}. {fio} не явился на предзащиту ВКРМ.")
                        missed_index += 1
                else:
                    if was_present:
                        students_data.append(f"{len(students_data)+1}. {fio} ({full_group_name})")
                    else:
                        missed_lines.append(f"{missed_index}. {fio} не явился на предзащиту ВКРБ.")
                        missed_index += 1

                if pd.notna(new_tema) and new_tema.strip() and new_tema.strip() != str(tema).strip():
                    fio_dative = fio_in_case(fio, 'datv')
                    missed_lines.append(f"{missed_index}. {fio_dative} рекомендуется сменить тему на \"{clean_text(new_tema)}\".")
                    missed_index += 1
                if pd.notna(new_ruk) and new_ruk.strip() and new_ruk.strip() != str(ruk).strip():
                    fio_dative = fio_in_case(fio, 'datv')
                    missed_lines.append(f"{missed_index}. {fio_dative} рекомендуется сменить руководителя на {new_ruk}.")
                    missed_index += 1

            base_context = {
                'Номер': str(protocol_number),
                'Число': day_str,
                'Месяц': month_str,
                'Неявившиеся': "\n".join(missed_lines),
                'Защита': defense_date_str,
                'Председатель': chair_name,
                'Член_комиссии_1': members_final[0],
                'Член_комиссии_2': members_final[1],
                'Член_комиссии_3': members_final[2],
                'Член_комиссии_4': members_final[3],
                'Секретарь': SECRETARY_DEFAULT,
            }
            if is_master:
                context = {**base_context, 'Студенты_таблица': students_data}
            else:
                context = {**base_context, 'Студенты': "\n".join(students_data)}

            doc = DocxTemplate(template_path)
            doc.render(context)

            prefix = 'M' if is_master else 'B'
            filename = f'{prefix}_{day_str}_{month_str}_{full_group_name}_{protocol_number}.docx'
            output_path = os.path.join(output_dir_pred, filename)
            doc.save(output_path)

            if is_master:
                remove_empty_paragraphs_around_tables(output_path)
                insert_empty_paragraphs_after_text(output_path, "Защита выпускной квалификационной работы состоится", 1)

            st.write(f"✔ Протокол сохранён: {filename}")
            protocol_number += 1
            saved_count += 1

    return saved_count

def commission_protocol(protocol_type, selected_date, group_num):
    df_students = get_students()
    if df_students.empty:
        return 0

    if protocol_type == 'Защита':
        df_dates = get_dates_defense()
        if df_dates.empty:
            return 0

        date_row = df_dates[df_dates['Группа_номер'] == group_num].iloc[0]
        dates_str = str(date_row['Даты_основные']).strip()
        main_part = dates_str.split(',')[0].strip()
        dates_list = [d.strip() for d in main_part.replace('\\', '/').split('/') if d.strip()]
        
        try:
            date_idx = dates_list.index(selected_date)
        except ValueError:
            st.error("Выбранная дата не в списке основных дат группы.")
            return 0

        all_studs = df_students[(df_students['Группа'] == group_num) &
                                 (~df_students['Присутствие_защ'].astype(str).str.strip().isin(['Академ', 'полное отсутствие']))]
        if all_studs.empty:
            st.warning("Нет студентов для этой группы.")
            return 0

        indices_groups = split_students_by_dates(all_studs, dates_list)
        if date_idx >= len(indices_groups) or not indices_groups[date_idx]:
            st.warning("Нет студентов на выбранную дату.")
            return 0

        selected_indices = indices_groups[date_idx]
        studs = all_studs.iloc[selected_indices]

        table = []
        for _, s in studs.iterrows():
            fio = s['ФИО']
            tema = s['Тема']
            ruk = s['Руководитель']
            ocenka = s.get('Оценка_защ', '')
            primech = ''
            table.append({'ФИО': fio, 'Тема': tema, 'Руководитель': ruk, 'Оценка': ocenka, 'Примечание': primech})
        protocol_type_str = "ЗАЩИТЫ"
        prefix = "Защита"
        date_part = selected_date.replace('.', '_')
        filename = f"{prefix}_{group_num}_{date_part}.docx"

    else:
        df_dates_pred = get_dates_predefense()
        if df_dates_pred.empty:
            return 0

        date_rows = df_dates_pred[df_dates_pred['Группа_номер'] == group_num]
        if date_rows.empty:
            st.error("Ошибка: не найдено расписание для группы.")
            return 0
        date_row = date_rows.iloc[0]
        dates_list, _ = parse_predefense_dates(date_row['Дата_предзащиты'])
        if not dates_list:
            st.error("Ошибка: не удалось получить список дат.")
            return 0

        try:
            date_idx = dates_list.index(selected_date)
        except ValueError:
            st.error("Выбранная дата не соответствует списку дат группы.")
            return 0

        all_studs = df_students[(df_students['Группа'] == group_num) &
                                 (~df_students['Присутствие_пред'].astype(str).str.strip().isin(['Академ']))]
        if all_studs.empty:
            st.warning("Нет студентов для этой группы.")
            return 0

        indices_groups = split_students_by_dates(all_studs, dates_list)
        if date_idx >= len(indices_groups) or not indices_groups[date_idx]:
            st.warning("Нет студентов на выбранную дату.")
            return 0

        selected_indices = indices_groups[date_idx]
        sub_students = all_studs.iloc[selected_indices]

        table = []
        for _, s in sub_students.iterrows():
            fio = s['ФИО']
            tema = s['Тема']
            ruk = s['Руководитель']
            primech_parts = []
            new_tema = s.get('Новая_тема', '')
            new_ruk = s.get('Новый_рук', '')
            if pd.notna(new_tema) and new_tema.strip() and new_tema.strip() != str(s['Тема']).strip():
                primech_parts.append(f"Новая тема: {new_tema}")
            if pd.notna(new_ruk) and new_ruk.strip() and new_ruk.strip() != str(s['Руководитель']).strip():
                primech_parts.append(f"Новый руководитель: {new_ruk}")
            primech = "; ".join(primech_parts)
            table.append({'ФИО': fio, 'Тема': tema, 'Руководитель': ruk, 'Оценка': '', 'Примечание': primech})
        protocol_type_str = "ПРЕДЗАЩИТЫ"
        prefix = "Предзащита"
        date_part = selected_date.replace('.', '_')
        filename = f"{prefix}_{group_num}_{date_part}.docx"

    context = {
        'тип': protocol_type_str,
        'Студенты_таблица': table
    }

    doc = DocxTemplate(template_path_commission)
    doc.render(context)

    set_landscape(doc)
    target_width = doc.sections[0].page_width
    target_height = doc.sections[0].page_height

    temp_path = os.path.join(output_dir_protocols, '_temp_' + filename)
    doc.save(temp_path)

    remove_empty_paragraphs_around_tables(temp_path)

    doc_fixed = Document(temp_path)
    ensure_section(doc_fixed, target_width, target_height)
    set_landscape(doc_fixed)

    final_path = os.path.join(output_dir_protocols, filename)
    doc_fixed.save(final_path)

    if os.path.exists(temp_path):
        os.remove(temp_path)

    st.write(f"✔ Протокол комиссии сохранён: {filename}")
    return 1