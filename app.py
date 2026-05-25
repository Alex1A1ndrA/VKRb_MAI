import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime, timedelta

# Это должна быть первая команда Streamlit
st.set_page_config(page_title="Генератор протоколов ВКР", layout="wide")

# Импорты из модулей
from constants import (
    output_dir_defense, output_dir_pred, output_dir_protocols,
    BASE_PATH_DATA, excel_students, excel_programs
)
from utils import parse_defense_dates, parse_predefense_dates, split_students_by_dates, extract_group
from data_loader import (
    get_students, get_dates_defense, get_dates_predefense, get_program_info,
    get_group_full_info, get_gek_members, get_questions_zash, get_questions_pred, normalize_profile_text
)
from generators import process_defenses, process_pre_defenses, commission_protocol
from parsers import (
    save_uploaded_pdfs, parse_programs_from_pdfs, parse_students_from_pdfs,
    analyze_students_fio, save_students_and_programs
)

# Создание директорий
os.makedirs(output_dir_defense, exist_ok=True)
os.makedirs(output_dir_pred, exist_ok=True)
os.makedirs(output_dir_protocols, exist_ok=True)
os.makedirs(BASE_PATH_DATA, exist_ok=True)

# Кэш для данных
@st.cache_data
def load_students_cached():
    return get_students()

@st.cache_data
def load_programs_cached():
    from data_loader import get_programs_data
    return get_programs_data()

# Загружаем данные
students_df = load_students_cached()
if students_df.empty:
    st.sidebar.error("Нет данных о студентах. Загрузите приказы или проверьте файл Студенты.xlsx")
else:
    all_groups = sorted(students_df['Группа'].unique())

# Боковое меню
doc_type = st.sidebar.radio(
    "Выберите раздел",
    ["Защита", "Предзащита", "Протокол комиссии", "Загрузка приказов", "Анализ ФИО"]
)

# ------------------------------------------------------------
# 1. ЗАЩИТА
# ------------------------------------------------------------
if doc_type == "Защита":
    if students_df.empty:
        st.stop()
    if 'last_doc_type' not in st.session_state:
        st.session_state.last_doc_type = doc_type
    if st.session_state.last_doc_type != doc_type:
        st.session_state.pop('date_groups_df_defense', None)
        st.session_state.last_doc_type = doc_type
        st.header("Генерация протоколов защиты")
        st.subheader("Выбор дат и групп")

    df_dates = get_dates_defense()
    if df_dates.empty:
        st.error("❌ Файл 'Даты.xlsx' не содержит листа 'Даты защит' или в нём нет данных.")
        st.stop()

    # Номера групп студентов, допущенных к защите
    valid_groups_numbers = set()
    group_num_to_full = {}
    for group_full in students_df[~students_df['Присутствие_защ'].astype(str).str.strip().isin(['Академ', 'полное отсутствие'])]['Группа'].unique():
        num = extract_group(group_full)
        if num:
            valid_groups_numbers.add(num)
            group_num_to_full[num] = group_full

    if not valid_groups_numbers:
        st.warning("Нет студентов, допущенных к защите.")
        st.stop()

    # Проверка наличия дат
    groups_without_dates = []
    for group_num in valid_groups_numbers:
        if df_dates[df_dates['Группа_номер'] == group_num].empty:
            groups_without_dates.append(group_num)
    if groups_without_dates:
        st.error(f"⚠ Для следующих групп не найдены даты защит: {', '.join(groups_without_dates)}")
        st.info("Добавьте строки в лист 'Даты защит'. Протоколы для этих групп не будут сгенерированы.")

    # Проверка комиссии (сокращённо)
    groups_without_commission = []
    for group_num in valid_groups_numbers:
        full_group_name = group_num_to_full.get(group_num)
        if not full_group_name:
            continue
        date_row = df_dates[df_dates['Группа_номер'] == group_num].iloc[0] if not df_dates[df_dates['Группа_номер'] == group_num].empty else None
        if date_row is None:
            continue
        predsedatel = date_row.get('Председатель', '')
        if pd.isna(predsedatel) or str(predsedatel).strip() == '':
            groups_without_commission.append(f"{group_num} (нет председателя в Даты.xlsx)")
            continue
        info = get_program_info(full_group_name, None)
        if info:
            code = info['code']
            profile = info['profile']
            norm_profile = normalize_profile_text(profile)
            gek_members = get_gek_members()
            key = (code, norm_profile)
            if key not in gek_members:
                groups_without_commission.append(f"{group_num} (нет состава ГЭК для {code} / '{profile}')")
        else:
            groups_without_commission.append(f"{group_num} (группа {full_group_name} не найдена в Программы.xlsx)")
    if groups_without_commission:
        st.error(f"⚠ Для следующих групп не хватает данных о комиссии: {', '.join(groups_without_commission)}")
        st.info("Проверьте 'Даты защит' (Председатель), 'Состав ГЭК.docx' и 'Программы.xlsx'.")

    # Формируем список дат с полными именами групп
    date_group_pairs = []
    for _, row in df_dates.iterrows():
        group_num = row['Группа_номер']
        if group_num not in valid_groups_numbers:
            continue
        full_group_name = group_num_to_full.get(group_num)
        if not full_group_name:
            full_group_name, _ = get_group_full_info(group_num)
            if not full_group_name or full_group_name == group_num:
                full_group_name = f"М8О-{group_num}"
        dates_str = str(row['Даты_основные'])
        dates = re.findall(r'\d{2}\.\d{2}\.\d{2}', dates_str)
        for d in dates:
            date_group_pairs.append((d, group_num, full_group_name))

    if not date_group_pairs:
        st.error("Нет доступных дат для выбранных групп.")
        st.stop()

    def sort_key(pair):
        try:
            return datetime.strptime(pair[0], '%d.%m.%y')
        except:
            return pair[0]
    date_group_pairs.sort(key=sort_key)

    df_pairs = pd.DataFrame(date_group_pairs, columns=["Дата", "Группа_номер", "Группа"])
    df_pairs = df_pairs.drop_duplicates().reset_index(drop=True)
    df_pairs["Выбрать"] = True

    if 'date_groups_df_defense' not in st.session_state:
        st.session_state.date_groups_df_defense = df_pairs.copy()

    edited_df = st.data_editor(
        st.session_state.date_groups_df_defense,
        column_config={
            "Выбрать": st.column_config.CheckboxColumn("Выбрать"),
            "Дата": st.column_config.TextColumn("Дата", disabled=True),
            "Группа": st.column_config.TextColumn("Группа", disabled=True),
            "Группа_номер": None,
        },
        hide_index=True,
        use_container_width=True,
    )
    st.session_state.date_groups_df_defense = edited_df

    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("Выбрать все", key=f"select_all_{doc_type}"):
            st.session_state.date_groups_df_defense["Выбрать"] = True
            st.rerun()
    with col2:
        if st.button("Снять все", key=f"clear_all_{doc_type}"):
            st.session_state.date_groups_df_defense["Выбрать"] = False
            st.rerun()

    selected_rows = st.session_state.date_groups_df_defense[st.session_state.date_groups_df_defense["Выбрать"]]
    if selected_rows.empty:
        st.warning("Не выбрано ни одной записи.")
    else:
        st.info(f"Выбрано записей: {len(selected_rows)}")

    if st.button("🚀 Сгенерировать", type="primary"):
        selected_pairs = list(zip(selected_rows["Дата"], selected_rows["Группа_номер"]))
        if not selected_pairs:
            st.warning("Не выбрано ни одной даты.")
        else:
            with st.spinner("Генерация..."):
                count = process_defenses(selected_pairs)
                out_dir = output_dir_defense
            if count > 0:
                st.success(f"✅ Сгенерировано протоколов: {count}")
                st.info(f"📁 Файлы сохранены в: {out_dir}")
            else:
                st.warning("Не удалось сгенерировать ни одного протокола.")

# ------------------------------------------------------------
# 2. ПРЕДЗАЩИТА
# ------------------------------------------------------------
elif doc_type == "Предзащита":
    if students_df.empty:
        st.stop()
    if 'last_doc_type_pred' not in st.session_state:
        st.session_state.last_doc_type_pred = doc_type
    if st.session_state.last_doc_type_pred != doc_type:
        st.session_state.pop('date_groups_df_pred', None)
        st.session_state.last_doc_type_pred = doc_type
        st.header("Генерация протоколов предзащиты")
        st.subheader("Выбор дат и групп")

    # Загружаем даты предзащит
    df_dates_pred = get_dates_predefense()
    if df_dates_pred.empty:
        st.error("❌ Файл 'Даты.xlsx' не содержит листа 'Даты предзащит' или в нём нет данных.")
        st.info("Убедитесь, что лист называется 'Даты предзащит' и в нём есть колонки 'Группа' и 'Дата предзащиты'.")
        st.stop()

    # Группы студентов, допущенных к предзащите
    valid_groups = set()
    group_num_to_full_pred = {}
    for group_full in students_df[~students_df['Присутствие_пред'].astype(str).str.strip().isin(['Академ'])]['Группа'].unique():
        num = extract_group(group_full)
        if num:
            valid_groups.add(num)
            group_num_to_full_pred[num] = group_full

    if not valid_groups:
        st.warning("Нет студентов, допущенных к предзащите.")
        st.stop()

    # Проверка наличия дат
    groups_without_dates = []
    for group_num in valid_groups:
        if df_dates_pred[df_dates_pred['Группа_номер'] == group_num].empty:
            groups_without_dates.append(group_num)
    if groups_without_dates:
        st.error(f"⚠ Для следующих групп не найдены даты предзащит: {', '.join(groups_without_dates)}")
        st.info("Добавьте строки в лист 'Даты предзащит'. Протоколы для этих групп не будут сгенерированы.")

    # Формируем список дат с полными именами групп
    date_group_pairs = []
    for _, row in df_dates_pred.iterrows():
        group_num = row['Группа_номер']
        if group_num not in valid_groups:
            continue
        dates, _ = parse_predefense_dates(row['Дата_предзащиты'])
        full_group_name = group_num_to_full_pred.get(group_num)
        if not full_group_name:
            sample = students_df[students_df['Группа'].str.contains(group_num, na=False)]
            year_int = None
            if not sample.empty:
                year_match = re.search(r'-(\d{2})$', sample.iloc[0]['Группа'])
                if year_match:
                    year_int = 2000 + int(year_match.group(1))
            full_group_name, _ = get_group_full_info(group_num, year=year_int)
            if not full_group_name or full_group_name == group_num:
                full_group_name = f"М8О-{group_num}"
        for d in dates:
            date_group_pairs.append((d, group_num, full_group_name))

    if not date_group_pairs:
        st.error("Нет доступных дат для выбранных групп.")
        st.stop()

    def sort_key(pair):
        try:
            return datetime.strptime(pair[0], '%d.%m.%y')
        except:
            return pair[0]
    date_group_pairs.sort(key=sort_key)

    df_pairs = pd.DataFrame(date_group_pairs, columns=["Дата", "Группа_номер", "Группа"])
    df_pairs = df_pairs.drop_duplicates().reset_index(drop=True)
    df_pairs["Выбрать"] = True

    if 'date_groups_df_pred' not in st.session_state:
        st.session_state.date_groups_df_pred = df_pairs.copy()

    edited_df = st.data_editor(
        st.session_state.date_groups_df_pred,
        column_config={
            "Выбрать": st.column_config.CheckboxColumn("Выбрать"),
            "Дата": st.column_config.TextColumn("Дата", disabled=True),
            "Группа": st.column_config.TextColumn("Группа", disabled=True),
            "Группа_номер": None,
        },
        hide_index=True,
        use_container_width=True,
    )
    st.session_state.date_groups_df_pred = edited_df

    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("Выбрать все", key=f"select_all_{doc_type}"):
            st.session_state.date_groups_df_pred["Выбрать"] = True
            st.rerun()
    with col2:
        if st.button("Снять все", key=f"clear_all_{doc_type}"):
            st.session_state.date_groups_df_pred["Выбрать"] = False
            st.rerun()

    selected_rows = st.session_state.date_groups_df_pred[st.session_state.date_groups_df_pred["Выбрать"]]
    if selected_rows.empty:
        st.warning("Не выбрано ни одной записи.")
    else:
        st.info(f"Выбрано записей: {len(selected_rows)}")

    if st.button("🚀 Сгенерировать", type="primary"):
        selected_pairs = list(zip(selected_rows["Дата"], selected_rows["Группа_номер"]))
        if not selected_pairs:
            st.warning("Не выбрано ни одной даты.")
        else:
            with st.spinner("Генерация..."):
                count = process_pre_defenses(selected_pairs)
                out_dir = output_dir_pred
            if count > 0:
                st.success(f"✅ Сгенерировано протоколов: {count}")
                st.info(f"📁 Файлы сохранены в: {out_dir}")
            else:
                st.warning("Не удалось сгенерировать ни одного протокола.")

# ------------------------------------------------------------
# 3. ПРОТОКОЛ КОМИССИИ
# ------------------------------------------------------------
elif doc_type == "Протокол комиссии":
    if students_df.empty:
        st.stop()
    if 'last_doc_type_comm' not in st.session_state:
        st.session_state.last_doc_type_comm = doc_type
    if st.session_state.last_doc_type_comm != doc_type:
        st.session_state.pop('commission_df', None)
        st.session_state.last_doc_type_comm = doc_type
    st.header("Генерация итоговых протоколов комиссии")
    sub_choice = st.radio("Выберите тип", ["Защита", "Предзащита"], horizontal=True)
    if 'last_sub_choice' not in st.session_state:
        st.session_state.last_sub_choice = sub_choice
    if st.session_state.last_sub_choice != sub_choice:
        st.session_state.pop('commission_df', None)
        st.session_state.last_sub_choice = sub_choice

    if sub_choice == "Защита":
        # Группы с защитой
        all_student_groups = set()
        group_num_to_full = {}
        for group_full in students_df[~students_df['Присутствие_защ'].astype(str).str.strip().isin(['Академ', 'полное отсутствие'])]['Группа'].unique():
            num = extract_group(group_full)
            if num:
                all_student_groups.add(num)
                group_num_to_full[num] = group_full
        df_dates = get_dates_defense()
        if df_dates.empty:
            st.error("❌ Лист 'Даты защит' пуст или отсутствует.")
            st.stop()
        groups_with_dates = set(df_dates['Группа_номер'].dropna().unique())
        missing_groups = all_student_groups - groups_with_dates
        if missing_groups:
            st.error(f"⚠ Для следующих групп нет дат защит: {', '.join(sorted(missing_groups))}")
        pairs = []
        for _, row in df_dates.iterrows():
            group_num = row['Группа_номер']
            if group_num not in all_student_groups or group_num not in groups_with_dates:
                continue
            full_group_name = group_num_to_full.get(group_num)
            if not full_group_name:
                full_group_name, _ = get_group_full_info(group_num)
                if not full_group_name:
                    full_group_name = f"М8О-{group_num}"
            dates = re.findall(r'\d{2}\.\d{2}\.\d{2}', str(row['Даты_основные']))
            for d in dates:
                pairs.append((d, full_group_name, group_num))
        pairs.sort(key=lambda x: datetime.strptime(x[0], '%d.%m.%y') if re.match(r'\d{2}\.\d{2}\.\d{2}', x[0]) else x[0])
    else:  # Предзащита
        all_student_groups = set()
        group_num_to_full = {}
        for group_full in students_df[~students_df['Присутствие_пред'].astype(str).str.strip().isin(['Академ'])]['Группа'].unique():
            num = extract_group(group_full)
            if num:
                all_student_groups.add(num)
                group_num_to_full[num] = group_full
        df_dates_pred = get_dates_predefense()
        if df_dates_pred.empty:
            st.error("❌ Лист 'Даты предзащит' пуст или отсутствует.")
            st.stop()
        groups_with_dates = set(df_dates_pred['Группа_номер'].dropna().unique())
        missing_groups = all_student_groups - groups_with_dates
        if missing_groups:
            st.error(f"⚠ Для следующих групп нет дат предзащит: {', '.join(sorted(missing_groups))}")
        pairs = []
        for _, row in df_dates_pred.iterrows():
            group_num = row['Группа_номер']
            if group_num not in all_student_groups or group_num not in groups_with_dates:
                continue
            full_group_name = group_num_to_full.get(group_num)
            if not full_group_name:
                sample = students_df[students_df['Группа'].str.contains(group_num, na=False)]
                year_int = None
                if not sample.empty:
                    year_match = re.search(r'-(\d{2})$', sample.iloc[0]['Группа'])
                    if year_match:
                        year_int = 2000 + int(year_match.group(1))
                full_group_name, _ = get_group_full_info(group_num, year=year_int)
                if not full_group_name:
                    full_group_name = f"М8О-{group_num}"
            dates, _ = parse_predefense_dates(row['Дата_предзащиты'])
            for d in dates:
                pairs.append((d, full_group_name, group_num))
        pairs.sort(key=lambda x: datetime.strptime(x[0], '%d.%m.%y') if re.match(r'\d{2}\.\d{2}\.\d{2}', x[0]) else x[0])

    if not pairs:
        st.error("Нет доступных дат для выбранного типа.")
        st.stop()
    else:
        st.subheader("Выбор дат и групп")
        df_pairs = pd.DataFrame(pairs, columns=["Дата", "Группа (полная)", "Группа (номер)"])
        df_pairs["Выбрать"] = True

        if 'commission_df' not in st.session_state:
            st.session_state.commission_df = df_pairs.copy()

        edited_df = st.data_editor(
            st.session_state.commission_df,
            column_config={
                "Выбрать": st.column_config.CheckboxColumn("Выбрать"),
                "Дата": st.column_config.TextColumn("Дата", disabled=True),
                "Группа (полная)": st.column_config.TextColumn("Группа", disabled=True),
                "Группа (номер)": None,
            },
            hide_index=True,
            use_container_width=True,
        )
        st.session_state.commission_df = edited_df

        col1, col2, col3 = st.columns([1, 1, 4])
        with col1:
            if st.button("Выбрать все", key="select_all_commission"):
                st.session_state.commission_df["Выбрать"] = True
                st.rerun()
        with col2:
            if st.button("Снять все", key="clear_all_commission"):
                st.session_state.commission_df["Выбрать"] = False
                st.rerun()

        selected_rows = st.session_state.commission_df[st.session_state.commission_df["Выбрать"]]
        if selected_rows.empty:
            st.warning("Не выбрано ни одной записи.")
        else:
            st.info(f"Выбрано записей: {len(selected_rows)}")

        if st.button("🚀 Сгенерировать протоколы", type="primary"):
            with st.spinner("Генерация..."):
                total_count = 0
                for _, row in selected_rows.iterrows():
                    date = row["Дата"]
                    group_num = row["Группа (номер)"]
                    count = commission_protocol(sub_choice, date, group_num)
                    total_count += count
            if total_count > 0:
                st.success(f"✅ Сгенерировано протоколов: {total_count}")
                st.info(f"📁 Файлы сохранены в: {output_dir_protocols}")
            else:
                st.warning("Не удалось сгенерировать ни одного протокола.")

# ------------------------------------------------------------
# 4. ЗАГРУЗКА ПРИКАЗОВ
# ------------------------------------------------------------
elif doc_type == "Загрузка приказов":
    st.header("📄 Загрузка и парсинг приказов")
    st.markdown("""
    Загрузите PDF-файлы приказов (проекты приказов о допуске к ВКР).  
    После обработки будут созданы/обновлены файлы:
    - `Data/Студенты.xlsx` (листы: итог, защ, пред)
    - `Data/Программы.xlsx` (лист: Лист1)
    
    Эти файлы используются для генерации протоколов.
    """)

    uploaded_files = st.file_uploader(
        "Выберите PDF-файлы приказов",
        type=["pdf"],
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("📥 Обработать и сохранить в Data"):
            with st.spinner("Сохранение PDF-файлов..."):
                saved_paths = save_uploaded_pdfs(uploaded_files)
            st.success(f"Сохранено {len(saved_paths)} файлов в {BASE_PATH_DATA}")

            with st.spinner("Парсинг образовательных программ..."):
                df_programs = parse_programs_from_pdfs(saved_paths)
                if not df_programs.empty:
                    df_programs = df_programs.rename(columns={'Профиль/Специальность': 'Профиль'})
                    st.success(f"✅ Программы извлечены: {len(df_programs)} записей")
                else:
                    st.warning("Не удалось извлечь программы")
                    df_programs = pd.DataFrame()

            with st.spinner("Парсинг студентов..."):
                df_students = parse_students_from_pdfs(saved_paths)
                if not df_students.empty:
                    st.success(f"✅ Студенты извлечены: {len(df_students)} записей")
                else:
                    st.warning("Не удалось извлечь студентов")
                    df_students = pd.DataFrame()

            if not df_students.empty or not df_programs.empty:
                if not df_students.empty:
                    save_students_and_programs(df_students, df_programs)
                    st.success("✅ Файлы Студенты.xlsx и Программы.xlsx сохранены в нужном формате")
                else:
                    if not df_programs.empty:
                        df_programs.to_excel(excel_programs, sheet_name='Лист1', index=False)
                        st.warning("Студенты не найдены, сохранены только программы")
            else:
                st.error("Не удалось извлечь ни студентов, ни программ.")

            # Сбрасываем кэшированные данные и переменные состояния
            for key in ['date_groups_df_defense', 'date_groups_df_pred', 'commission_df']:
                if key in st.session_state:
                    del st.session_state[key]
            st.cache_data.clear()
            st.rerun()
    else:
        st.info("Загрузите один или несколько PDF-файлов приказов.")

    st.subheader("Текущие файлы в папке Data")
    if os.path.exists(BASE_PATH_DATA):
        files = os.listdir(BASE_PATH_DATA)
        for f in files:
            st.write(f"📄 {f}")
    else:
        st.write("Папка Data не существует")

# ------------------------------------------------------------
# 5. АНАЛИЗ ФИО
# ------------------------------------------------------------
elif doc_type == "Анализ ФИО":
    st.header("🔍 Анализ ФИО студентов и связанных протоколов")
    
    if students_df.empty:
        st.error("Нет данных о студентах.")
        st.stop()

    output_dirs = {
        "Защита": output_dir_defense,
        "Предзащита": output_dir_pred,
        "Протокол комиссии": output_dir_protocols
    }

    problematic_students, summary = analyze_students_fio(students_df, output_dirs)

    st.subheader("📊 Статистика")
    st.write(summary)

    if not problematic_students:
        st.success("✅ Все ФИО имеют корректный формат и пол определяется успешно.")
    else:
        st.warning(f"⚠️ Найдено {len(problematic_students)} проблемных записей")
        for p in problematic_students:
            with st.expander(f"❌ {p['fio']} (Группа: {p['group'] or '?'}) – {p['reason']}"):
                if p['files']:
                    st.markdown("**Найденные протоколы:**")
                    for file_info in p['files']:
                        with open(file_info['path'], "rb") as f:
                            st.download_button(
                                label=f"📎 {file_info['name']} ({file_info['type']})",
                                data=f,
                                file_name=file_info['name'],
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                else:
                    st.info("Протоколов для этого студента не найдено.")

    with st.expander("📋 Полный список студентов"):
        st.dataframe(students_df[['ФИО', 'Группа']], use_container_width=True)