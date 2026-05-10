import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# Это должна быть первая команда Streamlit
st.set_page_config(page_title="Генератор протоколов ВКР", layout="wide")

# Импорты из модулей
from constants import (
    output_dir_defense, output_dir_pred, output_dir_protocols
)
from utils import parse_defense_dates, parse_predefense_dates, split_students_by_dates
from data_loader import (
    get_students, get_dates_defense, get_dates_predefense, get_program_info,
    get_group_full_info, get_commission_data, get_commission_data_norm,
    get_gek_members, get_questions_zash, get_questions_pred
)
from generators import process_defenses, process_pre_defenses, commission_protocol

# Создание директорий
os.makedirs(output_dir_defense, exist_ok=True)
os.makedirs(output_dir_pred, exist_ok=True)
os.makedirs(output_dir_protocols, exist_ok=True)

# Загружаем данные для интерфейса (кэшированные)
students_df = get_students()
if students_df.empty:
    st.stop()

all_groups = sorted(students_df['Группа'].unique())

# Боковое меню выбора типа документа
doc_type = st.sidebar.radio("Выберите тип документа", ["Защита", "Предзащита", "Протокол комиссии"])

if doc_type in ["Защита", "Предзащита"]:
    # Отслеживаем смену типа документа
    if 'last_doc_type' not in st.session_state:
        st.session_state.last_doc_type = doc_type
    if st.session_state.last_doc_type != doc_type:
        # Сброс данных при смене типа
        st.session_state.pop('date_groups_df', None)
        st.session_state.last_doc_type = doc_type
        st.header(f"Генерация протоколов {doc_type.lower()}")
        st.subheader("Выбор дат и групп")

    # Формируем список всех доступных дат с полными названиями групп
    date_group_pairs = []
    if doc_type == "Защита":
        df_dates = get_dates_defense()
        # Получаем группы с допущенными к защите студентами
        valid_groups = set(
            students_df[~students_df['Присутствие_защ'].astype(str).str.strip().isin(['Академ', 'полное отсутствие'])]['Группа'].unique()
        )
        if not df_dates.empty:
            for _, row in df_dates.iterrows():
                group_num = row['Группа_номер']
                if group_num not in valid_groups:
                    continue
                dates_str = str(row['Даты_основные']).strip()
                main_part = dates_str.split(',')[0].strip()
                dates = [d.strip() for d in main_part.replace('\\', '/').split('/') if d.strip()]
                full_group_name = row['Группа_полная']
                for d in dates:
                    date_group_pairs.append((d, group_num, full_group_name))
    else:  # Предзащита
        df_dates_pred = get_dates_predefense()
        # Группы с допущенными к предзащите студентами
        valid_groups = set(
            students_df[~students_df['Присутствие_пред'].astype(str).str.strip().isin(['Академ'])]['Группа'].unique()
        )
        if not df_dates_pred.empty:
            for _, row in df_dates_pred.iterrows():
                group_num = row['Группа_номер']
                if group_num not in valid_groups:
                    continue
                dates, _ = parse_predefense_dates(row['Дата_предзащиты'])
                full_group_name, _ = get_group_full_info(group_num)
                if not full_group_name or full_group_name == group_num:
                    full_group_name = f"М8О-{group_num}"
                for d in dates:
                    date_group_pairs.append((d, group_num, full_group_name))

    if not date_group_pairs:
        st.error("Нет доступных дат для выбранного типа.")
    else:
        # Сортируем по дате
        def sort_key(pair):
            try:
                return datetime.strptime(pair[0], '%d.%m.%y')
            except:
                return pair[0]
        date_group_pairs.sort(key=sort_key)

        # Создаем DataFrame для отображения
        df_pairs = pd.DataFrame(date_group_pairs, columns=["Дата", "Группа_номер", "Группа"])
        df_pairs = df_pairs.drop_duplicates().reset_index(drop=True)
        df_pairs["Выбрать"] = True

        # Инициализация session state
        if 'date_groups_df' not in st.session_state:
            st.session_state.date_groups_df = df_pairs.copy()

        # Редактируемая таблица
        edited_df = st.data_editor(
            st.session_state.date_groups_df,
            column_config={
                "Выбрать": st.column_config.CheckboxColumn("Выбрать"),
                "Дата": st.column_config.TextColumn("Дата", disabled=True),
                "Группа": st.column_config.TextColumn("Группа", disabled=True),
                "Группа_номер": None,
            },
            hide_index=True,
            use_container_width=True,
            key="date_groups_editor"
        )
        st.session_state.date_groups_df = edited_df

        col1, col2, col3 = st.columns([1, 1, 4])
        with col1:
            if st.button("Выбрать все", key=f"select_all_{doc_type}"):
                st.session_state.date_groups_df["Выбрать"] = True
                st.rerun()
        with col2:
            if st.button("Снять все", key=f"clear_all_{doc_type}"):
                st.session_state.date_groups_df["Выбрать"] = False
                st.rerun()

        selected_rows = st.session_state.date_groups_df[st.session_state.date_groups_df["Выбрать"]]
        if selected_rows.empty:
            st.warning("Не выбрано ни одной записи.")
        else:
            st.info(f"Выбрано записей: {len(selected_rows)}")

        if st.button("🚀 Сгенерировать", type="primary"):
            selected_pairs = list(zip(selected_rows["Дата"], selected_rows["Группа_номер"]))
            if not selected_pairs:
                st.warning("Не выбрано ни одной даты для генерации.")
            else:
                with st.spinner("Генерация..."):
                    if doc_type == "Защита":
                        count = process_defenses(selected_pairs)
                        out_dir = output_dir_defense
                    else:
                        count = process_pre_defenses(selected_pairs)
                        out_dir = output_dir_pred

            if count > 0:
                st.success(f"✅ Сгенерировано протоколов: {count}")
                st.info(f"📁 Файлы сохранены в: {out_dir}")
            else:
                st.warning("Не удалось сгенерировать ни одного протокола.")

elif doc_type == "Протокол комиссии":
    # Отслеживаем смену типа документа
    if 'last_doc_type_comm' not in st.session_state:
        st.session_state.last_doc_type_comm = doc_type
    if st.session_state.last_doc_type_comm != doc_type:
        # Сброс данных при смене типа
        st.session_state.pop('commission_df', None)
        st.session_state.last_doc_type_comm = doc_type
    st.header("Генерация итоговых протоколов комиссии")
    sub_choice = st.radio("Выберите тип", ["Защита", "Предзащита"], horizontal=True)
    # также отслеживаем подтип (Защита/Предзащита)
    if 'last_sub_choice' not in st.session_state:
        st.session_state.last_sub_choice = sub_choice
    if st.session_state.last_sub_choice != sub_choice:
        st.session_state.pop('commission_df', None)
        st.session_state.last_sub_choice = sub_choice

    # Формируем список доступных дат с группами
    pairs = []
    if sub_choice == "Защита":
        df_dates = get_dates_defense()
        if not df_dates.empty:
            for _, row in df_dates.iterrows():
                dates_str = str(row['Даты_основные']).strip()
                main_part = dates_str.split(',')[0].strip()
                dates = [d.strip() for d in main_part.replace('\\', '/').split('/') if d.strip()]
                for d in dates:
                    pairs.append((d, row['Группа_полная'], row['Группа_номер']))
        def date_key(pair):
            try:
                return datetime.strptime(pair[0], '%d.%m.%y')
            except:
                return pair[0]
        pairs.sort(key=date_key)
    else:
        df_dates_pred = get_dates_predefense()
        if not df_dates_pred.empty:
            for _, row in df_dates_pred.iterrows():
                dates, _ = parse_predefense_dates(row['Дата_предзащиты'])
                for d in dates:
                    full_group_name, _ = get_group_full_info(row['Группа_номер'])
                    if not full_group_name or full_group_name == row['Группа_номер']:
                        full_group_name = f"М8О-{row['Группа_номер']}"
                    pairs.append((d, full_group_name, row['Группа_номер']))
        def date_key(pair):
            try:
                return datetime.strptime(pair[0], '%d.%m.%y')
            except:
                return pair[0]
        pairs.sort(key=date_key)

    if not pairs:
        st.error("Нет доступных дат для выбранного типа.")
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
            key="commission_editor"
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