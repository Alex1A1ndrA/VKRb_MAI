import streamlit as st
import pdfplumber
import pandas as pd
import re
import os

from constants import BASE_PATH_DATA, excel_students, excel_programs

def transform_group(grp: str, filename: str) -> str:
    if not grp:
        return "UNKNOWN"
    grp = grp.strip()
    if re.match(r'^[А-Яа-я0-9]+-[24]', grp):
        return grp
    m = re.match(r'^([А-Яа-я0-9]+)-(.+)$', grp)
    if m:
        prefix, rest = m.groups()
        suffix = "4" if "ВКРБ" in filename else "2" if "СпецВО" in filename else ""
        return f"{prefix}-{suffix}{rest}"
    return grp


def clean_text(x: str) -> str:
    return re.sub(r"\s+", " ", x).strip() if isinstance(x, str) else ""


def find_student_group(fio_raw: str, full_text: str, filename: str) -> str:
    fio_clean = clean_text(fio_raw).lower()
    text_lower = full_text.lower()
    pos = text_lower.find(fio_clean)
    if pos != -1:
        return _get_group_above(pos, full_text, filename)
    parts = fio_clean.split()
    if len(parts) >= 2:
        surname = parts[0]
        name = parts[1]
        patronymic = parts[2] if len(parts) > 2 else None
        start = 0
        while True:
            pos = text_lower.find(surname, start)
            if pos == -1:
                break
            context = text_lower[pos:pos + 150]
            if name in context:
                if patronymic is None or patronymic in context:
                    return _get_group_above(pos, full_text, filename)
            start = pos + 1
    return "UNKNOWN"


def _get_group_above(pos: int, full_text: str, filename: str) -> str:
    prefix = full_text[:pos]
    matches = list(re.finditer(r'(?i)группа:\s*([а-яa-z0-9\-]+)', prefix))
    if matches:
        raw_grp = matches[-1].group(1)
        return transform_group(raw_grp, filename)
    return "UNKNOWN"


def save_uploaded_pdfs(uploaded_files):
    os.makedirs(BASE_PATH_DATA, exist_ok=True)
    saved_paths = []
    for uploaded_file in uploaded_files:
        file_path = os.path.join(BASE_PATH_DATA, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        saved_paths.append(file_path)
    return saved_paths


def parse_programs_from_pdfs(pdf_paths):
    all_programs = []
    for file_path in pdf_paths:
        if not os.path.exists(file_path):
            continue
        with pdfplumber.open(file_path) as pdf:
            state = {"code": None, "direction": None, "profile": "Не указан"}
            programs = {}
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                flat = re.sub(r'\s+', ' ', text)

                events = []
                for m in re.finditer(r'(\d{2}\.\d{2}\.\d{2})\s*[«"]\s*([^»"]+?)\s*[»"]', flat):
                    events.append((m.start(), 'code_dir', m.group(1), m.group(2).strip()))
                for m in re.finditer(r'(?:Профиль|Специализация):\s*[«"]\s*([^»"]+?)\s*[»"]', flat):
                    events.append((m.start(), 'profile', m.group(1).strip()))
                for m in re.finditer(r'Группа:\s*([А-Яа-я0-9\-]+)', flat):
                    events.append((m.start(), 'group', m.group(1)))
                for m in re.finditer(r'[«"]\s*([^»"]*?)\s*(?:Профиль|Специализация):\s*([^»"]+?)\s*[»"]', flat):
                    reconstructed = f"{m.group(1).strip()} {m.group(2).strip()}".strip()
                    events.append((m.start(), 'profile', reconstructed))

                events.sort(key=lambda x: x[0])
                for ev in events:
                    typ = ev[1]
                    if typ == 'code_dir':
                        state["code"], state["direction"] = ev[2], ev[3]
                    elif typ == 'profile':
                        state["profile"] = ev[2]
                    elif typ == 'group':
                        grp = transform_group(ev[2], os.path.basename(file_path))
                        if state["code"] and state["direction"]:
                            key = (state["code"], state["direction"], state["profile"])
                            if key not in programs:
                                level = "Бакалавриат"
                                qualification = "бакалавр"
                                if "магистр" in state["direction"].lower() or "СВО" in state["direction"]:
                                    level = "Магистратура"
                                    qualification = "магистр"
                                if "СпецВО" in file_path:
                                    level = "СВО"
                                    qualification = "инженер-исследователь"
                                programs[key] = {
                                    "Код": state["code"],
                                    "Направление": state["direction"],
                                    "Профиль": state["profile"],
                                    "Квалификация": qualification,
                                    "Уровень": level,
                                    "Год_поступления": "",
                                    "Группы": set()
                                }
                            programs[key]["Группы"].add(grp)
                            if not programs[key]["Год_поступления"]:
                                year_match = re.search(r'-(\d{2})$', grp)
                                if year_match:
                                    programs[key]["Год_поступления"] = f"20{year_match.group(1)}"
        all_programs.extend(programs.values())

    df = pd.DataFrame([{
        "Код": p["Код"],
        "Направление": p["Направление"],
        "Профиль": p["Профиль"],
        "Квалификация": p["Квалификация"],
        "Уровень": p["Уровень"],
        "Год_поступления": p["Год_поступления"],
        "Группы": ", ".join(sorted(p["Группы"]))
    } for p in all_programs])
    if not df.empty:
        df = df.drop_duplicates().reset_index(drop=True)
    return df


def parse_students_from_pdfs(pdf_paths):
    all_students = []
    for file_path in pdf_paths:
        if not os.path.exists(file_path):
            continue

        full_text = ""
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                full_text += t + "\n"
        full_text = re.sub(r'\s+', ' ', full_text)

        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                for table in page.find_tables():
                    data = table.extract()
                    if not data:
                        continue

                    start = None
                    for i, row in enumerate(data):
                        if row and re.match(r"^\d+$", clean_text(row[0])):
                            start = i
                            break
                    if start is None:
                        continue

                    for row in data[start:]:
                        row = [clean_text(c) for c in row]
                        if not row or not re.match(r"^\d+$", row[0]):
                            continue

                        fio = row[1] if len(row) > 1 else ""
                        fio = re.sub(r'[\n\r]+', ' ', fio).strip()
                        if not re.search(r"[А-ЯЁ]", fio):
                            continue

                        theme = row[2] if len(row) > 2 else ""
                        supervisor = row[3] if len(row) > 3 else ""
                        consultant = row[4] if len(row) > 4 else ""

                        if not consultant and "ассистент" in supervisor.lower():
                            parts = supervisor.split("ассистент", 1)
                            supervisor, consultant = parts[0].strip(), "ассистент " + parts[1].strip()

                        group = find_student_group(fio, full_text, os.path.basename(file_path))

                        all_students.append({
                            "ФИО": fio,
                            "Группа": group,
                            "Тема": theme,
                            "Руководитель": supervisor,
                            "Консультант": consultant,
                            "Рецензент": "",
                            "Присутствие_пред": "",
                            "Оценка_пред": "",
                            "Новая_тема": "",
                            "Новый_рук": "",
                            "Присутствие_защ": "",
                            "Листов": "",
                            "Оценка_защ": ""
                        })

    df = pd.DataFrame(all_students)
    if not df.empty:
        df = df.drop_duplicates(subset=["ФИО", "Группа"]).reset_index(drop=True)
    return df


def save_students_and_programs(students_df, programs_df):
    with pd.ExcelWriter(excel_students, engine='openpyxl') as writer:
        students_df.to_excel(writer, sheet_name='итог', index=False)
        empty_questions = pd.DataFrame(columns=['ФИО', 'Вопросы'])
        empty_questions.to_excel(writer, sheet_name='защ', index=False)
        empty_questions.to_excel(writer, sheet_name='пред', index=False)
    
    programs_df.to_excel(excel_programs, sheet_name='Лист1', index=False)


def analyze_students_fio(students_df, output_dirs):
    from utils import detect_gender_by_patronymic

    problematic = []
    total = 0

    for idx, row in students_df.iterrows():
        fio = row.get('ФИО', '')
        if pd.isna(fio) or not isinstance(fio, str):
            continue
        fio = fio.strip()
        total += 1

        parts = fio.split()
        if len(parts) != 3:
            problematic.append({
                'fio': fio,
                'reason': 'не 3 части',
                'group': row.get('Группа', ''),
                'files': []
            })
            continue

        gender = detect_gender_by_patronymic(fio)
        if gender is None:
            problematic.append({
                'fio': fio,
                'reason': 'пол не определён по отчеству',
                'group': row.get('Группа', ''),
                'files': []
            })
            continue

    for p in problematic:
        fio_clean = p['fio'].replace(' ', '_').replace('.', '')
        group = p['group']
        found_files = []
        for dir_name, dir_path in output_dirs.items():
            if not os.path.exists(dir_path):
                continue
            for filename in os.listdir(dir_path):
                if filename.endswith('.docx'):
                    if fio_clean in filename or (group and group in filename):
                        found_files.append({
                            'path': os.path.join(dir_path, filename),
                            'name': filename,
                            'type': dir_name
                        })
        p['files'] = found_files

    summary = f"Всего студентов: {total}. Проблемных: {len(problematic)} ({len(problematic)/total*100:.1f}%)"
    return problematic, summary