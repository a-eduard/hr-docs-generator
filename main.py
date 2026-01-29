import streamlit as st
import pandas as pd
import os
import io
import zipfile
import re
import pdfplumber
from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm
from num2words import num2words
from datetime import date
import pymorphy3
from PIL import Image, ImageOps

# --- 1. НАСТРОЙКИ ---
st.set_page_config(page_title="Smart HR Architect", layout="wide", page_icon="🏗️")

st.markdown("""
<style>
    [data-testid="stFileUploaderDropzone"] div div::before {content:"";}
    [data-testid="stFileUploaderDropzone"] div div span {display:none;}
    [data-testid="stFileUploaderDropzone"] {min-height: 80px; padding: 10px;}
</style>
""", unsafe_allow_html=True)

# --- 2. ПОДКЛЮЧЕНИЕ МОДУЛЕЙ ---
try:
    morph = pymorphy3.MorphAnalyzer()
except:
    pass

try:
    from ai_utils import generate_ai_duties, extract_data_from_egrul
except ImportError:
    def generate_ai_duties(p): return ""
    def extract_data_from_egrul(t): return None

# --- 3. STATE ---
keys = ["c_name", "c_short_name", "c_inn", "c_kpp", "c_ogrn", "c_address", "c_boss", "c_boss_pos", "c_opf"]
for k in keys:
    if k not in st.session_state:
        st.session_state[k] = ""

# --- 4. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---

def clean_val(val):
    if pd.isna(val): return None
    s = str(val).strip()
    if s == "" or s.lower() == "nan": return None
    return s

def build_passport_string(row):
    row_lower = {str(k).lower().strip(): v for k, v in row.items()}
    passport_num = ""
    for key in row_lower:
        if any(x in key for x in ["паспорт", "серия", "номер", "документ"]):
            val = clean_val(row_lower[key])
            if val:
                if val.isdigit() and len(val) == 10:
                    val = f"{val[:4]} {val[4:]}"
                passport_num = val
                break
    
    issued_by = ""
    for key in row_lower:
        if any(x in key for x in ["кем выдан", "выдан", "кем"]):
            if "дата" in key or "когда" in key: continue
            val = clean_val(row_lower[key])
            if val: issued_by = val; break
    
    date_issued = ""
    for key in row_lower:
        if any(x in key for x in ["дата", "когда", "число"]):
            val = clean_val(row_lower[key])
            if val:
                try: 
                    date_issued = pd.to_datetime(val, dayfirst=True).strftime("%d.%m.%Y")
                except: date_issued = val 
                break

    parts = []
    if passport_num: parts.append(f"Паспорт: {passport_num}")
    else: parts.append("Паспорт: __________________")
    if issued_by: parts.append(f"выдан {issued_by}")
    if date_issued: parts.append(f"дата выдачи {date_issued}")
    return ", ".join(parts)

def clean_case(text):
    if not text: return ""
    text = str(text)
    # Если текст весь ВЕРХНИМ РЕГИСТРОМ (как в ЕГРЮЛ часто бывает), делаем первую заглавной
    # Но если там смешанный регистр (ООО "Ромашка"), не трогаем
    upper_chars = sum(1 for c in text if c.isupper())
    if len(text) > 4 and (upper_chars / len(text)) > 0.8:
        return text.capitalize() # БЫЛО: text.capitalize(). ТЕПЕРЬ: можно сделать умнее, но пока оставим
    return text

def try_read_csv(file_source, encoding, sep):
    try:
        if hasattr(file_source, 'seek'): file_source.seek(0)
        df = pd.read_csv(file_source, sep=sep, encoding=encoding, on_bad_lines='skip')
        if len(df.columns) > 1: return df
    except: pass
    return None

def load_data_file(key_label, local_filename):
    file_source = None
    local_path_xlsx = f"data/{local_filename}.xlsx"
    local_path_csv = f"data/{local_filename}.csv"
    
    uploaded = st.sidebar.file_uploader(f"Загрузить {key_label}", type=["csv", "xlsx"], key=local_filename)
    if uploaded: file_source = uploaded
    elif os.path.exists(local_path_xlsx): file_source = local_path_xlsx
    elif os.path.exists(local_path_csv): file_source = local_path_csv
            
    if not file_source: return None

    try:
        df = None
        if hasattr(file_source, 'name') and file_source.name.endswith('.xlsx'):
             df = pd.read_excel(file_source)
        elif isinstance(file_source, str) and file_source.endswith('.xlsx'):
             df = pd.read_excel(file_source)
        else:
            df = try_read_csv(file_source, 'cp1251', ';')
            if df is None: df = try_read_csv(file_source, 'utf-8-sig', ',')
            if df is None: df = try_read_csv(file_source, 'cp1251', ',')

        if df is not None:
            df.columns = df.columns.str.strip()
            if 'ФИО' in df.columns:
                if 'Должность' in df.columns:
                    df['search_key'] = df['ФИО'] + " — " + df['Должность']
                else:
                    df['search_key'] = df['ФИО']
            return df
        return None
    except Exception as e:
        st.sidebar.error(f"Ошибка {key_label}: {e}")
        return None

def parse_egrul_pdf_ai(pdf_file):
    full_text = ""
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted: full_text += extracted + "\n"
    except Exception as e:
        return None, f"Ошибка PDF: {e}"
    if not full_text: return None, "PDF пустой."
    data = extract_data_from_egrul(full_text)
    if not data: return None, "AI не вернул данные."
    return data, None

def make_times_new_roman(text):
    if not text: return ""
    rt = RichText()
    rt.add(str(text), font='Times New Roman', size=24)
    return rt

# --- ФУНКЦИИ ОБРАБОТКИ ТЕКСТА ---

def get_inflected(text: str, case_tag: str) -> str:
    if not text or 'morph' not in globals(): return text
    res = []
    for w in text.split():
        try:
            is_capitalized = w[0].isupper()
            p = morph.parse(w)[0]
            inflected = p.inflect({case_tag})
            
            if inflected:
                word = inflected.word
                if is_capitalized: word = word.capitalize()
                res.append(word)
            else:
                res.append(w)
        except:
            res.append(w)
    
    final_str = " ".join(res)
    if final_str:
        return final_str[0].upper() + final_str[1:]
    return ""

def get_initials(full_name: str) -> str:
    if not full_name: return ""
    p = full_name.split()
    if len(p) >= 3:
        return f"{p[0].capitalize()} {p[1][0].upper()}.{p[2][0].upper()}."
    return full_name

def get_gender_word(fio: str, word_masc: str, word_fem: str) -> str:
    if not fio: return word_masc
    parts = fio.split()
    if len(parts) >= 3:
        patr = parts[2].lower()
        if patr.endswith("вна") or patr.endswith("чна") or patr.endswith("шна"):
            return word_fem
        if patr.endswith("вич"):
            return word_masc
    if len(parts) >= 2 and 'morph' in globals():
        try:
            parsed = morph.parse(parts[1])[0] 
            if 'femn' in parsed.tag: return word_fem
        except: pass
    return word_masc

def increment_doc_number(base_num: str, step: int) -> str:
    if step == 0: return base_num
    match = re.search(r'\d+', base_num)
    if match:
        number_str = match.group()
        new_number = int(number_str) + step
        return base_num.replace(number_str, str(new_number), 1)
    return f"{base_num}-{step + 1}"

# --- ИЗОБРАЖЕНИЯ ---

def trim_whitespace(img):
    try:
        if img.mode != "RGBA":
            img = img.convert("RGBA")
        alpha = img.split()[-1]
        bbox = alpha.getbbox()
        if bbox: return img.crop(bbox)
        return img
    except: return img

def create_overlay_image(sign_path, stamp_path):
    try:
        if not sign_path or not os.path.exists(sign_path): return None
        sign_img = Image.open(sign_path).convert("RGBA")
        sign_img = trim_whitespace(sign_img)
        
        if stamp_path and os.path.exists(stamp_path):
            stamp_img = Image.open(stamp_path).convert("RGBA")
            stamp_img = trim_whitespace(stamp_img)
            target_h = int(sign_img.height * 1.3)
            if target_h < 150: target_h = 150 
            ratio = target_h / stamp_img.height
            target_w = int(stamp_img.width * ratio)
            stamp_img = stamp_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
            
            shift_x = int(sign_img.width * 0.6) 
            canvas_w = max(sign_img.width, shift_x + stamp_img.width) + 10
            canvas_h = max(sign_img.height, stamp_img.height) + 10
            new_img = Image.new('RGBA', (canvas_w, canvas_h), (255, 255, 255, 0))
            
            y_sign = (canvas_h - sign_img.height) // 2
            new_img.paste(sign_img, (0, y_sign), sign_img)
            y_stamp = (canvas_h - stamp_img.height) // 2
            new_img.paste(stamp_img, (shift_x, y_stamp), stamp_img)
            
            temp_path = "data/signatures/temp_combo.png"
            new_img.save(temp_path, format="PNG")
            return temp_path
            
        temp_path = "data/signatures/temp_sign_trimmed.png"
        sign_img.save(temp_path, format="PNG")
        return temp_path
    except: return sign_path

def get_image_object(doc, filename_or_path, width_mm, do_trim=True):
    if not filename_or_path: return "[ПУСТОЕ ИМЯ]"
    
    path = filename_or_path
    if not os.path.exists(path):
        base = os.path.join("data", "signatures", filename_or_path)
        if os.path.exists(base): path = base
        elif os.path.exists(base + ".png"): path = base + ".png"
        elif os.path.exists(base + ".jpg"): path = base + ".jpg"
        elif os.path.exists(base + ".jpeg"): path = base + ".jpeg"
        else:
            return f"[НЕТ ФАЙЛА: {filename_or_path}]"

    final_path = path
    if do_trim and "temp" not in path: 
        try:
            img = Image.open(path)
            img = trim_whitespace(img)
            trimmed_name = f"trimmed_{os.path.basename(path)}"
            final_path = os.path.join("data", "signatures", trimmed_name)
            img.save(final_path, format="PNG")
        except Exception as e:
            return f"[ОШИБКА ОБРАБОТКИ: {e}]"

    try: 
        return InlineImage(doc, final_path, width=Mm(width_mm))
    except Exception as e:
        return f"[ОШИБКА ВСТАВКИ: {e}]"

# --- 5. ИНТЕРФЕЙС ---

st.sidebar.header("📂 Базы данных")
df_emp = load_data_file("Сотрудников", "employees")
df_resp = load_data_file("Ответственных", "responsible")

st.sidebar.divider()
st.sidebar.header("⚙️ Настройки")
use_ai_duties = st.sidebar.toggle("🤖 Генерировать обязанности", value=True)
selected_style = st.sidebar.selectbox("Стиль шаблонов", ["style1", "style2", "style3", "style4", "style5", "style6"], index=0)

with st.sidebar.expander("✒️ Загрузить подписи сотрудников"):
    uploaded_sigs = st.file_uploader("Файлы (название = ФИО)", type=["png", "jpg"], accept_multiple_files=True)
    if uploaded_sigs:
        if not os.path.exists("data/signatures"): os.makedirs("data/signatures")
        for f in uploaded_sigs:
            with open(os.path.join("data/signatures", f.name), "wb") as dest:
                dest.write(f.getbuffer())
        st.success(f"Загружено {len(uploaded_sigs)} подписей")

st.title("🏗️ Генератор PRO (v8.0)")
st.markdown("---")

if df_emp is None:
    st.info("👈 Загрузите базу Сотрудников.")
    st.stop()

col_left, col_right = st.columns([1, 1.3])

with col_left:
    st.subheader("1. Выбор персонала")
    options = df_emp['search_key'].unique()
    selected_emp_keys = st.multiselect("Сотрудники:", options)
    
    st.markdown("")
    st.write("🧑‍💼 **Ответственное лицо:**")
    selected_resp_key = "--- Не указывать ---"
    
    if df_resp is not None:
        resp_options = ["--- Не указывать ---"] + list(df_resp['search_key'].unique())
        selected_resp_key = st.selectbox("Кто упоминается в документах:", resp_options)

    st.markdown("---")
    st.subheader("2. Параметры")
    c1, c2 = st.columns(2)
    with c1:
        start_doc_num = st.text_input("Номер документа", "12-К")
        salary = st.number_input("Оклад", value=120000, step=5000)
    with c2:
        doc_date = st.date_input("Дата", date.today())
        city = st.text_input("Город", "Москва")

with col_right:
    st.subheader("3. Данные Работодателя")
    uploaded_pdf = st.file_uploader("1. Загрузить ЕГРЮЛ (PDF)", type=["pdf"])
    
    if uploaded_pdf:
        if st.button("🚀 Распознать через YandexGPT", type="secondary"):
            with st.spinner("Анализирую..."):
                extracted, err = parse_egrul_pdf_ai(uploaded_pdf)
                if err: st.error(err)
                elif extracted:
                    if "inn" in extracted: st.session_state.c_inn = extracted["inn"]
                    if "kpp" in extracted: st.session_state.c_kpp = extracted["kpp"]
                    if "ogrn" in extracted: st.session_state.c_ogrn = extracted["ogrn"]
                    
                    # ОБНОВЛЕНО: Используем исходный регистр для названия (без clean_case), 
                    # или аккуратно чистим, но сохраняем структуру
                    name_extracted = extracted.get("name", "")
                    # Если все капсом - делаем красиво, если нет - оставляем как есть
                    if name_extracted.isupper():
                        st.session_state.c_name = clean_case(name_extracted)
                    else:
                        st.session_state.c_name = name_extracted
                        
                    if "short_name" in extracted: st.session_state.c_short_name = extracted["short_name"]
                    if "address" in extracted: st.session_state.c_address = clean_case(extracted["address"])
                    if "boss_name" in extracted: st.session_state.c_boss = clean_case(extracted["boss_name"])
                    if "boss_pos" in extracted: st.session_state.c_boss_pos = clean_case(extracted["boss_pos"])
                    if "opf" in extracted: st.session_state.c_opf = clean_case(extracted["opf"])
                    st.success(f"Распознано: {extracted.get('name')}")
                    st.rerun()

    st.markdown("##### 🖃 Печать и Подпись Директора:")
    c_stamp, c_dir = st.columns(2)
    stamp_path_temp = None
    director_path_temp = None
    if not os.path.exists("data/signatures"): os.makedirs("data/signatures")

    with c_stamp:
        up_stamp = st.file_uploader("Печать (PNG)", type=["png"], key="u_stamp")
        if up_stamp:
            stamp_path_temp = "data/signatures/temp_stamp_session.png"
            with open(stamp_path_temp, "wb") as f: f.write(up_stamp.getbuffer())

    with c_dir:
        up_dir = st.file_uploader("Подпись Директора (PNG)", type=["png"], key="u_dir")
        if up_dir:
            director_path_temp = "data/signatures/temp_director_session.png"
            with open(director_path_temp, "wb") as f: f.write(up_dir.getbuffer())

    st.markdown("##### 📝 Реквизиты:")
    st.text_input("Орг.-правовая форма", key="c_opf")
    st.text_input("Название (без ОПФ)", key="c_name")
    st.text_input("Сокращенное название", key="c_short_name")
    c_i, c_k, c_o = st.columns([1, 1, 1])
    with c_i: st.text_input("ИНН", key="c_inn")
    with c_k: st.text_input("КПП", key="c_kpp")
    with c_o: st.text_input("ОГРН", key="c_ogrn")
    st.text_area("Юридический адрес", key="c_address", height=68)
    c_b1, c_b2 = st.columns(2)
    with c_b1: st.text_input("ФИО Директора", key="c_boss")
    with c_b2: st.text_input("Должность", key="c_boss_pos")

st.markdown("---")
if st.button("🚀 Сформировать документы", type="primary", use_container_width=True):
    
    if not selected_emp_keys:
        st.error("❌ Выберите сотрудников!")
        st.stop()

    # --- ПОДГОТОВКА ОБЩИХ ДАННЫХ ---
    tasks = []
    for key in selected_emp_keys:
        row = df_emp[df_emp['search_key'] == key].iloc[0]
        tasks.append({"data": row, "role": "emp"})
        
    opf = st.session_state.c_opf.strip()
    name = st.session_state.c_name.strip()
    
    # === ИСПРАВЛЕНИЕ: ПРЯМАЯ СКЛЕЙКА ===
    # Больше программа не добавляет никаких кавычек автоматически.
    # Что написано в полях "ОПФ" и "Название" — то и будет в документе.
    full_company_name = f"{opf} {name}".strip()
    # ===================================
             
    b_name = st.session_state.c_boss
    b_pos = st.session_state.c_boss_pos
    short_name_val = st.session_state.c_short_name if st.session_state.c_short_name else full_company_name
    
    resp_name_str = ""
    resp_pos_str = ""
    resp_doc_str = ""
    if df_resp is not None and selected_resp_key != "--- Не указывать ---":
        r_row = df_resp[df_resp['search_key'] == selected_resp_key].iloc[0]
        resp_name_str = r_row.get('ФИО', '')
        resp_pos_str = r_row.get('Должность', '')
        for k_resp, v_resp in r_row.items():
            if any(x in str(k_resp).lower() for x in ["основание", "документ", "доверенность"]):
                 resp_doc_str = str(v_resp)
                 break

    reqs_str = f"{full_company_name}\nЮр. адрес: {st.session_state.c_address}\nИНН {st.session_state.c_inn}, КПП {st.session_state.c_kpp}, ОГРН {st.session_state.c_ogrn}"
    rt_reqs = make_times_new_roman(reqs_str)

    date_short = doc_date.strftime("%d.%m.%Y") + " г."
    months_ru = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"]
    date_full = f"«{doc_date.day:02d}» {months_ru[doc_date.month - 1]} {doc_date.year} г."

    combo_path = None
    if director_path_temp:
        combo_path = create_overlay_image(director_path_temp, stamp_path_temp)
    
    zip_buf = io.BytesIO()
    files_ok = 0
    progress = st.progress(0)
    
    # 2. ДОБАВЛЯЕМ СТИЛЬ В ИМЕНА ФАЙЛОВ
    style_suffix = f"_{selected_style}"
    
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        
        info_text = f"""Дата генерации: {date.today()}
Компания: {full_company_name}
Использован стиль: {selected_style}
Сотрудников обработано: {len(tasks)}
        """
        zf.writestr("00_INFO.txt", info_text)

        company_ctx = {
            "city": city, "contract_date": date_short, "date_ru": date_full,
            "company_name": full_company_name, "company_short": short_name_val,
            "company_address": st.session_state.c_address,
            "company_inn": st.session_state.c_inn, "company_kpp": st.session_state.c_kpp, "company_ogrn": st.session_state.c_ogrn,
            "head_name": b_name, "head_pos": b_pos, "head_short": get_initials(b_name),
            "head_name_gen": get_inflected(b_name, 'gent'), 
            "head_pos_gen": get_inflected(b_pos, 'gent'),
            "head_name_accs": get_inflected(b_name, 'accs'), 
            "head_pos_accs": get_inflected(b_pos, 'accs'),
            "head_pos_datv": get_inflected(b_pos, 'datv'),
            "employer_reqs": rt_reqs,
            "director_combo": get_image_object(DocxTemplate(io.BytesIO()), combo_path, 45, False) if combo_path else "",
        }

        # --- 1. ОПИСЬ ---
        inventory_path = "templates/inventory.docx"
        if os.path.exists(inventory_path):
            try:
                doc_inv = DocxTemplate(inventory_path)
                doc_inv.render(company_ctx)
                tmp_inv = io.BytesIO()
                doc_inv.save(tmp_inv)
                zf.writestr(f"00_Опись{style_suffix}.docx", tmp_inv.getvalue())
                files_ok += 1
            except Exception as e: pass

        # --- 2. СВОДНЫЙ ПРИКАЗ ---
        style_num = selected_style.replace("style", "") 
        order_tmpl_path = f"templates/orders/{style_num}.docx"
        
        if os.path.exists(order_tmpl_path):
            try:
                doc_ord = DocxTemplate(order_tmpl_path)
                employees_list = []
                for t in tasks:
                    emp_data = t["data"]
                    fio = emp_data['ФИО']
                    pos = emp_data.get('Должность', '')
                    
                    employees_list.append({
                        "name": fio,
                        "short": get_initials(fio),
                        "pos": pos,
                        "name_gen": get_inflected(fio, 'gent'),
                        "pos_gen": get_inflected(pos, 'gent'),
                        "name_accs": get_inflected(fio, 'accs'), 
                        "pos_accs": get_inflected(pos, 'accs'),
                        "accepted": get_gender_word(fio, "принят", "принята"),
                        "appointed": get_gender_word(fio, "назначен", "назначена"),
                        "sign": get_image_object(doc_ord, fio, 20, True) 
                    })
                
                ctx_ord = company_ctx.copy()
                ctx_ord["col_employees"] = employees_list
                if combo_path: 
                    ctx_ord["director_combo"] = get_image_object(doc_ord, combo_path, 45, False)
                    ctx_ord["director_sign"] = get_image_object(doc_ord, director_path_temp, 30, True)

                doc_ord.render(ctx_ord)
                tmp_ord = io.BytesIO()
                doc_ord.save(tmp_ord)
                zf.writestr(f"00_Сводный_приказ_Ответственные{style_suffix}.docx", tmp_ord.getvalue())
                files_ok += 1
            except Exception as e:
                st.error(f"Ошибка сводного приказа: {e}")

        # --- 3. ПРИКАЗ НА ОТВЕТСТВЕННОГО ---
        target_resp = {}
        if df_resp is not None and selected_resp_key != "--- Не указывать ---":
             r_row = df_resp[df_resp['search_key'] == selected_resp_key].iloc[0]
             target_resp = { "name": r_row['ФИО'], "pos": r_row.get('Должность', ''), "is_director": False }
             filename_resp = f"Приказ_Ответственный_{get_initials(r_row['ФИО'])}"
        else:
             target_resp = { "name": b_name, "pos": b_pos, "is_director": True }
             filename_resp = f"Приказ_Ответственный_Директор"
             
        if os.path.exists(order_tmpl_path):
             try:
                doc_r = DocxTemplate(order_tmpl_path)
                person_data = {
                    "name": target_resp["name"],
                    "short": get_initials(target_resp["name"]),
                    "pos": target_resp["pos"],
                    "name_gen": get_inflected(target_resp["name"], 'gent'),
                    "pos_gen": get_inflected(target_resp["pos"], 'gent'),
                    "name_accs": get_inflected(target_resp["name"], 'accs'),
                    "pos_accs": get_inflected(target_resp["pos"], 'accs'),
                    "accepted": get_gender_word(target_resp["name"], "принят", "принята"),
                    "appointed": get_gender_word(target_resp["name"], "назначен", "назначена"),
                    "sign": get_image_object(doc_r, director_path_temp, 30, True) if target_resp["is_director"] else get_image_object(doc_r, target_resp["name"], 20, True)
                }
                ctx_r = company_ctx.copy()
                ctx_r["col_employees"] = [person_data]
                if combo_path: 
                    ctx_r["director_combo"] = get_image_object(doc_r, combo_path, 45, False)
                    ctx_r["director_sign"] = get_image_object(doc_r, director_path_temp, 30, True)
                doc_r.render(ctx_r)
                tmp_r = io.BytesIO()
                doc_r.save(tmp_r)
                zf.writestr(f"00_{filename_resp}{style_suffix}.docx", tmp_r.getvalue())
                files_ok += 1
             except Exception as e: pass

        # --- 4. ЛИЧНЫЕ ДОКУМЕНТЫ ---
        for i, task in enumerate(tasks):
            emp = task["data"]
            role = task["role"]
            progress.progress((i + 1) / len(tasks))
            
            doc_num = increment_doc_number(start_doc_num, i)
            ai_duties = ""
            if use_ai_duties and role == "emp":
                try: ai_duties = generate_ai_duties(emp['Должность'])
                except: ai_duties = "Ошибка генерации"

            full_passport_str = build_passport_string(emp)
            pos_nom = emp.get('Должность', '')
            
            context = company_ctx.copy()
            context.update({
                "doc_number": doc_num,
                "resp_name": resp_name_str, "resp_pos": resp_pos_str, "resp_doc": resp_doc_str,
                "resp_short": get_initials(resp_name_str),
                "employee_name": emp['ФИО'], "employee_short": get_initials(emp['ФИО']),
                "employee_pos": pos_nom,
                "employee_pos_gen": get_inflected(pos_nom, 'gent'),
                "employee_pos_dat": get_inflected(pos_nom, 'datv'),
                "employee_pos_accs": get_inflected(pos_nom, 'accs'),
                "salary_digits": f"{salary:,}".replace(",", " "),
                "salary_words": num2words(salary, lang='ru').capitalize() + " рублей 00 копеек",
                "employee_reqs": make_times_new_roman(full_passport_str),
                "employee_passport": f"{full_passport_str}",
                "ai_duties": make_times_new_roman(ai_duties)
            })

            paths = {
                "Трудовой_договор": f"templates/contracts/{selected_style}.docx", 
                "Приказ": "templates/order.docx",
                "Должностная": f"templates/instructions/{emp.get('Должность','').strip()}_{selected_style}.docx"
            }
            
            for name, path in paths.items():
                if role == "resp" and name == "Должностная": continue
                if os.path.exists(path):
                    try:
                        doc = DocxTemplate(path)
                        if combo_path: context["director_combo"] = get_image_object(doc, combo_path, 45, do_trim=False)
                        if director_path_temp: context["director_sign"] = get_image_object(doc, director_path_temp, 30, do_trim=True)
                        context["employee_sign"] = get_image_object(doc, emp['ФИО'], 20, do_trim=True)
                        if resp_name_str: context["resp_sign"] = get_image_object(doc, resp_name_str, 20, do_trim=True)
                        
                        doc.render(context)
                        tmp = io.BytesIO()
                        doc.save(tmp)
                        safe_fio = get_initials(emp['ФИО']).replace(".", "")
                        suffix = "_RESP" if role == "resp" else ""
                        zf.writestr(f"{i+1:02d}_{safe_fio}{suffix}_{name}{style_suffix}.docx", tmp.getvalue())
                        files_ok += 1
                    except Exception: pass
    progress.progress(100)
    
    if files_ok > 0:
        zip_buf.seek(0)
        st.success(f"✅ Файлов создано: {files_ok}")
        st.download_button("💾 Скачать ZIP", zip_buf, f"Docs_{date.today()}.zip", "application/zip")
    else:
        st.error("Шаблоны не найдены!")