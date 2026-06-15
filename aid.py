import streamlit as st
import os
import json
from google import genai
from docx import Document
from pypdf import PdfReader
import pandas as pd
import io

# Настройка страницы
st.set_page_config(page_title="База знаний и Диспетчер ИИ", page_icon="🧠", layout="wide")

st.title("🧠 Интеллектуальная база знаний отдела (Gemini)")
st.write("Загружайте инструкции, мануалы и таблицы (Word, PDF, Excel). ИИ проанализирует базу и найдет решение или ответственного.")

# --- ГАРАНТИРОВАННОЕ СОЗДАНИЕ ПАПОК ---
BASE_DIR = "knowledge_base"
DOCS_DIR = os.path.join(BASE_DIR, "documents")
os.makedirs(DOCS_DIR, exist_ok=True)

# --- ИНИЦИАЛИЗАЦИЯ GEMINI ---
st.sidebar.header("🔑 Настройки ИИ")

# Проверяем, зашит ли ключ в Secrets приложения
secrets_key = os.environ.get("GEMINI_API_KEY", "")

# Поле ввода в сайдбаре (если в Secrets пусто, можно ввести вручную)
api_key = st.sidebar.text_input(
    "Введите Gemini API Key:", 
    type="password", 
    value=secrets_key if secrets_key else ""
)

# Выбираем, какой ключ использовать (приоритет секретам)
final_key = secrets_key if secrets_key else api_key

client = None
if final_key:
    try:
        # Инициализируем клиент строго с выбранным ключом
        client = genai.Client(api_key=final_key)
    except Exception as e:
        st.sidebar.error(f"Ошибка инициализации ИИ: {e}")
else:
    st.sidebar.warning("🔑 Пожалуйста, введите API ключ или пропишите его в Secrets.")

# --- ФУНКЦИИ ДЛЯ ЧТЕНИЯ ФОРМАТОВ ---

def parse_docx(file_bytes):
    try:
        doc = Document(io.BytesIO(file_bytes))
        full_text = []
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text:
                    full_text.append(" | ".join(row_text))
        return "\n".join(full_text)
    except Exception as e:
        return f"Ошибка при чтении файла Word: {e}"

def parse_pdf(file_bytes):
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        full_text = []
        for page in reader.pages:
            text = page.extract_text()
            if text:
                full_text.append(text)
        return "\n".join(full_text)
    except Exception as e:
        return f"Ошибка при чтении файла PDF: {e}"

def parse_excel(file_bytes):
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        full_text = []
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            df = df.fillna("")
            if not df.empty:
                full_text.append(f"--- Лист таблицы: {sheet_name} ---")
                headers = " | ".join(df.columns.astype(str))
                full_text.append(f"Колонки: {headers}")
                for idx, row in df.iterrows():
                    row_values = [str(val).strip() for val in row.values]
                    full_text.append(f"Строка {idx+1}: " + " | ".join(row_values))
        return "\n".join(full_text)
    except Exception as e:
        return f"Ошибка при чтении файла Excel: {e}"

# --- БОКОВАЯ ПАНЕЛЬ: ЗАГРУЗКА И УДАЛЕНИЕ ---
with st.sidebar:
    st.write("---")
    st.header("📥 Загрузка инструкций и таблиц")
    
    doc_type = st.selectbox("Тип документа:", ["Должностная инструкция", "Инструкция к программе / Ошибка", "Таблица Excel / Реестр"])
    doc_title = st.text_input("Название (ФИО, программа или имя таблицы):")
    uploaded_file = st.file_uploader("Выложите файл (.docx, .pdf, .xlsx):", type=["docx", "pdf", "xlsx"])
    
    save_btn = st.button("💾 Сохранить в базу знаний", type="secondary")
    
    if save_btn and doc_title and uploaded_file:
        file_bytes = uploaded_file.read()
        file_ext = uploaded_file.name.split(".")[-1].lower()
        
        if file_ext == "docx":
            parsed_text = parse_docx(file_bytes)
        elif file_ext == "pdf":
            parsed_text = parse_pdf(file_bytes)
        elif file_ext in ["xlsx", "xls"]:
            parsed_text = parse_excel(file_bytes)
        else:
            parsed_text = "Неподдерживаемый формат файла."
        
        if "Ошибка при чтении" in parsed_text:
            st.error(parsed_text)
        else:
            filename = f"{doc_type.replace(' ', '_')}_{doc_title.replace(' ', '_')}.json"
            filepath = os.path.join(DOCS_DIR, filename)
            
            data = {
                "type": doc_type,
                "title": doc_title,
                "content": parsed_text
            }
            
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            st.success(f"Документ '{doc_title}' успешно добавлен!")
            st.rerun()

    # --- СЕКЦИЯ УДАЛЕНИЯ ДОКУМЕНТОВ ---
    st.write("---")
    st.subheader("📚 Список документов в базе:")
    if os.path.exists(DOCS_DIR):
        files = [f for f in os.listdir(DOCS_DIR) if f.endswith('.json')]
        if files:
            for file in files:
                filepath = os.path.join(DOCS_DIR, file)
                with open(filepath, "r", encoding="utf-8") as f:
                    meta = json.load(f)
                
                if meta["type"] == "Должностная инструкция":
                    icon = "👤"
                elif meta["type"] == "Таблица Excel / Реестр":
                    icon = "📊"
                else:
                    icon = "💻"
                
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.write(f"{icon} **{meta['title']}**")
                with col2:
                    if st.button("🗑️", key=file, help=f"Удалить {meta['title']}"):
                        os.remove(filepath)
                        st.success("Удалено!")
                        st.rerun()
        else:
            st.info("База знаний пуста. Загрузите файлы.")

# --- ГЛАВНАЯ СТРАНИЦА: ЗАПРОСЫ, ОШИБКИ И АНАЛИЗ ---
st.subheader("📝 Запрос к ИИ (Поручение, Ошибка или поиск в Таблице)")
user_query = st.text_area("Вставьте текст задачи, описание технического сбоя или вопрос по таблице Excel:", height=150,
                         placeholder="Пример: 'Кто согласно таблице графика дежурит в эту субботу?' или 'Найди по реестру, за кем закреплен проект АВР'")

if st.button("🔍 Запустить анализ базы знаний", type="primary"):
    if not final_key or client is None:
        st.error("Ошибка: Не указан или неверен API-ключ Gemini!")
    elif not user_query.strip():
        st.warning("Пожалуйста, введите текст запроса.")
    else:
        all_docs_context = ""
        if os.path.exists(DOCS_DIR):
            files = [f for f in os.listdir(DOCS_DIR) if f.endswith('.json')]
            for idx, file in enumerate(files, 1):
                with open(os.path.join(DOCS_DIR, file), "r", encoding="utf-8") as f:
                    meta = json.load(f)
                    all_docs_context += f"--- ДОКУМЕНТ №{idx} ---\nТип: {meta['type']}\nНазвание: {meta['title']}\nСодержимое:\n{meta['content']}\n\n"
        
        if not all_docs_context:
            st.error("Ошибка: База знаний пуста! Пожалуйста, загрузите файлы в левой панели.")
        else:
            with st.spinner("Gemini штудирует базу данных, мануалы и таблицы..."):
                try:
                    system_instruction = (
                        "Ты — суперинтеллектуальный аналитик и корпоративный помощник. В твоем распоряжении база знаний, "
                        "состоящая из текстовых инструкций (Word/PDF) и структурированных таблиц (Excel).\n\n"
                        "Твоя задача — внимательно изучить предоставленный контекст и ответить на запрос пользователя:\n"
                        "1. Если вопрос касается ТАБЛИЦЫ/РЕЕСТРА (Excel): найди нужные строки, сопоставь данные и выдай точный "
                        "аналитический ответ (например, кто дежурит, за кем закреплена задача, какой статус у документа).\n"
                        "2. Если это ТЕХНИЧЕСКАЯ ОШИБКА: найди в мануалах пошаговый алгоритм исправления.\n"
                        "3. Если это ПОРУЧЕНИЕ: определи исполнителя по должностным инструкциям.\n\n"
                        "Отвечай строго по делу, опираясь только на факты из загруженных документов. Будь точен и лаконичен."
                    )
                    
                    full_prompt = f"""Вот полная база загруженных документов, инструкций и таблиц вашего отдела:
{all_docs_context}

---
ЗАПРОС ПОЛЬЗОВАТЕЛЯ:
"{user_query}"
---

Используя данные выше, сформируй точный ответ. Если информация взята из таблицы Excel, укажи конкретные строки или листы.
"""
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=full_prompt,
                        config={"system_instruction": system_instruction, "temperature": 0.2}
                    )
                    
                    st.success("🤖 Ответ сформирован!")
                    st.subheader("🎯 Результат анализа:")
                    st.markdown(response.text)
                    
                except Exception as e:
                    error_msg = str(e)
                    if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                        st.error("⏳ Исчерпан дневной лимит бесплатных запросов к ИИ (20 запросов в день). Пожалуйста, подождите или обновите API-ключ!")
                    elif "401" in error_msg or "UNAUTHENTICATED" in error_msg:
                        st.error("❌ Ошибка авторизации: Указан неверный или недействительный API-ключ Gemini. Пожалуйста, перепроверьте его в Secrets или введите вручную слева!")
                    else:
                        st.error(f"Произошла ошибка при вызове Gemini: {e}")
