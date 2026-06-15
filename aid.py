import streamlit as st
import os
import json
from google import genai
from docx import Document
from pypdf import PdfReader
import pandas as pd
import io
import sys

# Принудительно настраиваем окружение на UTF-8 для защиты от ошибок кодировки
import codecs
if sys.stdout.encoding != 'utf-8':
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# Настройка страницы
st.set_page_config(page_title="База знаний и Диспетчер ИИ", page_icon="🧠", layout="wide")

st.title("🧠 Интеллектуальная база знаний отдела (Gemini)")
st.write("Загружайте инструкции, мануалы и таблицы (Word, PDF, Excel). ИИ проанализирует базу и найдет решение или ответственного.")

# --- ГАРАНТИРОВАННОЕ СОЗДАНИЕ ПАПОК ---
BASE_DIR = "knowledge_base"
DOCS_DIR = os.path.join(BASE_DIR, "documents")
os.makedirs(DOCS_DIR, exist_ok=True)

# --- ИНИЦИАЛИЗАЦИЯ GEMINI (СТРОГО ВРУЧНУЮ) ---
st.sidebar.header("🔑 Настройки ИИ")

api_key = st.sidebar.text_input(
    "Введите Gemini API Key:", 
    type="password", 
    placeholder="Вставьте ключ AIzaSy..."
)

client = None
if api_key.strip():
    try:
        client = genai.Client(api_key=api_key.strip())
    except Exception as e:
        st.sidebar.error(f"Ошибка инициализации ИИ: {e}")
else:
    st.sidebar.warning("🔑 Пожалуйста, введите ваш рабочий API ключ слева для активации ИИ.")

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
        # Считываем страницы ровно так, как они идут физически
        for idx, page in enumerate(reader.pages, 1):
            text = page.extract_text()
            if text:
                full_text.append(f"[Страница {idx}]\n{text}")
        return "\n".join(full_text)
    except Exception as e:
        return f"Ошибка при чтении файла PDF: {e}"

def parse_excel(file_bytes):
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        full_text = []
        for sheet_name in excel_file.sheet_names:
            # header=None отключает превращение первой строки в заголовки колонок. 
            # Теперь отсчет строк идет честно, сверху вниз.
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            df = df.fillna("")
            if not df.empty:
                full_text.append(f"--- Лист таблицы: {sheet_name} ---")
                for idx, row in df.iterrows():
                    row_values = [str(val).strip() for val in row.values]
                    # Так как мы убрали +1, ИИ видит точный физический индекс строки в Excel (начиная с 1-й строки файла)
                    # Корректируем индекс для пользователя: idx + 1, чтобы совпадало со стандартной нумерацией строк Excel (1, 2, 3...)
                    full_text.append(f"Строка Excel {idx+1}: " + " | ".join(row_values))
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
    if not api_key.strip() or client is None:
        st.error("Ошибка: Пожалуйста, сначала введите действующий API-ключ Gemini в левой панели!")
    elif not user_query.strip():
        st.warning("Пожалуйста, введите текст запроса.")
    else:
        all_docs_context = ""
        if os.path.exists(DOCS_DIR):
            files = [f for f in os.listdir(DOCS_DIR) if f.endswith('.json')]
            for file in files:
                with open(os.path.join(DOCS_DIR, file), "r", encoding="utf-8") as f:
                    meta = json.load(f)
                    # Убрали искусственную нумерацию ДОКУМЕНТ №X, чтобы ИИ не путал ее с номерами страниц внутри файлов
                    all_docs_context += f"--- НАЧАЛО ДОКУМЕНТА: {meta['title']} ({meta['type']}) ---\n{meta['content']}\n--- КОНЕЦ ДОКУМЕНТА ---\n\n"
        
        if not all_docs_context:
            st.error("Ошибка: База знаний пуста! Пожалуйста, сначала загрузите файлы в левой панели.")
        else:
            with st.spinner("Gemini штудирует базу данных, мануалы и таблицы..."):
                try:
                    system_instruction = (
                        "You are an AI assistant for corporate knowledge base analysis. "
                        "Analyze the user's request using the provided context (memos, user manuals, Excel tables). "
                        "Provide answers in Russian language based strictly on the provided documents. "
                        "When referencing Excel tables, use the exact row numbers provided in the text (e.g., 'Строка Excel X'). "
                        "When referencing PDF files, use the exact page numbers tags provided (e.g., '[Страница X]'). "
                        "Do not add or subtract any numbers. Be extremely precise and concise."
                    )
                    
                    full_prompt = f"""Here is the corporate database:
{all_docs_context}

---
USER REQUEST:
"{user_query}"
---

Provide a detailed, helpful answer in RUSSIAN language based ONLY on the data above.
"""
                    full_prompt_clean = str(full_prompt).encode('utf-8', errors='ignore').decode('utf-8')
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=full_prompt_clean,
                        config={
                            "system_instruction": system_instruction, 
                            "temperature": 0.1  # Снизили температуру до 0.1 для максимальной точности в цифрах
                        }
                    )
                    
                    st.success("🤖 Ответ сформирован!")
                    st.subheader("🎯 Результат анализа:")
                    st.markdown(response.text)
                    
                except Exception as e:
                    error_msg = str(e)
                    if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                        st.error("⏳ Исчерпан дневной лимит бесплатных запросов для ЭТОГО ключа (20 запросов). Пожалуйста, вставьте СЛЕДУЮЩИЙ ключ слева!")
                    elif "401" in error_msg or "UNAUTHENTICATED" in error_msg:
                        st.error("❌ Ошибка авторизации: Этот API-ключ недействителен. Пожалуйста, перепроверьте его символы!")
                    elif "400" in error_msg or "API_KEY_INVALID" in error_msg or "expired" in error_msg.lower():
                        st.error("🔄 Срок действия этого API-ключа истек. Пожалуйста, сгенерируйте НОВЫЙ ключ в Google AI Studio и вставьте его слева!")
                    elif "503" in error_msg or "UNAVAILABLE" in error_msg:
                        st.error("🤖 Сервера Google сейчас сильно перегружены (ошибка 503). Подождите 1-2 минуты и нажмите кнопку анализа заново!")
                    else:
                        st.error(f"Произошла ошибка при вызове Gemini: {e}")
