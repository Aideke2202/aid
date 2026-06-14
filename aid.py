import streamlit as st
import os
import json
from google import genai
from docx import Document
import io

# Настройка страницы
st.set_page_config(page_title="База знаний и Диспетчер ИИ", page_icon="🧠", layout="wide")

st.title("🧠 Интеллектуальная база знаний отдела (Gemini)")
st.write("Загружайте должностные инструкции и мануалы к программам. ИИ поможет распределить задачи или найти решение ошибки.")

# Папки для хранения данных
BASE_DIR = "knowledge_base"
DOCS_DIR = os.path.join(BASE_DIR, "documents")
for folder in [BASE_DIR, DOCS_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# --- ИНИЦИАЛИЗАЦИЯ GEMINI ---
st.sidebar.header("🔑 Настройки ИИ")
api_key = st.sidebar.text_input("Введите Gemini API Key:", type="password", value=os.environ.get("GEMINI_API_KEY", ""))

client = None
if api_key:
    try:
        client = genai.Client(api_key=api_key)
    except Exception as e:
        st.sidebar.error(f"Ошибка инициализации ИИ: {e}")
else:
    st.sidebar.warning("🔑 Пожалуйста, введите API ключ для работы системы.")

# --- ФУНКЦИЯ ДЛЯ ЧТЕНИЯ WORD (.docx) ---
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

# --- БОКОВАЯ ПАНЕЛЬ: ЗАГРУЗКА И УДАЛЕНИЕ ---
with st.sidebar:
    st.write("---")
    st.header("📥 Загрузка инструкций")
    
    doc_type = st.selectbox("Тип документа:", ["Должностная инструкция", "Инструкция к программе / Ошибка"])
    doc_title = st.text_input("Название (ФИО или имя программы):")
    uploaded_file = st.file_uploader("Выложите файл инструкции (.docx):", type=["docx"])
    
    save_btn = st.button("💾 Сохранить в базу знаний", type="secondary")
    
    if save_btn and doc_title and uploaded_file:
        file_bytes = uploaded_file.read()
        parsed_text = parse_docx(file_bytes)
        
        if "Ошибка при чтении" in parsed_text:
            st.error(parsed_text)
        else:
            filename = f"{doc_type.replace(' ', '_
