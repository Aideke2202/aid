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
        # Также собираем текст из таблиц, если они есть в инструкции
        for table in doc.tables:
            for row in table.rows:
                row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_text:
                    full_text.append(" | ".join(row_text))
        return "\n".join(full_text)
    except Exception as e:
        return f"Ошибка при чтении файла Word: {e}"

# --- БОКОВАЯ ПАНЕЛЬ: ЗАГРУЗКА ИНСТРУКЦИЙ ---
with st.sidebar:
    st.write("---")
    st.header("📥 Загрузка инструкций")
    
    # Выбор типа инструкции
    doc_type = st.selectbox("Тип документа:", ["Должностная инструкция", "Инструкция к программе / Ошибка"])
    doc_title = st.text_input("Название (ФИО сотрудника или имя программы):")
    
    # Загрузчик файлов Word
    uploaded_file = st.file_uploader("Выложите файл инструкции (.docx):", type=["docx"])
    
    save_btn = st.button("💾 Сохранить в базу знаний", type="secondary")
    
    if save_btn and doc_title and uploaded_file:
        file_bytes = uploaded_file.read()
        parsed_text = parse_docx(file_bytes)
        
        if "Ошибка при чтении" in parsed_text:
            st.error(parsed_text)
        else:
            # Сохраняем структурированный JSON
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

    # Список загруженного
    st.write("---")
    st.subheader("📚 Уже в базе знаний:")
    if os.path.exists(DOCS_DIR):
        files = [f for f in os.listdir(DOCS_DIR) if f.endswith('.json')]
        if files:
            for file in files:
                with open(os.path.join(DOCS_DIR, file), "r", encoding="utf-8") as f:
                    meta = json.load(f)
                    icon = "👤" if meta["type"] == "Должностная инструкция" else "💻"
                    st.write(f"{icon} **{meta['title']}** ({meta['type']})")
        else:
            st.info("База знаний пуста. Загрузите файлы .docx")

# --- ГЛАВНАЯ СТРАНИЦА: ПОИСК И АНАЛИЗ ---
st.subheader("📝 Опишите проблему, ошибку или входящую задачу")
user_query = st.text_area("Вставьте текст (например: 'Выскочила ошибка базы данных Lotus Notes...' или текстовку поручения):", height=150)

if st.button("🔍 Найти решение или ответственного", type="primary"):
    if not api_key or client is None:
        st.error("Ошибка: Не указан или неверен API-ключ Gemini!")
    elif not user_query.strip():
        st.warning("Пожалуйста, введите текст запроса или описание ошибки.")
    else:
        # Собираем контекст из всех файлов базы данных
        all_docs_context = ""
        if os.path.exists(DOCS_DIR):
            files = [f for f in os.listdir(DOCS_DIR) if f.endswith('.json')]
            for idx, file in enumerate(files, 1):
                with open(os.path.join(DOCS_DIR, file), "r", encoding="utf-8") as f:
                    meta = json.load(f)
                    all_docs_context += f"--- ДОКУМЕНТ №{idx} ---\nТип: {meta['type']}\nНазвание/Объект: {meta['title']}\nСодержимое инструкции:\n{meta['content']}\n\n"
        
        if not all_docs_context:
            st.error("Ошибка: База знаний пуста! Пожалуйста, сначала загрузите файлы инструкций в левой панели.")
        else:
            with st.spinner("Gemini изучает инструкции и ищет решение..."):
                try:
                    # Строгий системный промпт универсального диспетчера-технаря
                    system_instruction = (
                        "Ты — интеллектуальный корпоративный помощник, эксперт по должностным инструкциям "
                        "и технической поддержке программного обеспечения. Твоя задача — проанализировать запрос "
                        "пользователя (это может быть описание технической ошибки в программе или рабочее поручение) "
                        "и строго на основании предоставленных документов вынести вердикт.\n\n"
                        "1. Если пользователь описал ОШИБКУ/ПРОБЛЕМУ в программе: найди в инструкциях по программам "
                        "подходящий алгоритм решения, распиши пошагово, что делать, и укажи, к какому документу ты обращался.\n"
                        "2. If пользователь ввёл ПОРУЧЕНИЕ/ЗАДАЧУ: на основе должностных инструкций определи, "
                        "кто из сотрудников или какой отдел должен этим заниматься, аргументируя выводы обязанностями.\n"
                        "Если в базе нет ответа или подходящего сотрудника, вежливо сообщи об этом."
                    )
                    
                    full_prompt = f"""Вот полная база загруженных документов и инструкций вашего отдела:
{all_docs_context}

---
ЗАПРОС ПОЛЬЗОВАТЕЛЯ (Текстовка / Ошибка):
"{user_query}"
---

На основе документов выше дай точный технический ответ или назначь ответственного. Обоснуй решение ссылками на текст инструкций.
"""
                    
                    # Отправляем запрос в модель
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=full_prompt,
                        config={"system_instruction": system_instruction, "temperature": 0.2}
                    )
                    
                    st.success("🤖 Ответ сформирован!")
                    st.subheader("🎯 Результат анализа:")
                    st.markdown(response.text)
                    
                except Exception as e:
                    st.error(f"Произошла ошибка при вызове Gemini: {e}")
