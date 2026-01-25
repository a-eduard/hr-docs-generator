import os
import json
import streamlit as st
try:
    from langchain_community.chat_models import ChatYandexGPT
    from langchain_core.prompts import PromptTemplate
except ImportError:
    pass 
from dotenv import load_dotenv

load_dotenv()

def get_llm(temp=0.1):
    api_key = os.getenv("YANDEX_API_KEY")
    folder_id = os.getenv("YANDEX_FOLDER_ID")

    if hasattr(st, "secrets"):
        api_key = st.secrets.get("YANDEX_API_KEY", api_key)
        folder_id = st.secrets.get("YANDEX_FOLDER_ID", folder_id)

    if not api_key or not folder_id:
        return None

    return ChatYandexGPT(
        api_key=api_key,
        folder_id=folder_id,
        model_uri=f"gpt://{folder_id}/yandexgpt/latest", 
        temperature=temp,
        max_tokens=7500
    )

def clean_json_response(content):
    content = content.strip()
    if "```" in content:
        content = content.replace("```json", "").replace("```", "")
    start_idx = content.find('{')
    end_idx = content.rfind('}')
    if start_idx != -1 and end_idx != -1:
        content = content[start_idx : end_idx + 1]
    return content

def generate_ai_duties(position: str) -> str:
    try:
        llm = get_llm(temp=0.6)
        if not llm: return ""
        template = """
        Ты — HR-директор. Напиши 5-7 обязанностей для должности: {position}.
        Стиль: Строгий, официальный.
        Формат: Только маркированный список.
        """
        chain = PromptTemplate(input_variables=["position"], template=template) | llm
        return chain.invoke({"position": position}).content
    except Exception as e:
        return f"Ошибка AI: {str(e)}"

def extract_data_from_egrul(text: str) -> dict:
    try:
        llm = get_llm(temp=0.1)
        if not llm: return None

        template = """
        Ты — алгоритм обработки ЕГРЮЛ. Извлеки данные и ОТФОРМАТИРУЙ их.
        
        ВХОДНОЙ ТЕКСТ:
        {text}
        
        ИНСТРУКЦИЯ (СТРОГО):
        1. Преобразуй ВЕСЬ текст из CAPS LOCK в обычный (Title Case).
        2. Исключения (оставь большими): ООО, АО, ИНН, КПП, ОГРН.
        
        ВЕРНИ ТОЛЬКО JSON с ключами:
        - "opf": ОПФ (Например: "Общество с ограниченной ответственностью")
        - "name": Название без кавычек (Например: Альянс)
        - "short_name": Сокращенное наименование (Например: ООО "Альянс")
        - "inn": ИНН
        - "kpp": КПП (9 цифр)
        - "ogrn": ОГРН
        - "address": Адрес
        - "boss_name": ФИО директора
        - "boss_pos": Должность
        """
        
        prompt = PromptTemplate(input_variables=["text"], template=template)
        chain = prompt | llm
        
        response = chain.invoke({"text": text[:30000]})
        return json.loads(clean_json_response(response.content))

    except Exception as e:
        print(f"Extraction Error: {e}")
        return None