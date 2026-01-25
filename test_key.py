import os
from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI

# 1. Загружаем переменные
print("--- ТЕСТ КЛЮЧА (GEMINI PRO) ---")
load_dotenv()

# 2. Проверяем ключ
api_key = os.getenv("GOOGLE_API_KEY")

if not api_key:
    print("❌ ОШИБКА: Ключ не найден в .env")
else:
    print(f"✅ Ключ найден: {api_key[:5]}...")
    
    # 3. Делаем запрос к СТАБИЛЬНОЙ модели gemini-pro
    try:
        # ВАЖНО: Тут теперь gemini-pro, а не flash
        llm = ChatGoogleGenerativeAI(
            model="gemini-pro", 
            google_api_key=api_key
        )
        res = llm.invoke("Напиши одно слово: Привет")
        print(f"✅ УСПЕХ! Ответ AI: {res.content}")
    except Exception as e:
        print(f"❌ Ошибка API: {e}")