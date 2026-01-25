import google.generativeai as genai
import os
from dotenv import load_dotenv

# Загружаем ключ
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

if not api_key:
    print("❌ Ключ не найден в .env")
    exit()

print(f"🔑 Используем ключ: {api_key[:5]}...")

# Настраиваем библиотеку
genai.configure(api_key=api_key)

print("⏳ Запрашиваю список доступных моделей у Google...")

try:
    models_found = False
    for m in genai.list_models():
        # Нам нужны только модели, которые умеют генерировать текст (generateContent)
        if 'generateContent' in m.supported_generation_methods:
            print(f"✅ Доступна модель: {m.name}")
            models_found = True
    
    if not models_found:
        print("⚠️ Список пуст! Google не вернул ни одной модели.")
        print("Причины: IP-адрес из РФ (нужен VPN) или неактивный проект в Google Cloud.")
        
except Exception as e:
    print(f"❌ Критическая ошибка соединения: {e}")
    if "400" in str(e) or "location" in str(e) or "404" in str(e):
        print("\n💡 СОВЕТ: Включите VPN (США/Европа) и попробуйте снова.")