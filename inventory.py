import pandas as pd
import matplotlib.pyplot as plt
import requests
import os
from dotenv import load_dotenv
import google.generativeai as genai

# Загрузка настроек
load_dotenv()
API_KEY = os.getenv("YOUR_API_KEY")
WEBHOOK_URL = os.getenv("N8N_WEBHOOK")

# Проверка для Машеньки (печать в терминал)
print(f"--- Отладка ---")
print(f"Файл .env найден: {os.path.exists('.env')}")
print(f"Ключ загружен: {'Да' if API_KEY else 'НЕТ'}")
print(f"----------------")

# Настройка Gemini (упрощенная, чтобы не было ошибок Metadata)
# Настройка Gemini (выровнено для Машеньки)
try:
    if API_KEY:
        print(f"✅ Ключ найден (первые символы): {API_KEY[:5]}...")
        genai.configure(api_key=API_KEY)
        model = genai.GenerativeModel("gemini-2.5-flash")
        print("✅ Подключение к AI настроено")
    else:
        print("⚠️ API ключ не заполнен в файле .env")
        model = None # Помечаем, что модели нет, чтобы не было ошибки дальше
except Exception as e:
    print(f"Ошибка настройки AI: {e}")
    model = None

# Загрузка данных
try:
    df = pd.read_excel("warehouse_data.xlsx")
    df["Количество"] = df["Количество"].fillna(0)

    # Расчёты
    kurs = 45
    tax_rate = 0.20
    df["Цена_UAH"] = df["Цена_EUR"] * kurs
    df["Сумма_UAH"] = df["Количество"] * df["Цена_UAH"]
    df["Налог_UAH"] = df["Сумма_UAH"] * tax_rate

    summary = df.groupby("Товар")["Количество"].sum().reset_index()
    print("✅ Данные обработаны")

    # AI анализ
    try:
        inventory_text = summary.to_string(index=False)
        prompt = f"Ты аналитик склада. Проанализируй остатки: {inventory_text}. Дай короткую рекомендацию по закупке на русском."
        
        response = model.generate_content(prompt, request_options={"timeout": 15})
        ai_advice = response.text
        print("✅ Анализ AI готов")
    except Exception as e:
        print(f"⚠️ Gemini пропущен: {e}")

    # Создание отчета Excel
    print("📂 Создаем Excel отчет...")
    with pd.ExcelWriter("Professional_Report.xlsx", engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Warehouse", index=False)
        summary.to_excel(writer, sheet_name="Analytics", index=False)
        sheet = writer.sheets["Analytics"]
        sheet.write("D1", "Рекомендация AI")
        sheet.write("D2", ai_advice)

    # График (важно: закрываем его!)
    summary.plot(kind="bar", x="Товар", y="Количество", legend=False)
    plt.title("Остатки товаров на складе")
    plt.tight_layout()
    plt.savefig("inventory_chart.png")
    plt.close() # <--- ВОТ ЭТО НЕ ДАСТ СКРИПТУ ЗАВИСНУТЬ!
    print("✅ График сохранен")

    # Отправка в n8n
    if WEBHOOK_URL:
        try:
            data = summary.to_dict(orient="records")
            # Добавим итоговую сумму для нашего письма в n8n
            final_data = {
                "items": data,
                "total_sum": float(df["Сумма_UAH"].sum()),
                "ai_recommendation": ai_advice
            }
            requests.post(WEBHOOK_URL, json=final_data, timeout=10)
            print("🚀 Данные успешно улетели в n8n!")
        except:
            print("⚠️ n8n недоступен, но отчет сохранен локально")

    print("⭐️ Скрипт успешно завершен!")

except Exception as e:
    print(f"❌ Критическая ошибка: {e}")