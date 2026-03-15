import os
import pandas as pd
import requests
import matplotlib.pyplot as plt
import google.generativeai as genai
from datetime import datetime, timedelta
from dotenv import load_dotenv

# --- Загружаем ключи ---
load_dotenv()
genai.configure(api_key=os.getenv("YOUR_API_KEY"))
webhook_url = os.getenv("N8N_WEBHOOK")

# --- Загружаем данные ---
df = pd.read_excel('warehouse_data.xlsx')
df['Дата_Поставки'] = pd.to_datetime(df['Дата_Поставки'])

# --- Фильтр по периоду ---
choice = input("1 - Все время, 2 - Последняя неделя: ")
if choice == '2':
    week_ago = datetime.now() - timedelta(days=7)
    df = df[df['Дата_Поставки'] >= week_ago]
    print("--- Отчет за неделю ---")
else:
    print("--- Отчет за все время ---")

# --- Расчеты ---
df['Количество'] = df['Количество'].fillna(0)
kurs = 45
tax_rate = 0.20
df['Цена_UAH'] = df['Цена_EUR'] * kurs
df['Сумма_UAH'] = (df['Количество'] * df['Цена_UAH']).round(2)
df['Налог_UAH'] = df['Сумма_UAH'] * tax_rate

summary = df.groupby('Товар')['Количество'].sum().reset_index()

# --- AI-Анализ ---
ai_advice = "AI сейчас отдыхает 😴"
try:
    inventory_text = summary.to_string(index=False)
    prompt = f"Ты аналитик склада. Проанализируй остатки:\n{inventory_text}\nДай короткий совет."
    response = genai.GenerativeModel('gemini-3-flash-latest').generate_content(prompt, request_options={"timeout": 20})
    ai_advice = response.text
except Exception as e:
    print(f"AI не ответил: {e}")

# --- Сохраняем Excel ---
with pd.ExcelWriter('Professional_Report.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Весь Склад', index=False)
    summary.to_excel(writer, sheet_name='Аналитика', index=False)
    
    workbook = writer.book
    sheet1 = writer.sheets['Весь Склад']
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
    money_fmt = workbook.add_format({'num_format': '#,##0.00 "₴"'})
    total_fmt = workbook.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1, 'num_format': '#,##0.00 "₴"'})

    for i, col in enumerate(df.columns):
        data_len = df[col].astype(str).str.len().max()
        header_len = len(str(col))
        width = max(data_len, header_len, 15)
        if "Дата" in col: width = 20
        elif "UAH" in col or "Цена" in col: width = 16
        elif "Артикул" in col: width = 10
        elif "Количество" in col: width = 11
        sheet1.set_column(i, i, width)
        sheet1.write(0, i, col, header_fmt)

    last_row = len(df) + 1
    sheet1.write(last_row, 6, 'ОБЩИЙ ИТОГ:', header_fmt)
    sheet1.write(last_row, 7, df['Сумма_UAH'].sum(), total_fmt)

    sheet2 = writer.sheets['Аналитика']
    sheet2.set_column('A:B', 20)
    sheet2.write('D1', 'СОВЕТ ОТ GEMINI 3:', header_fmt)
    sheet2.write('D2', ai_advice)

# --- Отправка в n8n ---
df_to_send = df.head(100).copy()
df_to_send = df_to_send.fillna(0)
for col in ['Артикул','Товар','Дата_Поставки']:
    if col in df_to_send.columns:
        df_to_send[col] = df_to_send[col].astype(str)

try:
    response = requests.post(webhook_url, json=df_to_send.to_dict(orient='records'))
    print(f"Ответ от n8n: {response.status_code}")
except Exception as e:
    print(f"Ошибка отправки: {e}")