import requests
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
import calendar
import numpy as np
import numpy as np
import pandas as pd
from pandas.tseries.offsets import DateOffset
import matplotlib.pyplot as plt

# === Патека до Excel фајлот ===
file_path = 'Indexi za trosoci na zivot.xlsx'

# === Читање на табелата ===
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()


df['Месец'] = pd.to_datetime(df['Месец'], dayfirst=True, errors='coerce')
df = df.sort_values('Месец').reset_index(drop=True)

# Заокружи ја колоната 'Индекс' на 2 децимали
df['Индекс'] = df['Индекс'].astype(float).round(2)
print(df['Индекс'].head(15))


# === Пресметка на месец за кој треба да се внесе податок ===
today = datetime.today()
target_month = today.replace(day=1) - timedelta(days=1)

# === Проверка дали веќе постои тој месец во табелата ===
if any((df['Месец'].dt.month == target_month.month) & (df['Месец'].dt.year == target_month.year)):
    print(f"✅ Податокот за {target_month.strftime('%B %Y')} веќе постои во табелата.")
else:
    month_code_map = {
        '2025-04': '258',
        # додади нови по потреба
    }
    month_key = target_month.strftime('%Y-%m')
    if month_key not in month_code_map:
        print(f"⚠️ Нема код за месецот {month_key}. Ажурирај ја мапата.")
    else:
        json_query = {
            "query": [
                {"code": "Месец", "selection": {"filter": "item", "values": [month_code_map[month_key]]}},
                {"code": "Базен период", "selection": {"filter": "item", "values": ["2"]}},
                {"code": "Главни COICOP групи", "selection": {"filter": "item", "values": ["0"]}}
            ],
            "response": {"format": "json"}
        }
        url = "https://makstat.stat.gov.mk:443/PXWeb/api/v1/mk/MakStat/Ceni/IndeksTrosZivot/TrosociZivot/120_CeniTr_Mk_IndTroZi_ml.px"
        response = requests.post(url, json=json_query)
        if response.status_code == 200:
            try:
                value = response.json()['data'][0]['values'][0]
                last_day = target_month.replace(day=calendar.monthrange(target_month.year, target_month.month)[1])
                new_row = pd.DataFrame([[last_day, round(float(value), 2)]], columns=['Месец', 'Индекс'])
                df = pd.concat([df, new_row], ignore_index=True)
                df['Индекс'] = df['Индекс'].astype(float).round(2)  # Ова гарантира дека новите податоци се заокружени
                print(f"✅ Успешно додаден индексот за {last_day.strftime('%d.%m.%Y')}: {value}")
            except (IndexError, KeyError):
                print("❌ Грешка при читање од JSON.")
        else:
            print("❌ API повик неуспешен. Код:", response.status_code)

df = df.sort_values('Месец').reset_index(drop=True)

# === ФУНКЦИИ ЗА ПРЕСМЕТКА ===
def days_between(d1, d2):
    return (d1 - d2).days

def calculate_excel_like_inflation(df, index, years):
    from pandas.tseries.offsets import DateOffset

    months_back = years * 12
    if index < months_back:
        return np.nan

    current_date = df.loc[index, 'Месец']
    try:
        # Земаме индексот за секои 12 месеци назад
        idxs = [index - i * 12 for i in range(years)]
        idxs.insert(0, index)  # додај го и тековниот месец
        indexes = df.loc[idxs, 'Индекс'].values

        if any(pd.isnull(indexes)):
            return np.nan

        product = np.prod(indexes) / (100 ** (years + 1))

        past_date = current_date - DateOffset(months=months_back)
        days_diff = (current_date - past_date).days

        if days_diff <= 0:
            return np.nan

        infl = (product ** (365 / days_diff) - 1) * 100
        return round(infl, 2)
    except Exception as e:
        print(f"Грешка за ред {index}: {e}")
        return np.nan


# === 1 ГОДИНА (се задржува оригиналната формула!) ===
infl1 = []
for i in range(len(df)):
    if i >= 12:
        now = df.at[i, 'Индекс']
        past = df.at[i - 12, 'Индекс']
        d1 = df.at[i, 'Месец']
        d2 = df.at[i - 12, 'Месец']
        days = days_between(d1, d2)
        try:
            val = ((now / 100) ** (365 / days) - 1) * 100
            infl1.append(round(val, 2))
        except:
            infl1.append(np.nan)
    else:
        infl1.append(np.nan)


# === ЗАПИШИ САМО КОРИГИРАНИ И ТОЧНИ КОЛОНИ ===
df['Инфлација Годишно (%)'] = infl1


# === Петгодишна инфлација (4-та колона) ===
df['Инфлација 5 Години (%)'] = [calculate_excel_like_inflation(df, i, 5) for i in range(len(df))]


# === Седумгодишна инфлација (5-та колона) ===
df['Инфлација 7 Години (%)'] = [calculate_excel_like_inflation(df, i, 7) for i in range(len(df))]

# === Форматирање и запис ===
df['Месец'] = df['Месец'].dt.strftime('%d.%m.%Y')
df = df[['Месец', 'Индекс', 'Инфлација Годишно (%)', 'Инфлација 5 Години (%)', 'Инфлација 7 Години (%)']]
df['Индекс'] = df['Индекс'].round(2)
df.to_excel(file_path, index=False, float_format="%.2f", sheet_name='baza_inflacija')

print("📁 Податоците се успешно запишани со ТОЧНИ инфлациски колони.")
# === ДОДАДЕНО: Полугодишен извештај и график ===
import matplotlib.pyplot as plt

# Врати 'Месец' назад во datetime облик (за групирање)
df['Месец'] = pd.to_datetime(df['Месец'], dayfirst=True)

# Групирање по година и полугодие
df['Година'] = df['Месец'].dt.year
df['Полугодие'] = df['Месец'].dt.month.apply(lambda x: 'H1' if x <= 6 else 'H2')

# Пресметка на просечен индекс по полугодие
semiannual_df = df.groupby(['Година', 'Полугодие'])['Индекс'].mean().reset_index()
semiannual_df['Индекс'] = semiannual_df['Индекс'].round(2)

# Запиши ја табелата во нов лист во Excel фајлот
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    semiannual_df.to_excel(writer, sheet_name='Полугодишен извештај', index=False)

print("📊 Полугодишната табела е додадена во Excel фајлот.")

# === Генерирање график ===
pivot_df = semiannual_df.pivot(index='Година', columns='Полугодие', values='Индекс')
pivot_df.plot(kind='bar', figsize=(10, 6), title='Полугодишен индекс на трошоци на живот по години')
plt.ylabel('Индекс')
plt.grid(True)
plt.tight_layout()
plt.savefig("Polugodishen_grafik.png")
plt.show()
print("📈 Графикот е снимен како Polugodishen_grafik.png")

