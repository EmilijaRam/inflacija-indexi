# 📈 Inflation Tracker – Cost of Living Indices

Python script that automatically downloads monthly cost of living indices, stores them in Excel, and calculates 1-year, 5-year, and 7-year inflation rates with correction by days and compound index.

---

## ✨ Features

- ✅ Automatic download of monthly consumer price indices (CPI)
- 🧮 Inflation calculations for 1, 5, and 7 years
- 📅 Half-year data organization (H1 and H2)
- 📊 Excel output with structured sheets and data updates
- 📈 Optional visualization (line chart for inflation trend)

---

## 🛠️ Technologies

- Python 3.x  
- pandas, openpyxl  
- Excel formulas  
- matplotlib (optional)

---

## 🚀 Installation

git clone https://github.com/EmilijaRam/inflacija-indexi.git
cd inflacija-indexi
pip install -r requirements.txt
---

## ▶️ Usage
python main.py
---

## 📸 Example Output
🧾 Excel Output Structure
   - Sheet: Indeksi – raw monthly CPI data
   - Sheet: Inflacija – inflation rates with formulas
   - Sheet: Polugodista – summary by half-years
![CoreInflation](https://github.com/EmilijaRam/inflacija-indexi/blob/main/CoreInflation.png)
![Half-yearly report on cost of living](https://github.com/EmilijaRam/inflacija-indexi/blob/main/Half-yearly%20report%20on%20cost%20of%20living.png)
![Half-yearly report on inflation](https://github.com/EmilijaRam/inflacija-indexi/blob/main/Half-yearly%20report%20on%20inflation.png)

