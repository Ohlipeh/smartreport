# SmartReport – Python Excel Automation Tool

SmartReport is a Python automation solution that processes raw sales data (CSV) and generates structured, formatted Excel reports with business insights.

This project demonstrates practical automation skills using Python, showcasing data processing, calculations, and Excel report generation.

---

## 🚀 Key Features

- Reads sales data from CSV
- Calculates total and average revenue
- Identifies top-selling products
- Groups revenue by category
- Generates a polished multi-sheet Excel report

---

## 📁 Project Structure

smartreport/
│
├── data/
├── output/
├── src/
├── requirements.txt
└── README.md

---

## 🛠 Technologies Used

- Python
- pandas
- openpyxl

---

## ▶ How to Run

1. Clone the repo  
   `git clone https://github.com/Ohlipeh/smartreport.git`

2. Create Python virtual environment  
   `python -m venv venv`

3. Activate environment  
   `venv\Scripts\activate`

4. Install dependencies  
   `pip install -r requirements.txt`

5. Run the automation script  
   `python src/main.py`

---

## 📌 Output

The project generates a formatted Excel file inside the `output` folder with:

- **Dados Brutos** — raw sales data
- **Resumo por Categoria** — category grouping
- **Resumo Executivo** — key performance information

---

## 🎯 Purpose

This project is built as a portfolio demonstration of Python automation skills focusing on **data processing and Excel reporting** — a common and valuable demand in business automation.

---

## 🧠 Next Improvements

Future versions may include:

- Support for dynamic file input
- Interactive UI (e.g., Streamlit)
- Scheduled automation (cron/jobs)
- Email or Slack delivery of reports
