# Excel Calc v1.0 🧮

A modern CustomTkinter-based GUI application designed to automate performance and salary calculations directly from active Excel WMS reports.

---

## 🇺🇸 English Description

### 🚀 Features
- **Smart Search:** Calculates individual earnings based on the employee's surname.
- **Automated Report Parsing:** Connects to active Excel instances and extracts data from 4 distinct types of logistics reports (RT operations, Inbound pallet splitting, Shipping/dispatch, and Internal transfers).
- **Data Visualization:** Generates interactive analytics pie charts using Matplotlib to display earnings breakdown.
- **Asynchronous UI:** Runs heavy Excel calculations in background threads via `threading` to prevent interface freezing.

### 🛠️ Tech Stack
- **GUI:** `CustomTkinter` (Modern dark-themed interface)
- **Excel Automation:** `xlwings` (Live interaction with open workbooks)
- **Data Analysis:** `pandas` (Fast data filtering and missing value handling)
- **Charts:** `matplotlib` (Embedded data visualization)

### 📦 Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Tos4-gr8/xl-report-parser
   ```

2. **Install the required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application:**
   ```bash
   python main.py
   ```

---

## 🇷🇺 Описание на русском

### 🚀 Функционал
- **Умный поиск:** Расчет индивидуального заработка по фамилии сотрудника.
- **Автоматический парсинг:** Подключение к открытым книгам Excel и сбор данных из 4 типов складских отчетов (Работа с РТ, Деление паллет на приемке, Экспедиция-отгрузка, Данные по перемещениям).
- **Визуализация данных:** Построение наглядных круговых диаграмм с помощью Matplotlib для отображения долей заработка с каждого отчета.
- **Асинхронный GUI:** Выполнение тяжелых расчетов в фоновом потоке (`threading`), благодаря чему окно программы не зависает в процессе работы.

### 🛠️ Технологический стек
- **Интерфейс:** `CustomTkinter`
- **Автоматизация Excel:** `xlwings`
- **Обработка данных:** `pandas`
- **Графики:** `matplotlib`

### 📦 Установка и запуск

1. **Клонируйте репозиторий:**
   ```bash
   git clone https://github.com/Tos4-gr8/xl-report-parser
   ```

2. **Установите необходимые зависимости:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Запустите приложение:**
   ```bash
   python main.py
   ```
