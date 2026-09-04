import threading
import customtkinter as ctk
import xlwings as xw

import config
from backend.excel_worker import find_active_reports
from backend.calculators import (
    calculate_rt_operations,
    calculate_pallet_splitting,
    calculate_shipping,
    calculate_transfer,
    calculate_clean_salary
)
from gui.chart_view import create_pie_chart


class ExcelCalcApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # Configure window settings from the configuration file
        ctk.set_appearance_mode(config.APPEARANCE_MODE)
        ctk.set_default_color_theme(config.DEFAULT_COLOR_THEME)
        
        self.geometry(config.WINDOW_GEOMETRY)
        self.title(config.WINDOW_TITLE)
        
        try:
            self.iconbitmap(config.ICON_PATH)
        except Exception as e:
            print(f"Failed to load application icon: {e}")

        # Initialize UI elements
        self._create_widgets()
        
        # Bind the Enter key to trigger calculations
        self.bind("<Return>", self._button_start_click)

    def _create_widgets(self):
        """Creates and positions all widgets on the window form."""
        # Main Tabview component
        self.tabview = ctk.CTkTabview(self, width=380, height=450)
        self.tabview.pack(pady=10, padx=10)

        self.tabview.add("Калькулятор")
        self.tabview.add("Статистика")
        self.tabview.add("Инструкция")

        # --- "Калькулятор" Tab ---
        self.label_status = ctk.CTkLabel(
            self.tabview.tab("Калькулятор"), 
            text="Добро пожаловать.\n\n", 
            font=("Arial", 17)
        )
        self.label_status.pack(pady=20)

        self.input_line = ctk.CTkEntry(
            self.tabview.tab("Калькулятор"), 
            placeholder_text="Введите вашу фамилию", 
            width=200
        )
        self.input_line.pack(pady=10)

        self.button_start = ctk.CTkButton(
            self.tabview.tab("Калькулятор"), 
            text="Рассчитать З/П", 
            command=self._button_start_click
        )
        self.button_start.pack(pady=10)

        # Status tracking frame for reports
        self.reports_frame = ctk.CTkFrame(self.tabview.tab("Калькулятор"), fg_color="transparent")
        self.reports_frame.pack(pady=10)

        self.label_report_1 = ctk.CTkLabel(self.reports_frame, text="Отчет: Работа с РТ - ❌", font=("Arial", 15), anchor="w")
        self.label_report_1.pack(pady=3, fill="x", anchor="w")

        self.label_report_2 = ctk.CTkLabel(self.reports_frame, text="Отчет: Деление паллет приемка - ❌", font=("Arial", 15), anchor="w")
        self.label_report_2.pack(pady=3, fill="x", anchor="w")

        self.label_report_3 = ctk.CTkLabel(self.reports_frame, text="Отчет: Экспедиция-отгрузка - ❌", font=("Arial", 15), anchor="w")
        self.label_report_3.pack(pady=3, fill="x", anchor="w")

        self.label_report_4 = ctk.CTkLabel(self.reports_frame, text="Отчет: Данные по перемещениям - ❌", font=("Arial", 15), anchor="w")
        self.label_report_4.pack(pady=3, fill="x", anchor="w")

        # --- "Инструкция" Tab ---
        self.label_instructions = ctk.CTkLabel(
            self.tabview.tab("Инструкция"),
            text="1. Сформируйте нужные отчеты:\n\n"
                 "   BE -> Склад -> Отчеты:\n"
                 "   •Отчет по производительности приемщиков и переместителей с РТ\n"
                 "   •Загруженность приемщиков при делении паллет\n"
                 "   •Данные по перемещениям\n\n"
                 "   BE -> Склад -> Экспедиция-отгрузка -> Формирование отчета:\n"
                 "   •Мск.Лог проверки за период\n\n"
                 "2. Введите вашу фамилию.\n"
                 "3. Нажмите кнопку «Рассчитать З/П»",                                         
            font=("Arial", 14),
            justify="left",
            anchor="e"
        )
        self.label_instructions.pack(pady=15)

    def _button_start_click(self, event=None):
        """Triggers calculation logic inside a separate background thread."""
        thread = threading.Thread(target=self._run_calculations, daemon=True)
        thread.start()

    def _run_calculations(self):
        """Executes core calculation logic inside a background worker thread."""
        # Reset interface status indicators safely using the main thread queue
        self.after(0, lambda: self.label_report_1.configure(text="Отчет: Работа с РТ - ❌"))
        self.after(0, lambda: self.label_report_2.configure(text="Отчет: Деление паллет приемка - ❌"))
        self.after(0, lambda: self.label_report_3.configure(text="Отчет: Экспедиция-отгрузка - ❌"))
        self.after(0, lambda: self.label_report_4.configure(text="Отчет: Данные по перемещениям - ❌"))

        # Check for active Excel connectivity
        try:
            _ = xw.books.active
        except Exception:
            self.after(0, lambda: self.label_status.configure(text="Excel файл не найден.\n\n"))
            return

        target_surname = self.input_line.get().strip().title()
        if not target_surname:
            self.after(0, lambda: self.label_status.configure(text="Введите фамилию в строку ввода.\n\n"))
            return

        # Scan active workbook scopes
        active_sheets = find_active_reports()

        user_full_name = None
        sum_1, sum_2, sum_3, sum_4 = 0.0, 0.0, 0.0, 0.0

        # 1. RT Operations Report
        if active_sheets["rt"]:
            self.after(0, lambda: self.label_report_1.configure(text="Отчет: Работа с РТ - ✅"))
            sum_1, name = calculate_rt_operations(active_sheets["rt"], target_surname)
            if name: user_full_name = name

        # 2. Pallet Splitting Report
        if active_sheets["pallet"]:
            self.after(0, lambda: self.label_report_2.configure(text="Отчет: Деление паллет приемка - ✅"))
            sum_2, name = calculate_pallet_splitting(active_sheets["pallet"], target_surname)
            if name: user_full_name = name

        # 3. Shipping Report
        if active_sheets["shipping"]:
            self.after(0, lambda: self.label_report_3.configure(text="Отчет: Экспедиция-отгрузка - ✅"))
            sum_3, name = calculate_shipping(active_sheets["shipping"], target_surname)
            if name: user_full_name = name

        # 4. Inventory Transfers Report
        if active_sheets["transfer"]:
            self.after(0, lambda: self.label_report_4.configure(text="Отчет: Данные по перемещениям - ✅"))
            sum_4, name = calculate_transfer(active_sheets["transfer"], target_surname)
            if name: user_full_name = name

        # Verify whether the target user was discovered across files
        if user_full_name is None:
            self.after(0, lambda: self.label_status.configure(
                text=f"Сотрудника с фамилией '{target_surname}' не найдено.\n"
                     f"Проверьте корректность введенных данных\n"
                     f"или правильность открытых отчетов."
            ))
        else:
            # Aggregate calculations
            total_dirty = int(sum_1 + sum_2 + sum_3 + sum_4)
            total_clean = calculate_clean_salary(total_dirty)

            # Safely delegate UI state mutation back to the main rendering loop
            self.after(0, self._update_ui, user_full_name, total_dirty, total_clean, sum_1, sum_2, sum_3, sum_4)

    def _update_ui(self, full_name, dirt, clean, sum_1, sum_2, sum_3, sum_4):
        """Updates text metrics on screen and re-renders the distribution chart."""
        name_parts = full_name.split()
        
        # Extract the first name if standard 'Last First Middle' formatting matches
        if len(name_parts) >= 2:
            display_name = name_parts[1]
        else:
            display_name = full_name
            
        self.label_status.configure(
            text=f"{display_name}, ваш результат:\n"
                 f"Грязный заработок: {dirt}₽\n"
                 f"Чистый заработок: {clean}₽"
        )
        
        # Build the chart on the statistics tab canvas
        create_pie_chart(self.tabview.tab("Статистика"), sum_1, sum_2, sum_3, sum_4)
