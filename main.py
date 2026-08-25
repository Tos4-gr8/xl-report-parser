import xlwings as xw
import customtkinter as ctk
import pandas as pd
import threading
import matplotlib.pyplot as plt
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# Creating the window visual
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Create the main window
app = ctk.CTk()
app.geometry("510x430")
app.title("Excel calc v1.0")

# Icon
current_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(current_dir, "icon.ico")
app.iconbitmap(icon_path)


# ==============================================================================
# Backend Logic
# ==============================================================================

# Check for empty cells
def value_is_not_none(value):
    if value is None:
        return 0
    return value

# Button trigger to start the calculator
def run_calculations():
    # Connect to the active Excel file
    try:
        wb = xw.books.active
    except Exception:
        label.configure(text="Excel файл не найден.\n\n")
        return

    # Report statuses
    report_status_1 = None     # Report: RT operations
    report_status_2 = None     # Report: Inbound pallet splitting
    report_status_3 = None     # Report: Shipping and dispatch
    report_status_4 = None     # Report: Transfer data

    # Check open reports
    for excel_app in xw.apps:
        for book in excel_app.books:
            try:
                sheet = book.sheets.active
        
                report_book_1 = str(sheet.range('A1').value or '').strip().lower()
                report_book_2 = str(sheet.range('C1').value or '').strip().lower()
                report_book_3 = str(sheet.range('A1').value or '').strip().lower()
                report_book_4 = str(sheet.range('H1').value or '').strip().lower()

                # Search for reports by specific phrases
                if "производительность приемщиков и переместителей с рт" in report_book_1:
                    report_status_1 = sheet

                if "загрузка приёмщиков за период" in report_book_2:
                    report_status_2 = sheet

                if "отгрузка" in report_book_3:
                    report_status_3 = sheet
            
                if "участок сборки" in report_book_4:
                    report_status_4 = sheet
            except Exception as e:
                print(f"Skipped hidden or system Excel file. Error: {e}")
                continue

    # Search target - Surname
    target_surname = input_line.get().strip().title()
    user_full_name = None

    # Error prevention - set default values to 0 in case open reports are missing
    total_result_report_1 = 0
    total_result_report_2 = 0
    total_result_report_3 = 0
    total_result_report_4 = 0

    if target_surname.strip() == "":
        label.configure(text=f"Введите фамилию в строку ввода.\n\n")
    else:
        # Report: RT operations
        if report_status_1:
            label_report_1.configure(text=f"Отчет: Работа с РТ - ✅")
            df_report_1 = report_status_1.used_range.options(pd.DataFrame, index=False, header=False).value

            df_report_1.rename(columns={
                0: "ФИО",
                10: "Позиций собрано (Перемещения)",
                11: "Коробок собрано (Перемещения)",
                14: "Позиций размещено (Розница)",
                15: "Коробок размещено (Розница)",
                18: "Позиций собрано (ОПТ сборка)",
                19: "Коробок собрано (ОПТ сборка)",
                22: "Внутренние перемещения",
            }, inplace=True)

            surname_report_1 = df_report_1["ФИО"].astype(str).str.strip().str.lower().str.startswith(target_surname.strip().lower())
            found_rows = df_report_1[surname_report_1]

            if not found_rows.empty:
                row = found_rows.iloc[0]
                if user_full_name is None:
                    user_full_name = str(row["ФИО"]).strip()

                def safe_val(val):
                    if pd.isna(val) or val is None or str(val).strip() == "":
                        return 0
                    return float(val)
                
                # Multiply data for Report: RT operations
                value_k = safe_val(row["Позиций собрано (Перемещения)"]) * 5.30
                value_l = safe_val(row["Коробок собрано (Перемещения)"]) * 2
                value_o = safe_val(row["Позиций размещено (Розница)"]) * 6.80
                value_p = safe_val(row["Коробок размещено (Розница)"]) * 3.80
                value_s = safe_val(row["Позиций собрано (ОПТ сборка)"]) * 4
                value_t = safe_val(row["Коробок собрано (ОПТ сборка)"]) * 2.20
                value_w = safe_val(row["Внутренние перемещения"]) * 18

                total_result_report_1 = value_k + value_l + value_o + value_p + value_s + value_t + value_w                        

        # Report: Inbound pallet splitting
        if report_status_2:
            label_report_2.configure(text=f"Отчет: Деление паллет приемка - ✅")
            report_data_2 = report_status_2.used_range.value

            df_report_2 = pd.DataFrame(data=report_data_2[3:], columns=report_data_2[2])
            df_report_2 = df_report_2.fillna(0)    # Fill missing values with 0

            surname_report_2 = df_report_2[df_report_2["ФИО"].astype(str).str.startswith(target_surname)]
            if not surname_report_2.empty:
                user_full_name = surname_report_2["ФИО"].iloc[0]

            wholesale = surname_report_2[surname_report_2["Зона размещения"].astype(str).str.contains("опт", case=False)]
            goods_receipt = surname_report_2[surname_report_2["Зона размещения"].astype(str).str.contains("0", case=False)]
            retail = surname_report_2[ ~ surname_report_2["Зона размещения"].astype(str).str.contains("0|опт", case=False)]
            
            ssci_click = surname_report_2["В т.ч. отсканировано КИЗ"].sum()
            wholesale_box = wholesale["Кол-во максималок"].sum()
            wholesale_position = wholesale["Количество позиций, шт"].sum()
            retail_box = retail["Кол-во максималок"].sum()
            retail_position = retail["Количество позиций, шт"].sum()
            goods_receipt_box = goods_receipt["Кол-во максималок"].sum()
            goods_receipt_position = goods_receipt["Количество позиций, шт"].sum()

            #Multiply data for Report: Inbound pallet splitting
            total_result_report_2 = wholesale_box * 1.50 + retail_box * 3.80 + goods_receipt_box * 2.25 + ssci_click * 1.20 + wholesale_position * 7 + retail_position * 6.80 + goods_receipt_position * 1.90

        # Report: Shipping and dispatch
        if report_status_3:
            label_report_3.configure(text=f"Отчет: Экспедиция-отгрузка - ✅")
            report_data_3 = report_status_3.used_range.value

            df_report_3 = pd.DataFrame(data=report_data_3[1:], columns=report_data_3[0])
            df_report_3 = df_report_3.fillna(0)    # Fill missing values with 0

            surname_report_3 = df_report_3[df_report_3["Экспедитор"].astype(str).str.startswith(target_surname)]
            if not surname_report_3.empty:
                user_full_name = surname_report_3["Экспедитор"].iloc[0]

            #Multiply data for Report: Shipping and dispatch
            total_result_report_3 = int(surname_report_3.shape[0] * 4.90)

        # Report: Transfer data
        if report_status_4:
            label_report_4.configure(text=f"Отчет: Данные по перемещениям - ✅")
            report_data_4 = report_status_4.used_range.value

            df_report_4 = pd.DataFrame(data=report_data_4[1:], columns=report_data_4[0])
            df_report_4 = df_report_4.fillna(0)    # Fill missing values with 0

            surname_report_4 = df_report_4[df_report_4["Разместитель"].astype(str).str.startswith(target_surname)]
            if not surname_report_4.empty:
                user_full_name = surname_report_4["Разместитель"].iloc[0]

            type_operation = surname_report_4[surname_report_4["Тип"].astype(str).str.contains("Зачистка комнаты", case=False)]
            operation_sector_retail = type_operation[type_operation["Участок размещения"].astype(str).str.contains("розничный участок сборки")]
            total_operation_sector_retail = operation_sector_retail["Позиций"].sum()
            result_retail = int(total_operation_sector_retail * 55)
            operation_sector_opt = type_operation[ ~ type_operation["Участок размещения"].str.contains("К- розничный участок сборки")]
            total_opt = operation_sector_opt[["Позиций", "Максималок"]].sum()
            result_position_opt = total_opt["Позиций"]
            result_box_opt = total_opt["Максималок"]
            result = int(result_position_opt * 12.1 + result_box_opt * 5.80)

            #Multiply data for Report: Transfer data
            total_result_report_4 = result_retail + result

        # Final output
        if user_full_name is None:
            label.configure(text=f"Сотрудника с фамилией '{target_surname}' не найдено.\nПроверьте корректность введенных данных\nили правильность открытых отчетов.")
        else:
            user_first_name = user_full_name.split()[1]
            total_all_report_dirt = int(total_result_report_1 + total_result_report_2 + total_result_report_3 + total_result_report_4)
            total_all_report_clean = int(total_all_report_dirt - (total_all_report_dirt / 100 * 13))
            app.after(0, update_ui, user_first_name, total_all_report_dirt, total_all_report_clean, total_result_report_1, total_result_report_2, total_result_report_3, total_result_report_4)       

# Create pie chart
def create_pie_chart(tab_name, sum_1, sum_2, sum_3, sum_4):
    for widget in tab_name.winfo_children():
        widget.destroy()    # Clear the "Статистика" tab

    labels = []     # Chart slice labels
    sizes = []      # Earnings per report

    if sum_1 > 0:
        labels.append("Работа с РТ")
        sizes.append(sum_1)
    if sum_2 > 0:
        labels.append("Приемка")
        sizes.append(sum_2)
    if sum_3 > 0:
        labels.append("Пики")
        sizes.append(sum_3)
    if sum_4 > 0:
        labels.append("Зачистки")
        sizes.append(sum_4)
    if not sizes:
        no_data_label = ctk.CTkLabel(tab_name, text="Нет данных для диаграммы", font=("Arial", 16))
        no_data_label.pack(pady=50)
        return

    # Render the chart
    plt.style.use("dark_background")
    fig, ax = plt.subplots(figsize=(5, 4), dpi=100)
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"]

    ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=140, colors=colors[:len(sizes)], radius=0.9, textprops={"fontsize": 11})
    ax.axis("equal")

    fig.patch.set_facecolor("#2b2b2b")

    canvas = FigureCanvasTkAgg(fig, master=tab_name)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10, padx=10, fill="both", expand=True)        

# Update UI
def update_ui(first_name, dirt, clean, sum_1, sum_2, sum_3, sum_4):
    label.configure(text=f"{first_name}, ваш результат:\nГрязный заработок: {dirt}₽\nЧистый заработок: {clean}₽")    # Update text
    create_pie_chart(tabview.tab("Статистика"), sum_1, sum_2, sum_3, sum_4)                                          # Render chart

# Handle button click
def button_start_click(event=None):
    thread = threading.Thread(target=run_calculations, daemon=True)
    thread.start()

# Bind Enter key to trigger calculation
app.bind("<Return>", button_start_click)


# ==============================================================================
# Frontend Logic
# ==============================================================================

# Tab View Component
tabview = ctk.CTkTabview(app, width=380, height=450)
tabview.pack(pady=10, padx=10)

# Add tabs
tabview.add("Калькулятор")
tabview.add("Статистика")
tabview.add("Инструкция")

# Output label for results
label = ctk.CTkLabel(tabview.tab("Калькулятор"), text="Добро пожаловать.\n\n", font=("Arial", 17))
label.pack(pady=20)

# Input field
input_line = ctk.CTkEntry(tabview.tab("Калькулятор"), placeholder_text="Введите вашу фамилию", width=200)
input_line.pack(pady=10)

# Calculation trigger button
button_start = ctk.CTkButton(tabview.tab("Калькулятор"), text="Рассчитать З/П", command=button_start_click)
button_start.pack(pady=10)

# ==============================================================================
# Report Status Bars
# ==============================================================================

# Container frame for status layout
reports_frame = ctk.CTkFrame(tabview.tab("Калькулятор"), fg_color="transparent")
reports_frame.pack(pady=10)

# Report: RT operations
label_report_1 = ctk.CTkLabel(reports_frame, text="Отчет: Работа с РТ - ❌", font=("Arial", 15), anchor="w")
label_report_1.pack(pady=3, fill="x", anchor="w")

# Report: Inbound pallet splitting
label_report_2 = ctk.CTkLabel(reports_frame, text="Отчет: Деление паллет приемка - ❌", font=("Arial", 15), anchor="w")
label_report_2.pack(pady=3, fill="x", anchor="w")

# Report: Shipping and dispatch
label_report_3 = ctk.CTkLabel(reports_frame, text="Отчет: Экспедиция-отгрузка - ❌", font=("Arial", 15), anchor="w")
label_report_3.pack(pady=3, fill="x", anchor="w")

# Report: Transfer data
label_report_4 = ctk.CTkLabel(reports_frame, text="Отчет: Данные по перемещениям - ❌", font=("Arial", 15), anchor="w")
label_report_4.pack(pady=3, fill="x", anchor="w")

# ==============================================================================
# Instructions Tab
# ==============================================================================

label_instructions = ctk.CTkLabel(tabview.tab("Инструкция"),
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
                                    anchor="e")
label_instructions.pack(pady=15)

# ==============================================================================
# Application Main Loop
# ==============================================================================

app.mainloop()