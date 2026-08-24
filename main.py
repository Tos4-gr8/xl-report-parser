import xlwings as xw
import customtkinter as ctk
import pandas as pd
import threading
import matplotlib.pyplot as plt
import os
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


#Базовый визуал окна
ctk.set_appearance_mode("dark")         #Темная тема
ctk.set_default_color_theme("blue")     #Цвет кнопок (синий)

#Главное окно
app = ctk.CTk()
app.geometry("510x430")
app.title("Excel calc v1.0")

#Иконка
current_dir = os.path.dirname(os.path.abspath(__file__))
icon_path = os.path.join(current_dir, "icon.ico")
app.iconbitmap(icon_path)


###Backend логика
#Проверка на пустые ячейки
def value_is_not_none(value):
    if value is None:
        return 0
    return value

#Кнопка запуска калькулятора
def run_calculations():
    #Подключение к активному файлу Excel
    try:
        wb = xw.books.active
    except:
        label.configure(text=f"Excel файл не найден.\n\n")
        return

    #Статусы отчетов
    report_status_1 = None     #Отчет работы с РТ
    report_status_2 = None     #Отчет при делении паллетов
    report_status_3 = None     #Отчет пики экспедии
    report_status_4 = None     #Отчет зачистки

    #Проверям открытые отчеты
    for excel_app in xw.apps:
        for book in excel_app.books:
            try:
                sheet = book.sheets.active
        
                report_book_1 = str(sheet.range('A1').value or '').strip().lower()
                report_book_2 = str(sheet.range('C1').value or '').strip().lower()
                report_book_3 = str(sheet.range('A1').value or '').strip().lower()
                report_book_4 = str(sheet.range('H1').value or '').strip().lower()
                #Поиск отчетов по заданной фразе
                if "производительность приемщиков и переместителей с рт" in report_book_1:
                    report_status_1 = sheet

                if "загрузка приёмщиков за период" in report_book_2:
                    report_status_2 = sheet

                if "отгрузка" in report_book_3:
                    report_status_3 = sheet
            
                if "участок сборки" in report_book_4:
                    report_status_4 = sheet
            except Exception as e:
                print(f"Пропущен скрытый или системный файл Excel. Ошибка: {e}")
                continue
    #Цель поиска - Фамилия
    target_sursname = input_line.get().strip().title()
    user_full_name = None

    #Защита от ошибки - Задаю 0-е значения, на случай остутсвия открытых отчетов
    total_result_report_1 = 0
    total_result_report_2 = 0
    total_result_report_3 = 0
    total_result_report_4 = 0

    if target_sursname.strip() == "":
        label.configure(text=f"Введите фамилию в строку ввода.\n\n")
    else:
        #Отчет по работе с РТ
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

            sursname_report_1 = df_report_1["ФИО"].astype(str).str.strip().str.lower().str.startswith(target_sursname.strip().lower())
            found_rows = df_report_1[sursname_report_1]

            if not found_rows.empty:
                row = found_rows.iloc[0]
                if user_full_name is None:
                    user_full_name = str(row["ФИО"]).strip()

                def safe_val(val):
                    if pd.isna(val) or val is None or str(val).strip() == "":
                        return 0
                    return float(val)
                #Перемножение позиций из отчета РТ
                value_k = safe_val(row["Позиций собрано (Перемещения)"]) * 5.30
                value_l = safe_val(row["Коробок собрано (Перемещения)"]) * 2
                value_o = safe_val(row["Позиций размещено (Розница)"]) * 6.80
                value_p = safe_val(row["Коробок размещено (Розница)"]) * 3.80
                value_s = safe_val(row["Позиций собрано (ОПТ сборка)"]) * 4
                value_t = safe_val(row["Коробок собрано (ОПТ сборка)"]) * 2.20
                value_w = safe_val(row["Внутренние перемещения"]) * 18

                total_result_report_1 = value_k + value_l + value_o + value_p + value_s + value_t + value_w                        

        #Отчет по делению паллет
        if report_status_2:
            label_report_2.configure(text=f"Отчет: Деление паллет приемка - ✅")
            report_data_2 = report_status_2.used_range.value

            df_report_2 = pd.DataFrame(data=report_data_2[3:], columns=report_data_2[2])     #DataFrame Pandas
            df_report_2 = df_report_2.fillna(0)                                              #Убираем пустые ячейки

            sursname_report_2 = df_report_2[df_report_2["ФИО"].astype(str).str.startswith(target_sursname)]
            if not sursname_report_2.empty:
                user_full_name = sursname_report_2["ФИО"].iloc[0]

            wholesale = sursname_report_2[sursname_report_2["Зона размещения"].astype(str).str.contains("опт", case=False)]       #Берем значения по заданой фразе (опт с машины - коробки)
            goods_receipt = sursname_report_2[sursname_report_2["Зона размещения"].astype(str).str.contains("0", case=False)]     #Берем значения по заданой фразе (приемка товара - коробки)
            retail = sursname_report_2[ ~ sursname_report_2["Зона размещения"].astype(str).str.contains("0|опт", case=False)]     #Берем всё кроме заданной фразы (розница с машины - коробки)

            ssci_click = sursname_report_2["В т.ч. отсканировано КИЗ"].sum()          #Кизы приемка
            wholesale_box = wholesale["Кол-во максималок"].sum()                      #Коробки с машины размещено - ОПТ
            wholesale_position = wholesale["Количество позиций, шт"].sum()            #Позицый с приемки размещено - ОПТ
            retail_box = retail["Кол-во максималок"].sum()                            #Коробки с машины размещено - розница
            retail_position = retail["Количество позиций, шт"].sum()                  #Позиций с машины размещено - розница
            goods_receipt_box = goods_receipt["Кол-во максималок"].sum()              #Коробок принято с машины - приемка
            goods_receipt_position = goods_receipt["Количество позиций, шт"].sum()    #Позиций принято с машины - приемка

            total_result_report_2 = wholesale_box * 1.50 + retail_box * 3.80 + goods_receipt_box * 2.25 + ssci_click * 1.20 + wholesale_position * 7 + retail_position * 6.80 + goods_receipt_position * 1.90     #Результат по отчету: "Деленение паллет приемки"

        #Отчет по отгрузке экспедиция (зики)
        if report_status_3:
            label_report_3.configure(text=f"Отчет: Экспедиция-отгрузка - ✅")
            report_data_3 = report_status_3.used_range.value

            df_report_3 = pd.DataFrame(data=report_data_3[1:], columns=report_data_3[0])    #DataFrame Pandas
            df_report_3 = df_report_3.fillna(0)                                             #Убираем пустые ячейки

            sursname_report_3 = df_report_3[df_report_3["Экспедитор"].astype(str).str.startswith(target_sursname)]

            if not sursname_report_3.empty:
                user_full_name = sursname_report_3["Экспедитор"].iloc[0]

            total_result_report_3 = int(sursname_report_3.shape[0] * 4.90)      #Сумма с пиков

        #Отчет по перемещениям (зачистки)
        if report_status_4:
            label_report_4.configure(text=f"Отчет: Данные по перемещениям - ✅")
            report_data_4 = report_status_4.used_range.value

            df_report_4 = pd.DataFrame(data=report_data_4[1:], columns=report_data_4[0])    #DataFrame Pandas
            df_report_4 = df_report_4.fillna(0)     #Убираем пустые ячейки

            sursname_report_4 = df_report_4[df_report_4["Разместитель"].astype(str).str.startswith(target_sursname)]

            if not sursname_report_4.empty:
                user_full_name = sursname_report_4["Разместитель"].iloc[0]

            type_operation = sursname_report_4[sursname_report_4["Тип"].astype(str).str.contains("Зачистка комнаты", case=False)]
            operation_sector_retail = type_operation[type_operation["Участок размещения"].astype(str).str.contains("розничный участок сборки")]
            total_operation_sector_retail = operation_sector_retail["Позиций"].sum()        #Розничные зачистки (позиции)
            result_retail = int(total_operation_sector_retail * 55)
            operation_sector_opt = type_operation[ ~ type_operation["Участок размещения"].str.contains("К- розничный участок сборки")]
            total_opt = operation_sector_opt[["Позиций", "Максималок"]].sum()
            result_position_opt = total_opt["Позиций"]
            result_box_opt = total_opt["Максималок"]
            result = int(result_position_opt * 12.1 + result_box_opt * 5.80)
            total_result_report_4 = result_retail + result

        #Финальный результат
        if user_full_name is None:
            label.configure(text=f"Сотрудника с фамилией '{target_sursname}' не найдено.\nПроверьте корректность введенных данных\nили правильность открытых отчетов.")
        else:
            user_first_name = user_full_name.split()[1]
            total_all_report_dirt = int(total_result_report_1 + total_result_report_2 + total_result_report_3 + total_result_report_4)
            total_all_report_clean = int(total_all_report_dirt - (total_all_report_dirt / 100 * 13))
            app.after(0, update_ui, user_first_name, total_all_report_dirt, total_all_report_clean, total_result_report_1, total_result_report_2, total_result_report_3, total_result_report_4)       

#Создание диаграммы
def create_pie_chart(tab_name, sum_1, sum_2, sum_3, sum_4):
    for widget in tab_name.winfo_children():
        widget.destroy()      #Очищаем вкладку "Статистика"

    labels = []     #Наименование "Долек"
    sizes = []      #Заработок с отчета

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

    #Отрисовка графика
    plt.style.use("dark_background")                                     #Темная тема для графика
    fig, ax = plt.subplots(figsize=(5, 4), dpi=100)                      #Создаем фигуру
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728"]       #Цвет "Долек"

    ax.pie(sizes, labels=labels, autopct="%1.1f%%", startangle=140, colors=colors[:len(sizes)], radius=0.9, textprops={"fontsize": 11})
    ax.axis("equal")     #Ровный круг

    fig.patch.set_facecolor("#2b2b2b")      #Закрашиваем фон в цвет окна программы

    canvas = FigureCanvasTkAgg(fig, master=tab_name)
    canvas.draw()
    canvas.get_tk_widget().pack(pady=10, padx=10, fill="both", expand=True)        

#Обновление интерфейса
def update_ui(first_name, dirt, clean, sum_1, sum_2, sum_3, sum_4):
    label.configure(text=f"{first_name}, ваш результат:\nГрязный заработок: {dirt}₽\nЧистый заработок: {clean}₽")    #Обновляем текст
    create_pie_chart(tabview.tab("Статистика"), sum_1, sum_2, sum_3, sum_4)                                          #Рисуем график

#Кнопка
def button_start_click(event=None):
    thread = threading.Thread(target=run_calculations, daemon=True)
    thread.start()

#Enter
app.bind("<Return>", button_start_click)


###Fronted
##Окно программы
#Панель вкладок
tabview = ctk.CTkTabview(app, width=380, height=450)
tabview.pack(pady=10, padx=10)

#Вкладки
tabview.add("Калькулятор")
tabview.add("Статистика")
tabview.add("Инструкция")

#Строка вывода результата
label = ctk.CTkLabel(tabview.tab("Калькулятор"), text="Добро пожаловать.\n\n", font=("Arial", 17))
label.pack(pady=20)

#Строка ввода
input_line = ctk.CTkEntry(tabview.tab("Калькулятор"), placeholder_text="Введите вашу фамилию", width=200)
input_line.pack(pady=10)

#Кнопка
button_start = ctk.CTkButton(tabview.tab("Калькулятор"), text="Рассчитать З/П", command=button_start_click)
button_start.pack(pady=10)


##Статусбары отчетов
#Фрэйм для красивого текста
reports_frame = ctk.CTkFrame(tabview.tab("Калькулятор"), fg_color="transparent")
reports_frame.pack(pady=10)
#РТ
label_report_1 = ctk.CTkLabel(reports_frame, text="Отчет: Работа с РТ - ❌", font=("Arial", 15), anchor="w")
label_report_1.pack(pady=3, fill="x", anchor="w")

#Деление паллет приемщиков
label_report_2 = ctk.CTkLabel(reports_frame, text="Отчет: Деление паллет приемка - ❌", font=("Arial", 15), anchor="w")
label_report_2.pack(pady=3, fill="x", anchor="w")

#Отгрузка экспедиции (пики)
label_report_3 = ctk.CTkLabel(reports_frame, text="Отчет: Экспедиция-отгрузка - ❌", font=("Arial", 15), anchor="w")
label_report_3.pack(pady=3, fill="x", anchor="w")

#Данные по перемещениям(зачистки)
label_report_4 = ctk.CTkLabel(reports_frame, text="Отчет: Данные по перемещениям - ❌", font=("Arial", 15), anchor="w")
label_report_4.pack(pady=3, fill="x", anchor="w")

label_instructions = ctk.CTkLabel(tabview.tab("Инструкция"), text=f"Инструкция", font=("Arial", 17))
label_instructions.pack(pady=4)

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
label_instructions.pack(pady=4)

###Бесконечный цикл
app.mainloop()