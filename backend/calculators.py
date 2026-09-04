import pandas as pd
import xlwings as xw
from config import RATES_REPORT_1, RATES_REPORT_2, RATES_REPORT_3, RATES_REPORT_4, TAX_PERCENT


def safe_val(val):
    """Safely convert a value to a float, returning 0 if it is missing or empty."""
    if pd.isna(val) or val is None or str(val).strip() == "":
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


# ==============================================================================
# Report-specific calculation functions
# ==============================================================================

def calculate_rt_operations(sheet, target_surname):
    """Calculate the 'RT Operations' report metrics."""
    df = sheet.used_range.options(pd.DataFrame, index=False, header=False).value
    if df.empty:
        return 0.0, None
        
    df.rename(columns={
        0: "ФИО",
        10: "Позиций собрано (Перемещения)",
        11: "Коробок собрано (Перемещения)",
        14: "Позиций размещено (Розница)",
        15: "Коробок размещено (Розница)",
        18: "Позиций собрано (ОПТ сборка)",
        19: "Коробок собрано (ОПТ сборка)",
        22: "Внутренние перемещения",
    }, inplace=True)

    # Case-insensitive search preventing failures on missing rows
    surname_filter = df["ФИО"].astype(str).str.strip().str.lower().str.startswith(target_surname.lower(), na=False)
    found_rows = df[surname_filter]

    if not found_rows.empty:
        row = found_rows.iloc[0]
        user_full_name = str(row["ФИО"]).strip()

        value_k = safe_val(row["Позиций собрано (Перемещения)"]) * RATES_REPORT_1["pos_movement"]
        value_l = safe_val(row["Коробок собрано (Перемещения)"]) * RATES_REPORT_1["box_movement"]
        value_o = safe_val(row["Позиций размещено (Розница)"]) * RATES_REPORT_1["pos_retail"]
        value_p = safe_val(row["Коробок размещено (Розница)"]) * RATES_REPORT_1["box_retail"]
        value_s = safe_val(row["Позиций собрано (ОПТ сборка)"]) * RATES_REPORT_1["pos_wholesale"]
        value_t = safe_val(row["Коробок собрано (ОПТ сборка)"]) * RATES_REPORT_1["box_wholesale"]
        value_w = safe_val(row["Внутренние перемещения"]) * RATES_REPORT_1["internal_move"]

        total = value_k + value_l + value_o + value_p + value_s + value_t + value_w
        return total, user_full_name
    return 0.0, None


def calculate_pallet_splitting(sheet, target_surname):
    """Calculate the 'Pallet Splitting & Receiving' report metrics."""
    report_data = sheet.used_range.value
    if not report_data or len(report_data) <= 3:
        return 0.0, None
        
    df = pd.DataFrame(data=report_data[3:], columns=report_data[2]).fillna(0)

    # Case-insensitive search preserving the original structural logic
    surname_filter = df[df["ФИО"].astype(str).str.strip().str.lower().str.startswith(target_surname.lower())]
    
    if not surname_filter.empty:
        user_full_name = surname_filter["ФИО"].iloc[0]

        # Target storage zones filtering (Wholesale, Goods Receipt, Retail)
        wholesale = surname_filter[surname_filter["Зона размещения"].astype(str).str.contains("опт", case=False)]
        goods_receipt = surname_filter[surname_filter["Зона размещения"].astype(str).str.contains("0", case=False)]
        retail = surname_filter[~surname_filter["Зона размещения"].astype(str).str.contains("0|опт", case=False)]

        box_col = "Кол-во maximalлок" if "Кол-во maximalлок" in df.columns else "Кол-во максималок"

        ssci_click = surname_filter["В т.ч. отсканировано КИЗ"].sum()
        wholesale_box = wholesale[box_col].sum()
        wholesale_position = wholesale["Количество позиций, шт"].sum()
        retail_box = retail[box_col].sum()
        retail_position = retail["Количество позиций, шт"].sum()
        goods_receipt_box = goods_receipt[box_col].sum()
        goods_receipt_position = goods_receipt["Количество позиций, шт"].sum()

        total = (wholesale_box * RATES_REPORT_2["wholesale_box"] + 
                 retail_box * RATES_REPORT_2["retail_box"] + 
                 goods_receipt_box * RATES_REPORT_2["receipt_box"] + 
                 ssci_click * RATES_REPORT_2["ssci_click"] + 
                 wholesale_position * RATES_REPORT_2["wholesale_pos"] + 
                 retail_position * RATES_REPORT_2["retail_pos"] + 
                 goods_receipt_position * RATES_REPORT_2["receipt_pos"])
        return total, user_full_name
    return 0.0, None


def calculate_shipping(sheet, target_surname):
    """Calculate the 'Shipping & Dispatch' report metrics."""
    report_data = sheet.used_range.value
    df = pd.DataFrame(data=report_data[1:], columns=report_data[0]).fillna(0)

    # Case-insensitive search keeping shape[0] calculation intact
    surname_filter = df[df["Экспедитор"].astype(str).str.strip().str.lower().str.startswith(target_surname.lower())]
    if not surname_filter.empty:
        user_full_name = surname_filter["Экспедитор"].iloc[0]
        total = int(surname_filter.shape[0] * RATES_REPORT_3["base_rate"])
        return total, user_full_name
    return 0, None


def calculate_transfer(sheet, target_surname):
    """Calculate the 'Inventory Transfers' report metrics."""
    report_data = sheet.used_range.value
    df = pd.DataFrame(data=report_data[1:], columns=report_data[0]).fillna(0)

    # Case-insensitive placement search
    surname_filter = df[df["Разместитель"].astype(str).str.strip().str.lower().str.startswith(target_surname.lower())]
    if not surname_filter.empty:
        user_full_name = surname_filter["Разместитель"].iloc[0]
        
        # Casting to string to prevent unexpected failures on null types
        type_operation = surname_filter[surname_filter["Тип"].astype(str).str.contains("Зачистка комнаты", case=False)]
        
        # Filtering retail sector
        operation_sector_retail = type_operation[type_operation["Участок размещения"].astype(str).str.contains("розничный участок сборки")]
        total_operation_sector_retail = operation_sector_retail["Позиций"].sum()
        result_retail = int(total_operation_sector_retail * RATES_REPORT_4["retail_sector"])
        
        # Original inversion logic ensuring strict data alignment
        operation_sector_opt = type_operation[~type_operation["Участок размещения"].astype(str).str.contains("К- розничный участок сборки")]
        total_opt = operation_sector_opt[["Позиций", "Максималок"]].sum()
        result_position_opt = total_opt["Позиций"]
        result_box_opt = total_opt["Максималок"]
        result = int(result_position_opt * RATES_REPORT_4["opt_position"] + result_box_opt * RATES_REPORT_4["opt_box"])
        
        return (result_retail + result), user_full_name
    return 0, None


def calculate_clean_salary(dirty_salary):
    """Calculate the net salary after tax deduction."""
    return int(dirty_salary - (dirty_salary / 100 * TAX_PERCENT))
