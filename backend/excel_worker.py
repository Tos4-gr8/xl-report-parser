import xlwings as xw

def find_active_reports():
    """
    Scans opened Excel workbooks and matches them with required reports.
    
    Returns:
        dict: Active report sheets mapped to their specific types.
    """
    reports = {
        "rt": None,
        "pallet": None,
        "shipping": None,
        "transfer": None
    }
    
    for excel_app in xw.apps:
        for book in excel_app.books:
            try:
                sheet = book.sheets.active
        
                report_book_1 = str(sheet.range('A1').value or '').strip().lower()
                report_book_2 = str(sheet.range('C1').value or '').strip().lower()
                report_book_3 = str(sheet.range('A1').value or '').strip().lower()
                report_book_4 = str(sheet.range('H1').value or '').strip().lower()

                # Match reports based on specific phrases in cells
                if "производительность приемщиков и переместителей с рт" in report_book_1:
                    reports["rt"] = sheet
                if "загрузка приёмщиков за период" in report_book_2:
                    reports["pallet"] = sheet
                if "отгрузка" in report_book_3:
                    reports["shipping"] = sheet
                if "участок сборки" in report_book_4:
                    reports["transfer"] = sheet
                    
            except Exception as e:
                print(f"Skipped hidden or system Excel workbook. Error: {e}")
                continue
                
    return reports
