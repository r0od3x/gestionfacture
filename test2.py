import xlwings as xw
from datetime import datetime
import os


FILE_PATH = "C:/tesseract/excelfiles/CESSION AT PESBAK.xlsx"    
SHEET_NAME = "ATT"                   


def update_excel_layout(d25_value, d26_value, new_data, export_pdf=False,result_folder_path="C:/tesseract/resultat" ):
    outputfolder = os.path.join(result_folder_path, d25_value, datetime.today().strftime("%d-%m-%Y"))
    os.makedirs(outputfolder, exist_ok=True)
    insert_at = 36
    rows_to_delete = 3
    rows_to_insert = len(new_data)
    today = datetime.today().strftime("%d/%m/%Y")

    
    app = xw.App(visible=False)
    wb = xw.Book(FILE_PATH)
    ws = wb.sheets[SHEET_NAME]

    try:
        
        ws.range("G7").value = f"MOHAMMEDIA LE, {today}"

        
        ws.range("D25").value = d25_value
        ws.range("D26").value = d26_value + " CAISSES EN CARTON"

       
        ws.api.Rows(f"{insert_at}:{insert_at + rows_to_delete - 1}").Delete()

        
        ws.api.Rows(f"{insert_at}:{insert_at + rows_to_insert - 1}").Insert()

        
        for i, row in enumerate(new_data):
            ws.range(f"C{insert_at + i}").value = row[0]
            ws.range(f"D{insert_at + i}").value = row[1]

        output = outputfolder +"/"+ d25_value+" ATTESTATION.xlsx"
        wb.save(output)

        
        


        if export_pdf:
    
            pdf_file = os.path.abspath(output.replace(".xlsx", ".pdf"))

        try:
        
            ws.api.PageSetup.Zoom = False  
            ws.api.PageSetup.FitToPagesWide = 1
            ws.api.PageSetup.FitToPagesTall = 1
            ws.api.PageSetup.Orientation = 1 # 1 = Portrait, 2 = Landscape (choose based on your layout)
            ws.api.PageSetup.CenterHorizontally = True
            ws.api.PageSetup.CenterVertically = True

            ws.api.ExportAsFixedFormat(0, pdf_file)
            print(f"📄 PDF export attempt to: {pdf_file}")

       
            if os.path.exists(pdf_file):
                print(f"✅ PDF successfully saved as: {pdf_file}")
            else:
                print("❌ PDF export failed. File not found.")

        except Exception as e:
            print(f"❌ Error during PDF export: {e}")


    finally:
        wb.close()
        app.quit()


