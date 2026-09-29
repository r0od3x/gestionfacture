import xlwings as xw
from datetime import datetime
import os
import time

import config

FILE_PATH = config.CALCUL_TEMPLATE
SHEET_NAME = "FICHE"


def update_excel_layout_full(h8_value, new_data,poidsfinal,export_pdf=True,result_folder_path=config.RESULT_DIR):
    outputfolder = os.path.join(result_folder_path, h8_value, datetime.today().strftime("%d-%m-%Y"))

    os.makedirs(outputfolder, exist_ok=True)
    insert_at = 14
    rows_to_delete = 5
    rows_to_insert = len(new_data)
    max_col = 12

    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12

    xlContinuous = 1
    xlDash = 2
    xlNone = -4142
    xlCenter = -4108

    FONT_SIZE = 10
    YELLOW = (255, 255, 0)
    
    app = xw.App(visible=False)
    wb = xw.Book(FILE_PATH)
    ws = wb.sheets[SHEET_NAME]

    

    try:
        # Set L23 to 'hiii' before anything else
        ws.range("L23").value = "hiii"

        # 1) Clear and unmerge rows 14–18
        for row in range(insert_at, insert_at + rows_to_delete):
            for col in range(1, max_col + 1):
                cell = ws.api.Cells(row, col)
                if cell.MergeCells:
                    cell.MergeArea.UnMerge()
                cell.Value = ""
                cell.ClearFormats()

        # 2) Shift rows BELOW 22 downward
        last_row = ws.api.UsedRange.Rows.Count + ws.api.UsedRange.Row - 1
        ws.api.Rows(f"23:{last_row}").Insert(Shift=1)

        # 3) Insert data
        gpc_cycle = ["G", "P", "C"] 
        for i, row_data in enumerate(new_data):
            row_num = insert_at + i
            is_even = i % 2 == 0
            gpc_val = gpc_cycle[i % len(gpc_cycle)]

            for j in range(1, max_col + 1):
                col_letter = chr(ord('A') + j - 1)
                cell = ws.range(f"{col_letter}{row_num}")

                # Alternating values
                val = ""
                if col_letter == 'B' and len(row_data) > 0:
                    val = row_data[0] if is_even else ""
                if col_letter == 'C' and len(row_data) > 0:
                    val = row_data[1] if is_even else ""
                elif col_letter == 'D' and len(row_data) > 1:
                    val = row_data[3] if is_even else ""
                elif col_letter == 'E' and len(row_data) > 2:
                    val = "" if is_even else ""
                elif col_letter == 'F':
                    val = "G P C" if is_even else ""
                elif col_letter == 'G' and len(row_data) > 3:
                    val = 'PAPIER' if is_even else ""
                elif col_letter == 'H' and len(row_data) > 4:
                    val = row_data[2] if is_even else ""
                elif col_letter == 'I':
                    val = "15%" if is_even else ""
                elif col_letter == 'J':
                    val = "B." if is_even else "N."
                elif col_letter == 'K' and len(row_data) > 5:
                    val = row_data[4] if is_even else row_data[5]
                elif col_letter == 'L':
                    val = "OUI" if is_even else ""
                    cell.color = YELLOW

                cell.value = val
                cell.api.Font.Size = FONT_SIZE
                cell.api.HorizontalAlignment = xlCenter
                cell.api.VerticalAlignment = xlCenter

        # 4) Apply borders every 2 rows
        for i in range(0, rows_to_insert, 2):
            top_row = insert_at + i
            bottom_row = top_row + 1 if (i + 1) < rows_to_insert else top_row

            for col in range(1, max_col + 1):
                col_letter = chr(ord('A') + col - 1)

                top_cell = ws.range(f"{col_letter}{top_row}").api
                bottom_cell = ws.range(f"{col_letter}{bottom_row}").api

                top_border = top_cell.Borders(xlEdgeTop)
                top_border.LineStyle = xlDash
                top_border.Weight = 2
                top_border.ColorIndex = 1

                bottom_border = bottom_cell.Borders(xlEdgeBottom)
                bottom_border.LineStyle = xlDash
                bottom_border.Weight = 2
                bottom_border.ColorIndex = 1

                for edge in [xlEdgeLeft, xlEdgeRight]:
                    b1 = top_cell.Borders(edge)
                    b2 = bottom_cell.Borders(edge)
                    for b in [b1, b2]:
                        b.LineStyle = xlContinuous
                        b.Weight = 2
                        b.ColorIndex = 1

                if bottom_row != top_row:
                    bottom_cell.Borders(xlInsideHorizontal).LineStyle = xlNone

                # 5) Delete everything below the inserted block
                delete_start = insert_at + rows_to_insert + 1
                last_row = ws.api.UsedRange.Rows.Count + ws.api.UsedRange.Row - 1
                if delete_start <= last_row:
                    ws.api.Rows(f"{delete_start}:{last_row}").Delete()

                 # 6) Add value in K and L two rows after inserted block
                extra_row = insert_at + rows_to_insert + 2
                ws.range(f"K{extra_row}").value = "PDS A DECLARER"
                ws.range(f"L{extra_row}").value = poidsfinal
                ws.range("H8").value = h8_value  # You can change this to a real total if needed

        # Apply style to these two
                for col in ["K", "L"]:
                    cell = ws.range(f"{col}{extra_row}")
                    cell.api.Font.Size = FONT_SIZE
                    cell.api.HorizontalAlignment = xlCenter
                    cell.api.VerticalAlignment = xlCenter
                    if col == "L":
                        cell.color = YELLOW

        output = outputfolder +"/"+ h8_value+" CALCUL.xlsx"
        wb.save(output)

        if export_pdf:
            pdf_file = os.path.abspath(output.replace(".xlsx", ".pdf"))
            try:
                ps = ws.api.PageSetup
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = 1
                ps.Orientation = 2
                ps.CenterHorizontally = True
                ps.CenterVertically = True
                ws.api.ExportAsFixedFormat(0, pdf_file)
                print(f"📄 PDF exported: {pdf_file}")
            except Exception as e:
                print(f"❌ PDF export error: {e}")

    finally:
        wb.close()
        app.quit()

def append_row_to_excel(data, filepath='excel_table.xlsx', sheet_name='Sheet1'):
    
    app = xw.App(visible=False)
    wb = xw.Book(filepath)
    sheet = wb.sheets[sheet_name]

    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
    next_row = last_row + 1 if sheet.range(f"A{last_row}").value else last_row

    # Write the data to the next row
    sheet.range(f"A{next_row}").value = data

    # Save and close
    wb.save()
    wb.close()
    app.quit()
    print(f"Row added successfully to {filepath}")