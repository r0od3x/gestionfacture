
import re
from pdfminer.high_level import extract_text
from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import os
import main as backend
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import psycopg2
import re
import pandas as pd
import re
from datetime import datetime, timedelta, date



root = tk.Tk()
root.withdraw()


file_path = filedialog.askopenfilename(
    title="Select Excel file",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

# Load the selected file into a DataFrame
if file_path:
    df = pd.read_excel(file_path, engine="openpyxl")
    print("File loaded successfully.")
    print(df.head())
else:
    print("No file selected.")


print(df.columns)


df.columns = df.columns.str.strip().str.lower()

df = df.drop_duplicates(subset=['nomenclature', 'declaration', 'qenkg'], keep='first')




pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"


df['date'] = df['date'].apply(
    lambda d: datetime.strptime(d, "%d/%m/%Y").date() if isinstance(d, str)
    else d.date() if hasattr(d, 'date')
    else None
)


print("hiiiiiiiiiiiii")

# Connect to PostgreSQL
conn = psycopg2.connect(
    host="localhost",
    database="sommier",
    user="postgres",
    password="reda2005"
)
cursor = conn.cursor()

cursor.execute("DELETE FROM sommier;")

# Convert DataFrame to list of tuples
data = list(df.itertuples(index=False, name=None))

# Bulk insert
execute_values(
    cursor,
    "INSERT INTO sommier  VALUES %s",
    data
)

conn.commit()
cursor.close()
conn.close()






import re

def extractdata(filepath):

    lines = extract_text(filepath).splitlines()
    lines = [line.strip() for line in lines if line.strip()]
    full_text = "\n".join(lines)


    patterns = {
        "Nom Client": r"Nom Client\s*[:\-]?\s*(.+)",
        "Code article": r"Code article\s*[:\-]?\s*(.+)",
        "Cannelure": r"Cannelure\s*[:\-]?\s*(.+)",        
        "Composition": r"Composition\s*[:\-]?\s*(.+)",    
        "Surface brut": r"Surface brut\s*[:\-]?\s*(.+)",
        "Poids brut": r"Poids brut\s*[:\-]?\s*(.+)",
        "Nom fiche": r"Num\s*:\s*(.*?)\s*-\s*Date\s*:"
    }

    results = []
    for label, pattern in patterns.items():
        matches = re.findall(pattern, full_text, re.IGNORECASE)
    
        if label in ["Cannelure", "Composition"]:
        
            value = matches[1].strip() if len(matches) >= 2 else "N/A"
        else:
        
            value = matches[0].strip() if matches else "N/A"
    
        results.append(value)

    # Remove everything after the last digit in results[1]
    match = re.match(r'^(\d.*\d)', results[1])
    if match:
        results[1] = match.group(1)

    if '/' in results[1]:
        slash_index = results[1].index('/')
        if slash_index >= 6:
            before = results[1][slash_index - 6:slash_index]
        else:
            before = results[1][:slash_index]
        after = results[1][slash_index:]  # includes '/'
    else:
        before = results[1]
        after = ""

    if len(before) >= 2 and before[1] == '5':
        before = before[0] + 'S' + before[2:]

    results[1] = before + after

    results[2] = results[2][:-2]
    results[4] = results[4][:-3]
    results[5] = results[5][:-3]
    
    # Remove everything after the first occurrence of double space in results[3]
    if ' ' in results[3]:
        results[3] = results[3].split(' ')[0]
        results[3] = results[3][:-4]
    print(results)

    return results





import re

def verifierfiles(fprod, fcess):
    k = 0
    missing_codes = []

    # Clean codes in fprod by removing trailing non-digit chars
    for i in range(len(fprod)):
        s = fprod[i][1]
        if not s[-1].isdigit():
            fprod[i][1] = s[:-1]  
    m = []
    p = []
    f = []
    for x in range(len(fcess[0])):
        text = fcess[0][x]
        match = re.match(r'^([^,\s]+)', text)
        if match:
            first_part = match.group(1)

            found = False
            for y in range(len(fprod)):
                if first_part == fprod[y][1]:
                    if fprod[y][6] in m:
                        continue
                    else:
                        p.append(fprod[y])
                        f.append(fcess[0][x])
                        m.append(fprod[y][6])
                        found = True
                    

            if not found:
                missing_codes.append(first_part)

    if missing_codes:
        raise ValueError(f"Code(s) not found in fprod: {', '.join(missing_codes)}")

    print("muhaha")
    print(p)
    print(f)
    return p,f




def extractcession(filepath):
    images = convert_from_path(filepath)
    full_text = ""

    for img in images:
        full_text += pytesseract.image_to_string(img, lang='fra') + "\n"

    # Clean lines
    lines = [line.strip() for line in full_text.splitlines() if line.strip()]

    start_marker = "Montant"
    end_marker = "TOTAL"

    inside_block = False
    extracted_lines = []

    for line in lines:

        if start_marker.lower() in line.lower():
            inside_block = True
            continue
        
        elif end_marker.lower() in line.lower() and inside_block:
            break

        if inside_block:
            if re.match(r'^(mad|eur|usd)\b', line.lower()):
                continue

            if not re.search(r'\d.*\/.*\d', line): 
                continue

            extracted_lines.append(line.strip())
    cleaned_lines = []
    for line in extracted_lines:
        if '/' in line:
            slash_index = line.index('/')
            if slash_index >= 6:
                before = line[slash_index - 6:slash_index]
            else:
                before = line[:slash_index]
            after = line[slash_index:]  # includes the '/'
        else:
            before = line
            after = ""

        if len(before) >= 2 and before[1] == '5':
            before = before[0] + 'S' + before[2:]

        cleaned_lines.append(before + after)

    extracted_lines = cleaned_lines



    print(extracted_lines)
    return extracted_lines



def calculpoidsbrut(composition,cannelure,surface):
    surface=float(surface)
    B = 1.38
    C = 1.48
    E = 1.32
    parts = composition.split("/")
    parts= [int(p[2:]) for p in parts]
    cann = list(cannelure)
    j=0
    for i in range(len(parts)):
        if i%2 == 0:
            continue

        if cann[j] == 'B' :
            parts[i]=parts[i]*B
            j=j+1
            
        elif cann[j] == 'C':
            parts[i]=parts[i]*C
            j=j+1
        elif cann[j] == 'e':
            parts[i]=parts[i]*E
            j=j+1
            
    total = sum(parts)*surface
    total=total/1000

    
    
    return total


def calculpoidsnet(total, dechet):
    if isinstance(total, list):
        total = total[0]  # Extract the float value
    x = total * (dechet / 100)
    x = total - x
    return x


 

import re
from collections import defaultdict

import re
from collections import defaultdict

def calculercomposition(composition, cannelure, surface, quantite_str):
    words = re.findall(r'\S+', quantite_str)  # Matches non-space tokens
    print("Extracted words:", words)

    if len(words) >= 3:
        raw_value = words[-3]  # 3rd from last
        print("3rd from last word:", raw_value)

        clean_value = raw_value.replace('.', '').replace(',', '')  # Remove formatting
        if clean_value.isdigit():
            quantite = int(clean_value)
            print("Numeric quantite:", quantite)
            
        else:
            raise ValueError("Extracted value is not numeric")
    else:
        raise ValueError("Quantite string format invalid")

    surface = float(surface)
    B = 1.38
    C = 1.48
    E = 1.32

    parts = composition.split("/")                     
    new_parts = [part[:-3] for part in parts]          
    parts = [int(p[-3:]) for p in parts]

    cann = list(cannelure)                             
    j = 0

    for i in range(len(parts)):
        if i % 2 == 0:
            continue
        if j < len(cann):
            if cann[j] == 'B':
                parts[i] *= B
            elif cann[j] == 'C':
                parts[i] *= C
            elif cann[j] == 'E' or cann[j] == 'e':
                parts[i] *= E
            else:
                raise ValueError(f"Cannelure coefficient '{cann[j]}' is invalid. Expected 'B', 'C', or 'E'.")
            j += 1

    parts = [x * surface for x in parts]

    result = defaultdict(float)

    for name, value in zip(new_parts, parts):
        result[name] += value

    unique_names = list(result.keys())
    total_values = [result[name] for name in unique_names]
    total_values = [(x * quantite / 1000) for x in total_values]

    return unique_names, total_values



def convert_to_datetime(value):
    if isinstance(value, datetime):
        return value
    elif isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    elif isinstance(value, str):
        try:
            return datetime.strptime(value, "%Y-%m-%d")
        except ValueError:
            pass  # Try other formats if needed
    elif isinstance(value, float):
        # Excel serial date handling (1900-based)
        excel_start_date = datetime(1899, 12, 30)  # Excel's day 0
        return excel_start_date + timedelta(days=int(value))
    return None  # Unrecognized type






def get_and_process_rows(names, values_to_remove, output_excel='SOMMIER.xlsx'):
    weight = []
    
    conn = psycopg2.connect(
        host="localhost",
        database="sommier",
        user="postgres",
        password="reda2005"
    )
    cursor = conn.cursor()

    placeholders = ','.join(['%s'] * len(names))
    query = f"""
         SELECT * 
         FROM sommier 
         WHERE split_part(qualite, ' ', 1) IN ({placeholders}) 
         AND COALESCE(NULLIF(resteenkg, ''), '0')::numeric <> 0
         ORDER BY date;
    """
    cursor.execute(query, names)
    rows = cursor.fetchall()
    ten_days_ago = datetime.now() - timedelta(days=10)
    filtered_rows = []
    
    for row in rows:
        raw_date = row[4]  # Adjust index if needed
        row_datetime = convert_to_datetime(raw_date)
        if row_datetime and row_datetime < ten_days_ago:
            filtered_rows.append(row)

    # ✅ Raise exception if no rows are found
    if not filtered_rows:
        raise ValueError(f"No matching entries found in the database for: {names}")

    colnames = [desc[0] for desc in cursor.description]
    df_filtered = pd.DataFrame(filtered_rows, columns=colnames)

    df_filtered['original_qenkg'] = df_filtered['qenkg']

    df_filtered['resteenkg'] = pd.to_numeric(df_filtered['resteenkg'], errors='coerce')
    df_filtered['qenkg'] = pd.to_numeric(df_filtered['qenkg'], errors='coerce')
    df_filtered['reste'] = pd.to_numeric(df_filtered['reste'], errors='coerce')
    
    remaining_value = values_to_remove
    modified_rows = []

    for idx, row in df_filtered.iterrows():
        if remaining_value <= 0:
            print(f"Remaining value depleted at index {idx}")
            break

        current_resteenkg = row['resteenkg']
        qenkg = row['qenkg']

        if pd.isna(current_resteenkg) or pd.isna(qenkg):
            print(f"Skipping index {idx} due to NaN values")
            continue

        # Determine how much to remove from this row
        if current_resteenkg <= remaining_value:
            amount_removed = current_resteenkg
            new_resteenkg = 0
            new_reste = 0
        else:
            amount_removed = remaining_value
            new_resteenkg = current_resteenkg - amount_removed
            new_reste = (new_resteenkg / qenkg) * 100

        # Append the removed amount to weight list
        weight.append(amount_removed)
        print(f"Index {idx} - Removed: {amount_removed}, New resteenkg: {new_resteenkg}, New reste: {new_reste}")

        # Update the database
        update_query = """
            UPDATE sommier
            SET resteenkg = %s, reste = %s
            WHERE nomenclature = %s AND declaration = %s AND qenkg = %s;
        """
        cursor.execute(update_query, (
            new_resteenkg,
            new_reste,
            str(row['nomenclature']),
            float(row['declaration']),
            str(row['original_qenkg'])
        ))

        # Track changes in DataFrame
        df_filtered.at[idx, 'resteenkg'] = new_resteenkg
        df_filtered.at[idx, 'reste'] = new_reste
        modified_rows.append(df_filtered.loc[idx])

        remaining_value -= amount_removed
        print(f"Remaining value after index {idx}: {remaining_value}")

    conn.commit()

    full_df = pd.read_sql_query("SELECT * FROM sommier ORDER BY date", conn)
    full_df.to_excel(output_excel, index=False)

    cursor.close()
    conn.close()

    return pd.DataFrame(modified_rows), weight



def calculpoidsdeclarer(fprod, fcess):
    x = []
    for i in range(len(fcess)):
        match = re.search(r'(?:\S+\s+)*(\S+)\s+\S+\s+\S+$', fcess[i])
        if match:
            raw_value = match.group(1)
            
            clean_value = raw_value.replace('.', '')
            numeric_value = int(clean_value)  
            print("Numeric:", numeric_value)
            

            poids_brut = backend.calculpoidsbrut(fprod[i][3], fprod[i][2], fprod[i][4])
            poids_net = backend.calculpoidsnet(poids_brut, 15)

            x.append(poids_net * numeric_value)

    total = sum(x) - (sum(x) * 2 / 100)
    return total



def sum_fourth_numbers(matrix):
    total = 0
    for row in matrix:
        
        parts = row.split()
        if len(parts) >= 4:
            number_str = parts[-3].replace('.', '')  
            try:
                number = int(number_str)
                total += number
            except ValueError:
                continue 
    
    return str(total)



