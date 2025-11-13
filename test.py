import re
from pdfminer.high_level import extract_text
from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import os
import main as backend
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog
import test2 as export
import test3 as exportcalcul
import customtkinter as ctk
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import matplotlib.pyplot as plt
from tkinter import ttk, filedialog
import matplotlib.gridspec as gridspec
import platform
from datetime import datetime
from main import file_path
from tkinter import messagebox

# --- GUI Setup ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

app = ctk.CTk()
app.title("📄 PDF Folder Processor")
app.resizable(False, False)
app.state('zoomed')

sidebar_width = 220

sidebar = ctk.CTkFrame(app, width=sidebar_width, corner_radius=0)
sidebar.pack(side="left", fill="y")

container = ctk.CTkFrame(app, fg_color="#1e1e2f", corner_radius=15)
container.pack(side="right", fill="both", expand=True, padx=25, pady=25)

def keep_maximized():
    app.state('zoomed')
app.after(100, keep_maximized)

def raise_frame(frame):
    frame.tkraise()

# AnimatedStatus class (unchanged)
class AnimatedStatus(ctk.CTkLabel):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.base_text = ""
        self._running = False
        self._dots = 0
    def start(self, text):
        self.base_text = text
        self._dots = 0
        self._running = True
        self._animate()
    def stop(self):
        self._running = False
    def _animate(self):
        if not self._running:
            return
        dots = '.' * (self._dots % 4)
        self.configure(text=f"{self.base_text}{dots}")
        self._dots += 1
        self.after(500, self._animate)

# ------------------ Frames ------------------
home_frame = ctk.CTkFrame(container)
home_frame.place(relwidth=1, relheight=1)

settings_frame = ctk.CTkFrame(container)
settings_frame.place(relwidth=1, relheight=1)

charts_frame = ctk.CTkFrame(container)
charts_frame.place(relwidth=1, relheight=1)

summary_frame = ctk.CTkFrame(container)
summary_frame.place(relwidth=1, relheight=1)

treeview_frame = ctk.CTkFrame(container)
treeview_frame.place(relwidth=1, relheight=1)
# ------------------ Home Frame UI ------------------
status_label = AnimatedStatus(
    home_frame,
    text="",
    font=ctk.CTkFont(size=20, weight="bold"),
    text_color="#50fa7b"
)
status_label.pack(pady=(25, 15))

header = ctk.CTkLabel(
    home_frame,
    text="Generateur d'ATT GPC",
    font=ctk.CTkFont(size=34, weight="bold"),
    text_color="#3fb0ac"
)
header.pack(pady=(0, 40))

# Notice: Buttons moved to sidebar — Home frame will show status and header only

# ------------------ Settings Frame UI ------------------
settings_label = ctk.CTkLabel(
    settings_frame,
    text="Settings",
    font=ctk.CTkFont(size=28, weight="bold"),
    text_color="#50fa7b"
)
settings_label.pack(pady=(30, 20))



def show_excel_treeview():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path)

        # Clear previous treeview content
        for widget in treeview_frame.winfo_children():
            widget.destroy()

        tree = ttk.Treeview(treeview_frame, columns=list(df.columns), show="headings")
        tree.pack(expand=True, fill="both")

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        raise_frame(treeview_frame)  # ✅ SWITCH TO treeview frame properly!

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")



def change_appearance_mode(choice):
    ctk.set_appearance_mode(choice)

appearance_label = ctk.CTkLabel(settings_frame, text="Appearance Mode:", font=ctk.CTkFont(size=18))
appearance_label.pack(pady=(10, 5))

appearance_optionmenu = ctk.CTkOptionMenu(settings_frame,
                                          values=["Light", "Dark", "System"],
                                          command=change_appearance_mode)
appearance_optionmenu.set("Dark")
appearance_optionmenu.pack(pady=(0, 20))

# ---- Default Paths ----
result_folder_path = r"C:\tesseract\resultat"
excel_file_path = r"C:\tesseract\excelfiles\default.xlsx"

# ---- Save Functions ----
def save_result_folder():
    folder = result_folder_entry.get()
    if os.path.isdir(folder):
        global result_folder_path
        result_folder_path = folder
        messagebox.showinfo("Settings", f"Result folder path set to:\n{folder}")
    else:
        messagebox.showerror("Error", "Invalid folder path!")

def save_excel_file():
    file = excel_file_entry.get()
    if os.path.isfile(file) and file.lower().endswith((".xlsx", ".xls")):
        global excel_file_path
        excel_file_path = file
        messagebox.showinfo("Settings", f"Excel file path set to:\n{file}")
    else:
        messagebox.showerror("Error", "Invalid Excel file!")

def browse_and_set_folder(entry_widget):
    folder = filedialog.askdirectory(title="Select Folder")
    if folder:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder)

def browse_and_set_file(entry_widget):
    file = filedialog.askopenfilename(title="Select Excel File",
                                      filetypes=[("Excel Files", "*.xlsx *.xls")])
    if file:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file)

# ---- Result Folder UI ----
result_folder_label = ctk.CTkLabel(settings_frame, text="Result Folder Path:", font=ctk.CTkFont(size=18))
result_folder_label.pack(pady=(10, 5))

result_folder_entry = ctk.CTkEntry(settings_frame, width=400)
result_folder_entry.insert(0, result_folder_path)
result_folder_entry.pack(pady=(0, 5))

browse_result_folder_btn = ctk.CTkButton(settings_frame, text="Browse...",
                                         command=lambda: browse_and_set_folder(result_folder_entry))
browse_result_folder_btn.pack(pady=(0, 5))

save_folder_btn = ctk.CTkButton(settings_frame, text="Save Result Folder", command=save_result_folder)
save_folder_btn.pack(pady=(0, 20))

# ---- Excel File UI ----
excel_file_label = ctk.CTkLabel(settings_frame, text="Excel File to Append To:", font=ctk.CTkFont(size=18))
excel_file_label.pack(pady=(10, 5))

excel_file_entry = ctk.CTkEntry(settings_frame, width=400)
excel_file_entry.insert(0, excel_file_path)
excel_file_entry.pack(pady=(0, 5))

browse_excel_file_btn = ctk.CTkButton(settings_frame, text="Browse...",
                                      command=lambda: browse_and_set_file(excel_file_entry))
browse_excel_file_btn.pack(pady=(0, 5))

save_excel_file_btn = ctk.CTkButton(settings_frame, text="Save Excel File", command=save_excel_file)
save_excel_file_btn.pack(pady=(0, 20))


# ---------------- Sidebar Buttons ----------------
sidebar_buttons = {}

def set_active(button_name):
    for name, btn in sidebar_buttons.items():
        if name == button_name:
            btn.configure(fg_color="#50fa7b", text_color="#282c34")
        else:
            btn.configure(fg_color="transparent", text_color="#8be9fd")

def on_select_folder():
    folder = filedialog.askdirectory(title="Select a folder")  # must be on main thread
    if folder:
        threading.Thread(target=run_processing, args=(folder,), daemon=True).start()

def on_open_result_folder():
    threading.Thread(target=open_tesseract_result_folder, daemon=True).start()

def on_home_click():
    raise_frame(home_frame)
    set_active("home")

def on_settings_click():
    raise_frame(settings_frame)
    set_active("settings")

def on_process_click():
    on_select_folder()

def on_open_results_click():
    on_open_result_folder()

btn_home = ctk.CTkButton(
    sidebar,
    text="🏠 Home",
    font=ctk.CTkFont(size=18, weight="bold"),
    command=on_home_click,
    fg_color="#50fa7b",
    hover_color="#70fa90",
    corner_radius=0,
    height=60,
    width=sidebar_width
)
btn_home.pack(fill="x")




btn_settings = ctk.CTkButton(
    sidebar,
    text="⚙️ Settings",
    font=ctk.CTkFont(size=18, weight="bold"),
    command=on_settings_click,
    fg_color="transparent",
    hover_color="#44475a",
    corner_radius=0,
    height=60,
    width=sidebar_width
)
btn_settings.pack(fill="x", pady=(15,0))

sidebar_buttons = {
    "home": btn_home,
    "settings": btn_settings,
    
}


raise_frame(home_frame)
set_active("home")


def on_charts_click():
    raise_frame(charts_frame)
    set_active("charts")

def on_summary_click():
    raise_frame(summary_frame)
    set_active("summary")

btn_charts = ctk.CTkButton(
    sidebar,
    text="📊 Charts",
    font=ctk.CTkFont(size=18, weight="bold"),
    command=on_charts_click,
    fg_color="transparent",
    hover_color="#44475a",
    corner_radius=0,
    height=60,
    width=sidebar_width
)
btn_charts.pack(fill="x", pady=(5, 0))

excel_button = ctk.CTkButton(
    sidebar,
    text="📊 Excel Viewer",
    font=ctk.CTkFont(size=18, weight="bold"),
    command=show_excel_treeview,
    fg_color="transparent",
    hover_color="#44475a",
    corner_radius=0,
    height=60,
    width=sidebar_width
)
excel_button.pack(pady=10, fill="x")



btn_summary = ctk.CTkButton(
    sidebar,
    text="📜 Summary",
    font=ctk.CTkFont(size=18, weight="bold"),
    command=on_summary_click,
    fg_color="transparent",
    hover_color="#44475a",
    corner_radius=0,
    height=60,
    width=sidebar_width
)
btn_summary.pack(fill="x", pady=(5, 10))

# Register buttons in the state tracker
sidebar_buttons.update({"charts": btn_charts, "summary": btn_summary})

# ----------- Move buttons to home frame -----------

button_frame = ctk.CTkFrame(home_frame, fg_color="transparent")
button_frame.pack(pady=(0, 30))

home_process_btn = ctk.CTkButton(
    button_frame,
    text="📁 Select Folder & Process",
    font=ctk.CTkFont(size=18),
    command=on_process_click,
    fg_color="#6272a4",
    hover_color="#7083c7",
    width=280,
    height=50,
    corner_radius=12
)
home_process_btn.grid(row=0, column=0, padx=10, pady=5)

home_open_btn = ctk.CTkButton(
    button_frame,
    text="📂 Open Result Folder",
    font=ctk.CTkFont(size=18),
    command=on_open_results_click,
    fg_color="#6272a4",
    hover_color="#7083c7",
    width=280,
    height=50,
    corner_radius=12
)
home_open_btn.grid(row=0, column=1, padx=10, pady=5)
chart_placeholder = ctk.CTkFrame(charts_frame, fg_color="transparent")
chart_placeholder.pack(fill="both", expand=True, padx=20, pady=10)
# ----------- Charts Frame Content -----------

def extract_and_plot():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        df = pd.read_excel(file_path, engine="openpyxl")

        # Parse the 4th column (index 3) as datetime with format dd-mm-yyyy
        df.iloc[:, 3] = pd.to_datetime(df.iloc[:, 3], format="%d-%m-%Y", errors='coerce')

        # Read and parse input dates from entries
        start_date_str = start_entry.get().strip()
        end_date_str = end_entry.get().strip()

        start_date = pd.to_datetime(start_date_str, format="%d-%m-%Y", errors='coerce')
        end_date = pd.to_datetime(end_date_str, format="%d-%m-%Y", errors='coerce')

        if pd.isnull(start_date) or pd.isnull(end_date):
            messagebox.showerror("Invalid Date", "Please enter valid dates in dd-mm-yyyy format.")
            return

        # Filter rows within date range
        df_filtered = df[(df.iloc[:, 3] >= start_date) & (df.iloc[:, 3] <= end_date)]

        if df_filtered.empty:
            messagebox.showinfo("No Data", "No data found in this date range.")
            return

        # Now parse column E (index 4) cells with pattern name+number (e.g. 'a10/b3/c5') within filtered rows only
        col_e = df_filtered.iloc[:, 4].dropna()

        sums = {}
        for cell in col_e:
            parts = str(cell).lower().split('/')
            for part in parts:
                match = re.match(r'([a-z]+)(\d+)', part.strip())
                if match:
                    name = match.group(1)
                    value = int(match.group(2))
                    sums[name] = sums.get(name, 0) + value

        if not sums:
            messagebox.showinfo("No Data", "No valid paper types found in this date range.")
            return

        names = list(sums.keys())
        values = list(sums.values())

        # Show results in a message box
        result_str = "\n".join([f"{name.upper()}: {val}" for name, val in zip(names, values)])
        messagebox.showinfo("Summed Values", result_str)

        # Plotting multiple chart types
        fig = plt.figure(figsize=(14, 10))
        gs = gridspec.GridSpec(3, 3)

        def make_bar(ax): ax.bar(names, values, color='skyblue'); ax.set_title("Bar Chart")
        def make_horizontal_bar(ax): ax.barh(names, values, color='salmon'); ax.set_title("Horizontal Bar")
        def make_pie(ax): ax.pie(values, labels=names, autopct='%1.1f%%'); ax.set_title("Pie Chart")
        def make_line(ax): ax.plot(names, values, marker='o'); ax.set_title("Line Chart")
        def make_area(ax): ax.fill_between(range(len(values)), values, alpha=0.5); ax.set_title("Area Chart"); ax.set_xticks(range(len(values))); ax.set_xticklabels(names)
        def make_scatter(ax): ax.scatter(names, values, c='green'); ax.set_title("Scatter Plot")
        def make_step(ax): ax.step(names, values, where='mid'); ax.set_title("Step Chart")
        def make_bar_stacked(ax): ax.bar(names, values, color='purple'); ax.set_title("Stacked (1 layer)")
        def make_donut(ax): 
            wedges, texts, autotexts = ax.pie(values, labels=names, autopct='%1.1f%%')
            center_circle = plt.Circle((0,0),0.70,fc='white')
            ax.add_artist(center_circle)
            ax.set_title("Donut Chart")

        chart_funcs = [make_bar, make_horizontal_bar, make_pie, make_line, make_area, make_scatter, make_step, make_bar_stacked, make_donut]

        for i, chart_func in enumerate(chart_funcs):
            ax = fig.add_subplot(gs[i // 3, i % 3])
            chart_func(ax)

        plt.tight_layout()
        plt.show()

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

# ---------- UI Setup (Date Form + Button) ----------
date_form_frame = ctk.CTkFrame(charts_frame, fg_color="transparent")
date_form_frame.pack(pady=20)

start_label = ctk.CTkLabel(date_form_frame, text="Start Date (dd-mm-yyyy):", font=ctk.CTkFont(size=16))
start_label.grid(row=0, column=0, padx=10)
start_entry = ctk.CTkEntry(date_form_frame, width=150)
start_entry.grid(row=0, column=1, padx=10)

end_label = ctk.CTkLabel(date_form_frame, text="End Date (dd-mm-yyyy):", font=ctk.CTkFont(size=16))
end_label.grid(row=1, column=0, padx=10, pady=(10, 0))
end_entry = ctk.CTkEntry(date_form_frame, width=150)
end_entry.grid(row=1, column=1, padx=10, pady=(10, 0))

plot_btn = ctk.CTkButton(
    charts_frame,
    text="📈 Generate Charts",
    font=ctk.CTkFont(size=18, weight="bold"),
    command=extract_and_plot,
    fg_color="#6272a4",
    hover_color="#7083c7",
    corner_radius=10,
    height=45
)
plot_btn.pack(pady=20)

# ----------- Summary Frame Content -----------
summary_label = ctk.CTkLabel(
    summary_frame,
    text="📘 How to Use the Application",
    font=ctk.CTkFont(size=26, weight="bold"),
    text_color="#8be9fd"
)
summary_label.pack(pady=30)

summary_box = ctk.CTkTextbox(
    summary_frame,
    width=900,
    height=500,
    font=ctk.CTkFont(size=16),
    wrap="word",
    fg_color="#1e1e2f",
    text_color="#f8f8f2",
    border_width=2,
    corner_radius=10
)
summary_box.pack(pady=20)

steps_text = """
👋 Welcome to your smart document processing assistant!

🔹 Step 1: Import Excel 📂  
Make sure your Excel file is named **"SOMMIER.xlsx"** and contains the essential data. It will be used to link information from other PDFs.

🔹 Step 2: Set the Results Folder ⚙️  
In the **Settings** tab, choose or create a folder where all output files will be saved. This folder will store processed data and generated Excel files.

🔹 Step 3: Load Your Folder with PDFs 🗃️  
Click on **Import Folder** and select the folder that includes:
- `fiche_de_cession.pdf`
- `fiche_produit.pdf`  
Ensure both files are present and are correctly named.

🔹 Step 4: Process and Review 📈  
After processing:
- Check results directly in the **output Excel** file inside your results folder.
- Use the **“Open Results Folder”** button to explore your outputs instantly.

📊 Graphs View (under Charts tab):
- Displays product distribution and frequencies based on your Excel data.
- Ensure column 4 contains the dates in format `dd-mm-yyyy` for proper visualization.

✅ That’s it! You’re now ready to automate and visualize your paperwork like a pro!
"""

summary_box.insert("0.0", steps_text)
summary_box.configure(state="disabled")



# ---------------- Your unchanged functions below ----------------

def ask_and_edit_paper_types(unique_names, parent):
    from tkinter import messagebox

    answer = messagebox.askyesno("Modifier les types de papier", "Voulez-vous modifier les types de papier ?", parent=parent)

    if not answer:
        return unique_names  # Continue as-is

    updated_names = unique_names.copy()

    def on_submit():
        for i, entry in enumerate(entries):
            updated_names[i] = entry.get()
        form.destroy()

    form = tk.Toplevel(parent)
    form.title("Modifier les types de papier")
    form.grab_set()  # Make the popup modal

    entries = []
    for i, name in enumerate(unique_names):
        tk.Label(form, text=f"Type {i+1}:").grid(row=i, column=0, padx=10, pady=5, sticky="e")
        entry = tk.Entry(form)
        entry.insert(0, name)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries.append(entry)

    submit_btn = tk.Button(form, text="Valider", command=on_submit)
    submit_btn.grid(row=len(unique_names), column=0, columnspan=2, pady=10)

    form.wait_window()  # Pause execution until the form is closed
    return updated_names

def open_result_folder(result_folder_path):
    folder = os.path.abspath(result_folder_path)
    if not os.path.exists(folder):
        os.makedirs(folder)
    system_name = platform.system()
    try:
        if system_name == "Windows":
            os.startfile(folder)
        elif system_name == "Darwin":
            os.system(f"open \"{folder}\"")
        else:
            os.system(f'xdg-open "{folder}"')
    except Exception as e:
        print(f"Could not open folder: {e}")

def show_done_popup():
    popup = ctk.CTkToplevel(app)
    popup.title("✅ Process Complete")
    popup.geometry("360x160")
    popup.resizable(False, False)
    popup.grab_set()
    popup.configure(fg_color="#212936")

    label = ctk.CTkLabel(popup, text="Processing Completed Successfully!", font=ctk.CTkFont(size=18, weight="bold"), text_color="#50fa7b")
    label.pack(pady=(30, 15))

    btn_open = ctk.CTkButton(popup, text="📂 Take me to the generated files", command=lambda: [open_result_folder(result_folder_path), popup.destroy()],
                             fg_color="#6272a4", hover_color="#7083c7", width=280, height=40, corner_radius=12)
    btn_open.pack(pady=(0, 12))

    btn_close = ctk.CTkButton(popup, text="Close", command=popup.destroy,
                              fg_color="#44475a", hover_color="#565f72", width=100, height=35, corner_radius=12)
    btn_close.pack()

def run_processing(folder_path):
    try:
        status_label.start("Processing")
        root = tk.Tk()
        root.withdraw()
        pytesseract.pytesseract.tesseract_cmd = r"C:\\tesseract\\tesseract.exe"
        fsommier = []
        fprod = []
        fcess = []
        names = []
        values = []

        files = os.listdir(folder_path)
        files = [f for f in files if os.path.isfile(os.path.join(folder_path, f))]

        for i in range(len(files)):
            x = backend.extractdata(folder_path + '/' + files[i])
            if x[0] == 'N/A':
                continue
            else:
                fprod.append(x)

        for i in range(len(files)):
            x = backend.extractcession(folder_path + '/' + files[i])
            if x:
                fcess.append(x)
                break

        fprod,fcess = backend.verifierfiles(fprod, fcess)

        for i in range(len(fprod)):
            unique_names, _ = backend.calculercomposition(fprod[i][3], fprod[i][2], fprod[i][4], fcess[i])
            names.extend(unique_names)
            _, total_values = backend.calculercomposition(fprod[i][3], fprod[i][2], fprod[i][4], fcess[i])
            values.extend(total_values)
        

        sums = defaultdict(float)
        for name, value in zip(names, values):
            sums[name] += value

        unique_names = list(sums.keys())
        summed_values = list(sums.values())
        unique_names = ask_and_edit_paper_types(unique_names,app)


        all_rows = []
        all_weight = []

        for i in range(len(unique_names)):
            df, weight = backend.get_and_process_rows([unique_names[i]], summed_values[i],file_path)
            if not df.empty:
                all_rows.append(df)
                all_weight.append(weight)

        if all_rows:
            final_df = pd.concat(all_rows, ignore_index=True)
            output_folder = os.path.join("C:/tesseract/resultat", datetime.today().strftime("%d-%m-%Y"))
            os.makedirs(output_folder, exist_ok=True)
            

            new_rows = []
            for _, row in final_df.iterrows():
                bureau = str(row['bureau']).zfill(3)
                regime = str(row['regime']).zfill(2)
                declaration = str(int(row['declaration'])).zfill(6)
                date_str = pd.to_datetime(row['date']).strftime('%d/%m/%Y')
                full_string = bureau + ' ' + regime + ' ' + declaration + " DU " + date_str
                new_rows.append(["AT:", full_string])

            matrix = []
            

            new_weight = []
            for sublist in all_weight:
                if isinstance(sublist, list) and len(sublist) > 1:
                    for val in sublist:
                        new_weight.append([val])
                else:
                    new_weight.append(sublist)

            all_weight = new_weight
            data=[]
            for i in range(len(fprod)):
                data.append([fprod[i][6],fprod[i][0],fprod[i][1],datetime.today().strftime("%d-%m-%Y"),fprod[i][3],fprod[i][2],fprod[i][4],backend.calculpoidsbrut(fprod[i][3],fprod[i][2],fprod[i][4]),backend.calculpoidsnet(backend.calculpoidsbrut(fprod[i][3],fprod[i][2],fprod[i][4]),15),fprod[i][5],backend.calculpoidsbrut(fprod[i][3],fprod[i][2],fprod[i][4])-float(fprod[i][5]),backend.calculpoidsdeclarer(fprod, fcess)])
            exportcalcul.append_row_to_excel(data,excel_file_path)   
            for i in range(len(final_df)):
                bureau = final_df.loc[i, 'bureau']
                nomenclature = final_df.loc[i, 'nomenclature']
                regime = final_df.loc[i, 'regime']
                declaration = final_df.loc[i, 'declaration']
                date = final_df.loc[i, 'date']
                qualite = final_df.loc[i, 'qualite'].split(' ')[0]
                poids_brut = all_weight[i]
                poids_net = backend.calculpoidsnet(poids_brut, 15)
                regime_declaration = f"{regime}/{declaration}"
                matrix.append([bureau, date, nomenclature, regime_declaration, poids_brut, poids_net])
                matrix.append([bureau, date, nomenclature, regime_declaration, poids_brut, poids_net])

            export.update_excel_layout(fprod[0][0], backend.sum_fourth_numbers(fcess), new_rows, True,result_folder_path)
            exportcalcul.update_excel_layout_full(fprod[0][0], matrix, backend.calculpoidsdeclarer(fprod, fcess),True,result_folder_path)

        status_label.stop()
        status_label.configure(text="✅ Done!")
        app.after(500, show_done_popup)

    except Exception as err:
        status_label.stop()
        status_label.configure(text="❌ Erreur détectée!")

        def show_error_popup(error=err):
            popup = ctk.CTkToplevel(app)
            popup.title("Erreur")
            popup.geometry("420x220")
            popup.resizable(False, False)
            popup.grab_set()
            popup.configure(fg_color="#2a2d36")

            label = ctk.CTkLabel(
                popup,
                text="Une erreur est survenue :",
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color="#ff5555"
            )
            label.pack(pady=(20, 10))

            error_textbox = ctk.CTkTextbox(
                popup,
                width=360,
                height=80,
                wrap="word"
            )
            error_textbox.insert("0.0", str(error))
            error_textbox.configure(
                state="disabled",
                fg_color="#1e1e1e",
                text_color="#f8f8f2",
                font=ctk.CTkFont(size=13)
            )
            error_textbox.pack(pady=(0, 10), padx=20)

            btn = ctk.CTkButton(
                popup,
                text="Fermer",
                command=popup.destroy,
                fg_color="#44475a",
                hover_color="#6272a4",
                corner_radius=10
            )
            btn.pack(pady=(5, 15))

        app.after(300, show_error_popup)


def open_tesseract_result_folder():
    folder = result_folder_path
    if not os.path.exists(folder):
        os.makedirs(folder)
    try:
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            os.system(f'open "{folder}"')
        else:
            os.system(f'xdg-open "{folder}"')
    except Exception as e:
        print(f"Could not open folder: {e}")

app.mainloop()
