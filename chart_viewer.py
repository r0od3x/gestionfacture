import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec

def extract_and_plot():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        return

    try:
        # Read Excel and get column E
        df = pd.read_excel(filepath)
        col_e = df.iloc[:, 4].dropna()  # Column E is index 4

        sums = {}

        for cell in col_e:
            parts = str(cell).lower().split('/')
            for part in parts:
                match = re.match(r'([a-z]+)(\d+)', part.strip())
                if match:
                    name = match.group(1)
                    value = int(match.group(2))
                    sums[name] = sums.get(name, 0) + value

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
        def make_donut(ax): wedges, texts, autotexts = ax.pie(values, labels=names, autopct='%1.1f%%'); center_circle = plt.Circle((0,0),0.70,fc='white'); ax.add_artist(center_circle); ax.set_title("Donut Chart")

        chart_funcs = [make_bar, make_horizontal_bar, make_pie, make_line, make_area, make_scatter, make_step, make_bar_stacked, make_donut]

        for i, chart_func in enumerate(chart_funcs):
            ax = fig.add_subplot(gs[i // 3, i % 3])
            chart_func(ax)

        plt.tight_layout()
        plt.show()

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

# GUI setup
root = tk.Tk()
root.title("Paper Type Analyzer")
root.geometry("420x200")

btn = tk.Button(root, text="Select Excel File and Generate Charts", command=extract_and_plot,
                font=("Arial", 12), bg="#44475a", fg="white", padx=20, pady=10)
btn.pack(pady=60)

root.mainloop()
