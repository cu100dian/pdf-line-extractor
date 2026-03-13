import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

data_rows = []

def parse_pdf(path):
    global data_rows

    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):

            text = page.extract_text()

            if not text:
                continue

            lines = text.split("\n")

            for line_num, line in enumerate(lines, start=1):

                row = [
                    os.path.basename(path),
                    page_num,
                    line_num,
                    line
                ]

                data_rows.append(row)

                table.insert(
                    "",
                    "end",
                    values=row
                )


def open_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])

    for file in files:
        parse_pdf(file)


def open_folder():
    folder = filedialog.askdirectory()

    for file in os.listdir(folder):
        if file.lower().endswith(".pdf"):
            parse_pdf(os.path.join(folder, file))


def export_csv():

    if not data_rows:
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV file", "*.csv")]
    )

    df = pd.DataFrame(
        data_rows,
        columns=["file", "page", "line", "text"]
    )

    df.to_csv(path, index=False)


def export_excel():

    if not data_rows:
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel file", "*.xlsx")]
    )

    df = pd.DataFrame(
        data_rows,
        columns=["file", "page", "line", "text"]
    )

    df.to_excel(path, index=False)


def drop(event):

    files = root.tk.splitlist(event.data)

    for file in files:
        if file.lower().endswith(".pdf"):
            parse_pdf(file)


# GUI

root = TkinterDnD.Tk()
root.title("PDF Line Extractor")
root.geometry("900x500")

# buttons

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Button(frame, text="Open PDF Files", command=open_files).pack(side="left", padx=5)
tk.Button(frame, text="Open Folder", command=open_folder).pack(side="left", padx=5)
tk.Button(frame, text="Export CSV", command=export_csv).pack(side="left", padx=5)
tk.Button(frame, text="Export Excel", command=export_excel).pack(side="left", padx=5)

# table

columns = ("file", "page", "line", "text")

table = ttk.Treeview(root, columns=columns, show="headings")

for col in columns:
    table.heading(col, text=col)
    table.column(col, anchor="w")

table.pack(fill="both", expand=True)

# drag & drop

table.drop_target_register(DND_FILES)
table.dnd_bind("<<Drop>>", drop)

root.mainloop()
