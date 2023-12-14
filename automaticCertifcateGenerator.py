# import datetime
# from pathlib import Path
# import pandas as pd
# from docxtpl import DocxTemplate

# base_dir = Path(__file__).parent
# template_path = base_dir/"CertificateFormate.docx"
# output_dir = base_dir/"Generated Certificate"
# excel_dir = base_dir/'students_name.xlsx'
# date = datetime.datetime.today()
# output_dir.mkdir(exist_ok=True)

# df = pd.read_excel(excel_dir,sheet_name="Sheet1")
# df["date"] = pd.to_datetime(df["date"]).dt.date
# df["date"] = date.strftime("%d-%m-%y")

# for records in df.to_dict(orient="records"):
#     document = DocxTemplate(template_path)
#     document.render(records)
#     document.save(output_dir/f"{records['name']}-Certificate.docx")

import datetime
from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import filedialog

def generate_certificates(template_path, excel_path, output_dir):
    base_dir = Path(__file__).parent
    template_path = Path(template_path)
    excel_path = Path(excel_path)
    output_dir = Path(output_dir)
    date = datetime.datetime.today()

    df = pd.read_excel(excel_path, sheet_name="Sheet1")
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["date"] = date.strftime("%d-%m-%y")

    output_dir.mkdir(exist_ok=True)

    for records in df.to_dict(orient="records"):
        document = DocxTemplate(template_path)
        document.render(records)
        document.save(output_dir / f"{records['name']}-Certificate.docx")

def browse_template():
    template_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    template_entry.delete(0, tk.END)
    template_entry.insert(0, template_path)

def browse_excel():
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    excel_entry.delete(0, tk.END)
    excel_entry.insert(0, excel_path)

def browse_output():
    output_dir = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_dir)

def generate_certificates_from_gui():
    template_path = template_entry.get()
    excel_path = excel_entry.get()
    output_dir = output_entry.get()
    generate_certificates(template_path, excel_path, output_dir)
    tk.messagebox.showinfo("Success", "Certificates generated successfully!")

# Create the main application window
app = tk.Tk()
app.title("Certificate Generator")

# Create and place widgets
tk.Label(app, text="Template Path:").grid(row=0, column=0, padx=10, pady=10)
template_entry = tk.Entry(app, width=50)
template_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=browse_template).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Excel Path:").grid(row=1, column=0, padx=10, pady=10)
excel_entry = tk.Entry(app, width=50)
excel_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=browse_excel).grid(row=1, column=2, padx=10, pady=10)

tk.Label(app, text="Output Directory:").grid(row=2, column=0, padx=10, pady=10)
output_entry = tk.Entry(app, width=50)
output_entry.grid(row=2, column=1, padx=10, pady=10)
tk.Button(app, text="Browse", command=browse_output).grid(row=2, column=2, padx=10, pady=10)

tk.Button(app, text="Generate Certificates", command=generate_certificates_from_gui).grid(row=3, column=1, pady=20)

# Start the Tkinter event loop
app.mainloop()
