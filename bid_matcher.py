import PyPDF2
import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF document."""
    text = ""
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text() + "\n"
    return text

def load_product_data(file_path):
    """Loads product specifications from an Excel/CSV file."""
    return pd.read_csv(file_path) if file_path.endswith(".csv") else pd.read_excel(file_path)

def match_product(extracted_specs, product_data):
    """Finds the best matching product based on extracted specifications."""
    for _, row in product_data.iterrows():
        if all(spec.lower() in row["Specifications"].lower() for spec in extracted_specs.split("\n") if spec.strip()):
            return row["Model"]
    return "Specifications not matched"

def select_bid_file():
    """Opens a file dialog to select the BID document."""
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        extracted_text.set(extract_text_from_pdf(file_path))

def search_match():
    """Finds and displays the matching product model."""
    selected_product = product_var.get()
    selected_brand = brand_var.get()
    if not selected_product or not selected_brand:
        messagebox.showwarning("Selection Error", "Please select a product and brand.")
        return
    
    matched_model = match_product(extracted_text.get(), product_data)
    result_label.config(text=f"Matched Model: {matched_model}")

# Load product data (Modify this path to match your product file)
product_data = load_product_data("products.xlsx")

# GUI Setup
root = tk.Tk()
root.title("BID Specification Matcher")

tk.Label(root, text="Select BID Document:").pack()
tk.Button(root, text="Browse", command=select_bid_file).pack()
extracted_text = tk.StringVar()

tk.Label(root, text="Select Product:").pack()
product_var = tk.StringVar()
product_dropdown = ttk.Combobox(root, textvariable=product_var, values=product_data["Product"].unique().tolist())
product_dropdown.pack()

tk.Label(root, text="Select Brand:").pack()
brand_var = tk.StringVar()
brand_dropdown = ttk.Combobox(root, textvariable=brand_var, values=product_data["Brand"].unique().tolist())
brand_dropdown.pack()

tk.Button(root, text="Search Match", command=search_match).pack()
result_label = tk.Label(root, text="")
result_label.pack()

root.mainloop()
