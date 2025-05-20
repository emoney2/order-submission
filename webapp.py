import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from tkinter import ttk
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import re
import datetime
import threading

# For Google Drive operations:
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# For formatting cells in Sheets:
from gspread_formatting import (CellFormat, format_cell_range,
                                DataValidationRule, BooleanCondition,
                                set_data_validation_for_cell_range)

# For image previews:
from PIL import Image, ImageTk, ImageOps

# For PDF rendering (PyMuPDF)
import fitz

# ===================== Global Variables & Helper Functions =====================

def update_image_preview(filepath, target_widget, width, height):
    """
    Opens the file at 'filepath.'
      - If it's a PDF, uses PyMuPDF (fitz) to render the first page.
      - Otherwise, uses Pillow (with exif_transpose to account for orientation).
    Then uses ImageOps.contain so that the entire image fits within (width, height) without cropping.
    The resized image is then centered on a white background.
    Returns an ImageTk.PhotoImage.
    """
    try:
        if filepath.lower().endswith('.pdf'):
            doc = fitz.open(filepath)
            page = doc.load_page(0)
            matrix = fitz.Matrix(2, 2)  # Adjust scale as needed.
            pix = page.get_pixmap(matrix=matrix)
            image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        else:
            image = Image.open(filepath)
            image = ImageOps.exif_transpose(image)
        resized = ImageOps.contain(image, (width, height))
        background = Image.new("RGB", (width, height), "white")
        paste_x = (width - resized.width) // 2
        paste_y = (height - resized.height) // 2
        background.paste(resized, (paste_x, paste_y))
        return ImageTk.PhotoImage(background)
    except Exception as e:
        print("Error in update_image_preview:", e)
        target_widget.delete("all")
        target_widget.create_text(width // 2, height // 2, text="Preview not available", fill="black")
        return None

SPREADSHEET = None
MATERIAL_INVENTORY_WS = None  # Cache for "Material Inventory" worksheet

MATERIALS_CACHE = None
FUR_COLORS_CACHE = None
COMPANIES_CACHE = None
PRODUCTS_CACHE = None
THREADS_CACHE = None

ORDERED_INVENTORY = []
ORDERED_TAB = None  # Will be set later

def find_exact_header_index(header_list, target):
    for i, header in enumerate(header_list):
        if header == target:
            return i
    return None

def get_next_business_day(date_obj):
    next_day = date_obj + datetime.timedelta(days=1)
    while next_day.weekday() >= 5:
        next_day += datetime.timedelta(days=1)
    return next_day

def column_letter(n):
    result = ""
    while n:
        n, rem = divmod(n-1, 26)
        result = chr(65 + rem) + result
    return result

# ----- Cache Clearing Functions -----
def clear_materials_cache():
    global MATERIALS_CACHE
    MATERIALS_CACHE = None

def clear_fur_colors_cache():
    global FUR_COLORS_CACHE
    FUR_COLORS_CACHE = None

def clear_companies_cache():
    global COMPANIES_CACHE
    COMPANIES_CACHE = None

def clear_products_cache():
    global PRODUCTS_CACHE
    PRODUCTS_CACHE = None

def clear_threads_cache():
    global THREADS_CACHE
    THREADS_CACHE = None

# ----- NEW: helper to append a company and clear cache -----
def create_new_company(data: dict):
    """
    Appends a new row to the 'Directory' sheet using the existing headers,
    then clears our local companies cache.
    """
    ws = open_sheet().worksheet("Directory")
    headers = ws.row_values(1)
    # Build a row in the same order as your sheet’s headers:
    row = [ data.get(h, "") for h in headers ]
    ws.append_row(row)
    clear_companies_cache()

# ===================== FilterableCombobox =====================

class FilterableCombobox(ttk.Combobox):
    def __init__(self, master=None, validate_callback=None, **kw):
        kw.setdefault("state", "normal")
        super().__init__(master, **kw)
        self._completion_list   = []
        self._validate_callback = validate_callback
        self._after_id = None
        self.bind("<KeyRelease>", self._on_keyrelease)
        # fire validate_callback once when the user leaves the field
        self.bind("<FocusOut>",   self._on_focusout)

    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list, key=str.lower)
        self["values"] = self._completion_list

    def _on_keyrelease(self, event):
        if event.keysym in ("BackSpace","Left","Right","Delete",
                            "Shift_L","Shift_R","Control_L","Control_R"):
            return
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(300, self._autocomplete)

    def _autocomplete(self):
        typed = self.get()
        if not typed:
            return
        cands = [c for c in self._completion_list if c.lower().startswith(typed.lower())]
        if cands:
            comp = cands[0]
            if comp.lower() != typed.lower():
                self.delete(0, tk.END)
                self.insert(0, comp)
                self.select_range(len(typed), tk.END)

    def _on_focusout(self, event):
        if self._validate_callback:
            self._validate_callback(event)



# ===================== Data Loading Functions =====================

def load_list_from_sheet(sheet, header_name):
    try:
        records = sheet.get_all_records()
        values = [record[header_name] for record in records if record.get(header_name)]
        return sorted(list(set(values)))
    except Exception as e:
        messagebox.showerror("Error", f"Could not load data for {header_name}: {e}")
        return []

def get_companies():
    global COMPANIES_CACHE
    if COMPANIES_CACHE is not None:
        return COMPANIES_CACHE
    ws = open_sheet().worksheet("Directory")
    COMPANIES_CACHE = load_list_from_sheet(ws, "Company Name")
    return COMPANIES_CACHE

def get_products():
    global PRODUCTS_CACHE
    if PRODUCTS_CACHE is not None:
        return PRODUCTS_CACHE
    ws = open_sheet().worksheet("Table")
    PRODUCTS_CACHE = load_list_from_sheet(ws, "Products")
    return PRODUCTS_CACHE

def get_materials():
    global MATERIALS_CACHE
    if MATERIALS_CACHE is not None:
        return MATERIALS_CACHE
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Materials" in headers:
        col_index = headers.index("Materials") + 1
        MATERIALS_CACHE = [m for m in ws.col_values(col_index)[1:] if m]
        return MATERIALS_CACHE
    else:
        messagebox.showerror("Error", "Header 'Materials' not found in Material Inventory")
        return []

def get_fur_colors():
    global FUR_COLORS_CACHE
    if FUR_COLORS_CACHE is not None:
        return FUR_COLORS_CACHE
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Fur Color" in headers:
        col_index = headers.index("Fur Color") + 1
        FUR_COLORS_CACHE = [f for f in ws.col_values(col_index)[1:] if f]
        return FUR_COLORS_CACHE
    else:
        messagebox.showerror("Error", "Header 'Fur Color' not found in Material Inventory")
        return []

def get_threads_inventory():
    global THREADS_CACHE
    if THREADS_CACHE is not None:
        return THREADS_CACHE
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Thread Colors" in headers:
        col_index = headers.index("Thread Colors") + 1
        THREADS_CACHE = [t for t in ws.col_values(col_index)[1:] if t]
        return THREADS_CACHE
    else:
        messagebox.showerror("Error", "Header 'Thread Colors' not found in Material Inventory")
        return []

# ===================== Google Drive Functions =====================

def get_drive_service():
    drive_scope = ['https://spreadsheets.google.com/feeds',
                   'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', drive_scope)
    service = build('drive', 'v3', credentials=creds)
    return service

def make_file_public(file_id):
    service = get_drive_service()
    permission = {'type': 'anyone', 'role': 'reader'}
    service.permissions().create(fileId=file_id, body=permission).execute()

def upload_file_to_drive(filepath, folder_id):
    service = get_drive_service()
    file_metadata = {'name': os.path.basename(filepath), 'parents': [folder_id]}
    media = MediaFileUpload(filepath, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()
    make_file_public(file['id'])
    return file['id'], file['webViewLink']

def create_drive_folder(folder_name, parent_id=None):
    service = get_drive_service()
    file_metadata = {'name': folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        file_metadata['parents'] = [parent_id]
    folder = service.files().create(body=file_metadata, fields='id, webViewLink').execute()
    make_file_public(folder['id'])
    return folder['id'], folder['webViewLink']

import gspread

# Path to your service-account JSON key
GSHEET_KEYFILE = r"C:\Users\eckar\Desktop\OrderEntry,Inventory,QR,PrintPDF\Keys\poetic-logic-454717-h2-3dd1bedb673d.json"

# These will be initialized on first use
SPREADSHEET = None
MATERIAL_INVENTORY_WS = None

# ===================== Google Sheets Setup =====================

def init_google_client():
    """
    Create a gspread client using a service account JSON.
    This client will auto-refresh indefinitely.
    """
    return gspread.service_account(filename=GSHEET_KEYFILE)

def open_sheet():
    """
    Open (and cache) the main spreadsheet by title.
    """
    global SPREADSHEET
    if SPREADSHEET is None:
        gc = init_google_client()
        SPREADSHEET = gc.open("JR and Co.")
    return SPREADSHEET

def get_material_inventory_ws():
    """
    Return (and cache) the 'Material Inventory' worksheet.
    """
    global MATERIAL_INVENTORY_WS
    if MATERIAL_INVENTORY_WS is None:
        MATERIAL_INVENTORY_WS = open_sheet().worksheet("Material Inventory")
    return MATERIAL_INVENTORY_WS

# ===================== Material Log Update =====================

def update_material_log(material, qty, order_status, date_stamp):
    ws_log = open_sheet().worksheet("Material Log")
    headers = ws_log.row_values(1)
    required = ["Material", "Yards", "IN/OUT", "O/R", "Date"]
    for r in required:
        if r not in headers:
            messagebox.showerror("Error", f"Header '{r}' not found in Material Log")
            return None
    mat_col = headers.index("Material") + 1
    yards_col = headers.index("Yards") + 1
    inout_col = headers.index("IN/OUT") + 1
    or_col = headers.index("O/R") + 1
    date_col = headers.index("Date") + 1
    materials_vals = ws_log.col_values(mat_col)
    new_row = len(materials_vals) + 1
    ws_log.update_cell(new_row, mat_col, material)
    ws_log.update_cell(new_row, yards_col, qty)
    ws_log.update_cell(new_row, inout_col, "IN")
    ws_log.update_cell(new_row, or_col, order_status)
    ws_log.update_cell(new_row, date_col, date_stamp)
    return new_row

# ===================== NewMaterialDialog =====================
class NewMaterialDialog(simpledialog.Dialog):
    """
    Dialog for adding a new material.
    Displays:
     - The material name (read-only)
     - Min Inv.
     - Reorder
     - Price
     - Unit (dropdown with options, e.g., "yard" and "sqft")
    """
    def __init__(self, parent, material, title="New Material Entry"):
        self.material = material
        self.min_inv = None
        self.reorder = None
        self.price = None
        self.unit = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text="Name:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        tk.Label(master, text=self.material).grid(row=0, column=1, padx=5, pady=5)
        tk.Label(master, text="Min Inv.:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.min_inv_entry = tk.Entry(master)
        self.min_inv_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(master, text="Reorder:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.reorder_entry = tk.Entry(master)
        self.reorder_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Label(master, text="Price:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.price_entry = tk.Entry(master)
        self.price_entry.grid(row=3, column=1, padx=5, pady=5)
        tk.Label(master, text="Unit:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.unit_cb = ttk.Combobox(master, state="readonly", values=["yard", "sqft"])
        self.unit_cb.grid(row=4, column=1, padx=5, pady=5)
        self.unit_cb.current(0)  # Default to "yard"
        return self.min_inv_entry

    def apply(self):
        self.min_inv = self.min_inv_entry.get().strip()
        self.reorder = self.reorder_entry.get().strip()
        self.price = self.price_entry.get().strip()
        self.unit = self.unit_cb.get().strip()

# ===================== NewCompanyDialog & NewProductDialog =====================

class NewCompanyDialog(simpledialog.Dialog):
    def __init__(self, parent, title="New Company Setup"):
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text="Company Name:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.company_entry = tk.Entry(master)
        self.company_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(master, text="Contact First Name:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.first_name_entry = tk.Entry(master)
        self.first_name_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(master, text="Contact Last Name:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.last_name_entry = tk.Entry(master)
        self.last_name_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Label(master, text="Contact Email Address:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.email_entry = tk.Entry(master)
        self.email_entry.grid(row=3, column=1, padx=5, pady=5)
        tk.Label(master, text="Street Address 1:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.address1_entry = tk.Entry(master)
        self.address1_entry.grid(row=4, column=1, padx=5, pady=5)
        tk.Label(master, text="Street Address 2:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        self.address2_entry = tk.Entry(master)
        self.address2_entry.grid(row=5, column=1, padx=5, pady=5)
        tk.Label(master, text="City:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        self.city_entry = tk.Entry(master)
        self.city_entry.grid(row=6, column=1, padx=5, pady=5)
        tk.Label(master, text="State:").grid(row=7, column=0, sticky="e", padx=5, pady=5)
        self.state_entry = tk.Entry(master)
        self.state_entry.grid(row=7, column=1, padx=5, pady=5)
        tk.Label(master, text="Zip Code:").grid(row=8, column=0, sticky="e", padx=5, pady=5)
        self.zip_entry = tk.Entry(master)
        self.zip_entry.grid(row=8, column=1, padx=5, pady=5)
        return self.company_entry

    def apply(self):
        self.result = {
            "Company Name": self.company_entry.get().strip(),
            "Contact First Name": self.first_name_entry.get().strip(),
            "Contact Last Name": self.last_name_entry.get().strip(),
            "Contact Email Address": self.email_entry.get().strip(),
            "Street Address 1": self.address1_entry.get().strip(),
            "Street Address 2": self.address2_entry.get().strip(),
            "City": self.city_entry.get().strip(),
            "State": self.state_entry.get().strip(),
            "Zip Code": self.zip_entry.get().strip()
        }

class NewProductDialog(simpledialog.Dialog):
    def __init__(self, parent, title="New Product Setup"):
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        tk.Label(master, text="Products:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.product_entry = tk.Entry(master)
        self.product_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Label(master, text="Print Times (1 Machine):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.print_times_entry = tk.Entry(master)
        self.print_times_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Label(master, text="How Many Products Per Yard:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.products_per_yard_entry = tk.Entry(master)
        self.products_per_yard_entry.grid(row=2, column=1, padx=5, pady=5)
        return self.product_entry

    def apply(self):
        self.result = {
            "Products": self.product_entry.get().strip(),
            "Print Times (1 Machine)": self.print_times_entry.get().strip(),
            "How Many Products Per Yard": self.products_per_yard_entry.get().strip()
        }

# ===================== New Material and Fur Color Processing =====================
def create_new_material_inventory(material, unit, min_inv, reorder, price):
    ws_inv = get_material_inventory_ws()
    headers = ws_inv.row_values(1)
    # New required headers include "Unit"
    required = ["Materials", "Unit", "Min. Inv.", "Reorder", "On Order", "Inventory", "Value"]
    for r in required:
        if r not in headers:
            messagebox.showerror("Error", f"Header '{r}' not found in Material Inventory")
            return
    mat_col = headers.index("Materials") + 1
    unit_col = headers.index("Unit") + 1
    min_inv_col = headers.index("Min. Inv.") + 1
    reorder_col = headers.index("Reorder") + 1
    on_order_col = headers.index("On Order") + 1
    inventory_col = headers.index("Inventory") + 1
    value_col = headers.index("Value") + 1
    new_row = len(ws_inv.col_values(mat_col)) + 1
    ws_inv.update_cell(new_row, mat_col, material)
    ws_inv.update_cell(new_row, unit_col, unit)
    ws_inv.update_cell(new_row, min_inv_col, min_inv)
    ws_inv.update_cell(new_row, reorder_col, reorder)
    col_letter_on_order = column_letter(on_order_col)
    formula_on_order = ws_inv.acell(f"{col_letter_on_order}2", value_render_option='FORMULA').value
    if formula_on_order:
        new_formula_on_order = re.sub(r'([A-Z]+)2\b', lambda m: f"{m.group(1)}{new_row}", formula_on_order)
        ws_inv.update_cell(new_row, on_order_col, new_formula_on_order)
    else:
        ws_inv.update_cell(new_row, on_order_col, "0")
    col_letter_inventory = column_letter(inventory_col)
    formula_inventory = ws_inv.acell(f"{col_letter_inventory}2", value_render_option='FORMULA').value
    if formula_inventory:
        new_formula_inventory = re.sub(r'([A-Z]+)2\b', lambda m: f"{m.group(1)}{new_row}", formula_inventory)
        ws_inv.update_cell(new_row, inventory_col, new_formula_inventory)
    else:
        ws_inv.update_cell(new_row, inventory_col, "0")
    inv_cell_addr = f"{column_letter(inventory_col)}{new_row}"
    on_order_cell_addr = f"{column_letter(on_order_col)}{new_row}"
    value_formula = f"={price}*({inv_cell_addr}+{on_order_cell_addr})"
    ws_inv.update_cell(new_row, value_col, value_formula)

def create_new_fur_color_inventory(fur_color, min_inv, reorder, price):
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    required = ["Fur Color", "Min Inv.", "Reorder.", "On Order.", "Inventory.", "Value."]
    for r in required:
        if r not in headers:
            messagebox.showerror("Error", f"Header '{r}' not found in Material Inventory")
            return
    fur_col = headers.index("Fur Color") + 1
    min_inv_col = headers.index("Min Inv.") + 1
    reorder_col = headers.index("Reorder.") + 1
    on_order_col = headers.index("On Order.") + 1
    inv_col = headers.index("Inventory.") + 1
    value_col = headers.index("Value.") + 1
    new_row = len(ws.col_values(fur_col)) + 1
    ws.update_cell(new_row, fur_col, fur_color)
    ws.update_cell(new_row, min_inv_col, min_inv)
    ws.update_cell(new_row, reorder_col, reorder)
    col_letter_on_order = column_letter(on_order_col)
    formula_on_order = ws.acell(f"{col_letter_on_order}2", value_render_option='FORMULA').value
    if formula_on_order:
        new_formula_on_order = re.sub(r'([A-Z]+)2\b', lambda m: f"{m.group(1)}{new_row}", formula_on_order)
        ws.update_cell(new_row, on_order_col, new_formula_on_order)
    else:
        ws.update_cell(new_row, on_order_col, "0")
    col_letter_inventory = column_letter(inv_col)
    formula_inventory = ws.acell(f"{col_letter_inventory}2", value_render_option='FORMULA').value
    if formula_inventory:
        new_formula_inventory = re.sub(r'([A-Z]+)2\b', lambda m: f"{m.group(1)}{new_row}", formula_inventory)
        ws.update_cell(new_row, inv_col, new_formula_inventory)
    else:
        ws.update_cell(new_row, inv_col, "0")
    inv_cell_addr = f"{column_letter(inv_col)}{new_row}"
    on_order_cell_addr = f"{column_letter(on_order_col)}{new_row}"
    value_formula = f"={price}*({inv_cell_addr}+{on_order_cell_addr})"
    ws.update_cell(new_row, value_col, value_formula)

# ===================== Inventory Material Processing =====================
def process_material_entry(material, qty, order_status, prompt=True):
    ws_inv = get_material_inventory_ws()
    headers = ws_inv.row_values(1)
    mat_col = headers.index("Materials") + 1 if "Materials" in headers else None
    fur_col = headers.index("Fur Color") + 1 if "Fur Color" in headers else None
    if mat_col is None or fur_col is None:
        messagebox.showerror("Error", "Required headers 'Materials' or 'Fur Color' not found.")
        return
    materials_list = ws_inv.col_values(mat_col)[1:]
    fur_list = ws_inv.col_values(fur_col)[1:]
    now = datetime.datetime.now()
    date_stamp = f"{now.month}/{now.day}/{now.year} {now.hour:02d}:{now.minute:02d}:{now.second:02d}"
    if material not in materials_list and material not in fur_list:
        response = messagebox.askyesno("New Material", f"Material '{material}' not found.\nDo you want to add it?")
        if response:
            dialog = NewMaterialDialog(tk._get_default_root(), material)
            if not dialog.min_inv or not dialog.reorder or not dialog.price or not dialog.unit:
                messagebox.showerror("Error", "All fields (Min Inv, Reorder, Price, Unit) are required.")
                return None
            if "fur" in material.lower():
                create_new_fur_color_inventory(material, dialog.min_inv, dialog.reorder, dialog.price)
                clear_fur_colors_cache()
            else:
                create_new_material_inventory(material, dialog.unit, dialog.min_inv, dialog.reorder, dialog.price)
                clear_materials_cache()
        else:
            messagebox.showinfo("Cancelled", "Entry submission cancelled.")
            return None
    return update_material_log(material, qty, order_status, date_stamp)

# ===================== Thread Inventory Processing =====================
def process_thread_entry(thread_color, qty):
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    if "Thread Colors" not in headers:
        messagebox.showerror("Error", "Header 'Thread Colors' not found in Material Inventory")
        return False
    t_col = headers.index("Thread Colors") + 1
    threads_list = ws.col_values(t_col)[1:]
    if any(t.strip().lower() == thread_color.strip().lower() for t in threads_list):
        update_thread_inventory_cell(thread_color, qty)
        return True
    else:
        response = messagebox.askyesno("New Thread Color", f"Thread color '{thread_color}' not found.\nDo you want to add it?")
        if response:
            dialog = NewEntryDialog(tk._get_default_root(), thread_color)
            if not dialog.min_inv or not dialog.reorder or not dialog.price:
                messagebox.showerror("Error", "Min Inv., Reorder, and Price values must be provided.")
                return False
            create_new_thread_inventory(thread_color, dialog.min_inv, dialog.reorder, dialog.price)
            update_thread_inventory_cell(thread_color, qty)
            return True
        else:
            messagebox.showinfo("Cancelled", "Thread entry submission cancelled.")
            return False

def update_thread_inventory_cell(thread_color, new_quantity):
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    records = ws.get_all_values()
    row_num = None
    if "Thread Colors" in headers:
        t_col = headers.index("Thread Colors")
        for i, row in enumerate(records[1:], start=2):
            if row[t_col].strip() == thread_color:
                row_num = i
                break
    if row_num is None:
        return False
    else:
        if "Inventory.." in headers:
            inv_col = headers.index("Inventory..") + 1
            ws.update_cell(row_num, inv_col, new_quantity)
            return True
        else:
            messagebox.showerror("Error", "Header 'Inventory..' not found in Material Inventory")
            return True

def create_new_thread_inventory(thread_color, min_inv, reorder, price):
    ws = get_material_inventory_ws()
    headers = ws.row_values(1)
    required = ["Thread Colors", "Min Inv..", "Reorder..", "On Order..", "Inventory..", "Value.."]
    for r in required:
        if r not in headers:
            messagebox.showerror("Error", f"Header '{r}' not found in Material Inventory")
            return
    t_col = headers.index("Thread Colors") + 1
    threads_vals = ws.col_values(t_col)
    new_row = len(threads_vals) + 1
    ws.update_cell(new_row, t_col, thread_color)
    ws.update_cell(new_row, headers.index("Min Inv..") + 1, min_inv)
    ws.update_cell(new_row, headers.index("Reorder..") + 1, reorder)
    inv_col = headers.index("Inventory..") + 1
    col_letter_inventory = column_letter(inv_col)
    formula_inventory = ws.acell(f"{col_letter_inventory}2", value_render_option='FORMULA').value
    if formula_inventory:
        new_formula_inventory = re.sub(r'([A-Z]+)2\b', lambda m: f"{m.group(1)}{new_row}", formula_inventory)
        ws.update_cell(new_row, inv_col, new_formula_inventory)
    else:
        ws.update_cell(new_row, inv_col, "0")
    on_order_col = headers.index("On Order..") + 1
    col_letter_on_order = column_letter(on_order_col)
    formula_on_order = ws.acell(f"{col_letter_on_order}2", value_render_option='FORMULA').value
    if formula_on_order:
        new_formula_on_order = re.sub(r'([A-Z]+)2\b', lambda m: f"{m.group(1)}{new_row}", formula_on_order)
        ws.update_cell(new_row, on_order_col, new_formula_on_order)
    else:
        ws.update_cell(new_row, on_order_col, "0")
    value_col = headers.index("Value..") + 1
    inv_cell_addr = f"{column_letter(inv_col)}{new_row}"
    on_order_cell_addr = f"{column_letter(on_order_col)}{new_row}"
    value_formula = f"={price}*({inv_cell_addr}+{on_order_cell_addr})"
    ws.update_cell(new_row, value_col, value_formula)

def post_thread_data(thread_color, thread_qty, thread_or, date_stamp):
    try:
        new_qty = float(thread_qty) * 16500
    except:
        new_qty = 0
    ws = open_sheet().worksheet("Thread Data")
    headers = ws.row_values(1)
    color_idx = find_exact_header_index(headers, "Color")
    length_idx = find_exact_header_index(headers, "Length (ft)")
    inout_idx = find_exact_header_index(headers, "IN/OUT")
    date_idx = find_exact_header_index(headers, "Date")
    or_idx = find_exact_header_index(headers, "O/R")
    if color_idx is None or length_idx is None or inout_idx is None:
        messagebox.showerror("Error", "Required headers ('Color', 'Length (ft)', 'IN/OUT') not found in Thread Data tab")
        return None
    rows = ws.col_values(color_idx + 1)
    new_row = len(rows) + 1
    ws.update_cell(new_row, color_idx + 1, thread_color)
    ws.update_cell(new_row, length_idx + 1, new_qty)
    ws.update_cell(new_row, inout_idx + 1, "IN")
    if date_idx is not None:
        ws.update_cell(new_row, date_idx + 1, date_stamp)
    if or_idx is not None:
        ws.update_cell(new_row, or_idx + 1, thread_or)
    return new_row

# ===================== Load Orders from Both Sheets =====================

def load_orders_from_sheets():
    orders = []
    try:
        ws_log = open_sheet().worksheet("Material Log")
        data = ws_log.get_all_values()
        if data and len(data) >= 2:
            headers = data[0]
            or_idx = find_exact_header_index(headers, "O/R")
            date_idx = find_exact_header_index(headers, "Date")
            material_idx = find_exact_header_index(headers, "Material")
            yards_idx = find_exact_header_index(headers, "Yards")
            for i, row in enumerate(data[1:], start=2):
                if or_idx is not None and len(row) > or_idx and row[or_idx].strip().lower() == "ordered":
                    orders.append({
                        "row": i,
                        "date": row[date_idx] if len(row) > date_idx else "",
                        "type": "Material",
                        "name": row[material_idx] if len(row) > material_idx else "",
                        "quantity": row[yards_idx] if len(row) > yards_idx else ""
                    })
    except Exception as e:
        pass
    try:
        ws_thread = open_sheet().worksheet("Thread Data")
        data = ws_thread.get_all_values()
        if data and len(data) >= 2:
            headers = data[0]
            or_idx = find_exact_header_index(headers, "O/R")
            date_idx = find_exact_header_index(headers, "Date")
            color_idx = find_exact_header_index(headers, "Color")
            length_idx = find_exact_header_index(headers, "Length (ft)")
            for i, row in enumerate(data[1:], start=2):
                if or_idx is not None and len(row) > or_idx and row[or_idx].strip().lower() == "ordered":
                    try:
                        qty_val = float(row[length_idx]) / 16500
                        display_qty = f"{qty_val:.2f} cones"
                    except:
                        display_qty = row[length_idx]
                    orders.append({
                        "row": i,
                        "date": row[date_idx] if len(row) > date_idx else "",
                        "type": "Thread",
                        "name": row[color_idx] if len(row) > color_idx else "",
                        "quantity": display_qty
                    })
    except Exception as e:
        pass
    return orders

# ===================== Inventory Ordered Tab =====================

class InventoryOrderedTab(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.tree = ttk.Treeview(self, columns=("Date", "Type", "Name", "Quantity"), show="headings", selectmode="extended")
        self.tree.heading("Date", text="Date", anchor="center", command=lambda: self.sort_column("Date", False))
        self.tree.heading("Type", text="Type", anchor="center", command=lambda: self.sort_column("Type", False))
        self.tree.heading("Name", text="Name", anchor="center", command=lambda: self.sort_column("Name", False))
        self.tree.heading("Quantity", text="Quantity", anchor="center", command=lambda: self.sort_column("Quantity", False))
        for col in ("Date", "Type", "Name", "Quantity"):
            self.tree.column(col, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True)
        receive_button = tk.Button(self, text="Receive", command=self.receive_orders)
        receive_button.pack(pady=5)
        self.refresh()

    def sort_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        if col == "Quantity":
            try:
                l = [(float(val.split()[0]), k) for val, k in l]
            except ValueError:
                pass
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        orders = load_orders_from_sheets()
        for order in orders:
            prefix = "T" if order["type"]=="Thread" else "M"
            item_id = f"{prefix}{order['row']}"
            self.tree.insert("", tk.END, iid=item_id,
                             values=(order["date"], order["type"], order["name"], order["quantity"]))

    def receive_orders(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showerror("Error", "No orders selected.")
            return
        now = datetime.datetime.now()
        new_date_stamp = f"{now.month}/{now.day}/{now.year} {now.hour:02d}:{now.minute:02d}:{now.second:02d}"
        for item in selected_items:
            values = self.tree.item(item)["values"]
            order_type = values[1]
            row_number = int(item[1:])  # Assuming IDs like "M23" or "T23"
            if order_type == "Material":
                ws = open_sheet().worksheet("Material Log")
                headers = ws.row_values(1)
                if "O/R" not in headers:
                    messagebox.showerror("Error", "Header 'O/R' not found in Material Log")
                    continue
                or_col = headers.index("O/R") + 1
                ws.update_cell(row_number, or_col, "Received")
                if "Date" in headers:
                    date_col = headers.index("Date") + 1
                    ws.update_cell(row_number, date_col, new_date_stamp)
            elif order_type == "Thread":
                ws = open_sheet().worksheet("Thread Data")
                headers = ws.row_values(1)
                if "O/R" not in headers:
                    messagebox.showerror("Error", "Header 'O/R' not found in Thread Data")
                    continue
                or_col = headers.index("O/R") + 1
                ws.update_cell(row_number, or_col, "Received")
                if "Date" in headers:
                    date_col = headers.index("Date") + 1
                    ws.update_cell(row_number, date_col, new_date_stamp)
            self.tree.delete(item)
        messagebox.showinfo("Success", "Selected orders marked as Received with updated timestamp.")
        self.refresh()

# ===================== Order Entry App =====================

class OrderEntryApp(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # ── Split pane: left for form, right for previews ──
        self.left_frame  = tk.Frame(self)
        self.right_frame = tk.Frame(self)

        # ── Order fields ──
        self.company_var       = tk.StringVar()
        self.referral_var      = tk.StringVar()
        self.design_var        = tk.StringVar()
        self.quantity_var      = tk.StringVar()
        self.product_var       = tk.StringVar()
        self.due_date_var      = tk.StringVar()
        self.price_var         = tk.StringVar()
        self.date_type_var     = tk.StringVar()
        self.material1_var     = tk.StringVar()
        self.material2_var     = tk.StringVar()
        self.material3_var     = tk.StringVar()
        self.material4_var     = tk.StringVar()
        self.material5_var     = tk.StringVar()
        self.back_material_var = tk.StringVar()
        self.fur_color_var     = tk.StringVar()
        self.backing_type_var  = tk.StringVar()
        self.notes_var         = tk.StringVar()
        self.prod_file_paths   = []
        self.print_file_paths  = []
        # ── watch company & fur text; fire check only once you’ve left the field ──
        self.company_var.trace_add("write", self._on_company_write)
        self.fur_color_var.trace_add("write",    self._on_fur_write)




        # ── Status bar (hidden until submit) ──
        self.loading_frame    = tk.Frame(self, relief="sunken", bd=1)
        self.loading_label    = tk.Label(self.loading_frame, text="Submitting order, please wait…")
        self.loading_progress = ttk.Progressbar(
            self.loading_frame, mode='determinate', maximum=100, value=0
        )
        self.loading_label.pack(side="left", padx=5, pady=2)
        self.loading_progress.pack(side="left", fill="x", expand=True, padx=5, pady=2)

        # ── Layout: reserve row 0 for form, row 1 for status bar ──
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0, minsize=30)
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=1)

        # ── Place the two main frames ──
        self.left_frame.grid (row=0, column=0, sticky="nsew")
        self.right_frame.grid(row=0, column=1, sticky="nsew")

        # ── Build sub‑panels ──
        self.build_order_details_frame   (self.left_frame)
        self.build_materials_frame       (self.left_frame)
        self.build_additional_info_frame (self.left_frame)
        self.build_file_previews         (self.right_frame)

        # ── Status bar placement ──
        self.loading_frame.grid_remove()
        self.loading_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.loading_frame.grid_remove()


    def date_type_key_handler(self, event):
        ch = event.char.lower()
        if ch == 'h':
            event.widget.set("Hard Date")
        elif ch == 's':
            event.widget.set("Soft Date")

    def show_loading(self):
        self.loading_progress.config(value=0)
        self.loading_frame.grid()        # make it visible
        self.update_idletasks()

    def hide_loading(self):
        self.loading_frame.grid_remove() # hide it again

    def build_order_details_frame(self, master):
        order_frame = tk.LabelFrame(master, text="Order Details", padx=10, pady=10)
        order_frame.pack(fill=tk.X, padx=5, pady=5)
        order_frame.grid_columnconfigure(1, weight=1)
        order_frame.grid_columnconfigure(3, weight=1)

            # Company Name
        tk.Label(order_frame, text="Company Name*") \
        .grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        self.company_cb = FilterableCombobox(
        order_frame,
        validate_callback=self.company_selected,
        textvariable=self.company_var
    )
        self.company_cb.set_completion_list(get_companies())
        self.company_cb.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        # still hook the dropdown‐select event if you like:
        self.company_cb.bind("<<ComboboxSelected>>", self.company_selected)



        tk.Label(order_frame, text="Design Name*").grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)
        self.design_entry = tk.Entry(order_frame, textvariable=self.design_var, state="readonly")
        self.design_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        tk.Label(order_frame, text="Quantity*").grid(row=2, column=0, sticky=tk.E, padx=5, pady=5)
        tk.Entry(order_frame, textvariable=self.quantity_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        tk.Label(order_frame, text="Product*").grid(row=2, column=2, sticky=tk.E, padx=5, pady=5)
        self.product_cb = FilterableCombobox(order_frame, textvariable=self.product_var)
        self.product_cb.set_completion_list(get_products())
        self.product_cb.grid(row=2, column=3, sticky="ew", padx=5, pady=5)
        self.product_cb.bind("<<ComboboxSelected>>", self.product_selected)

        tk.Label(order_frame, text="Due Date*").grid(row=3, column=0, sticky=tk.E, padx=5, pady=5)
        tk.Entry(order_frame, textvariable=self.due_date_var).grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        tk.Label(order_frame, text="Price*").grid(row=3, column=2, sticky=tk.E, padx=5, pady=5)
        tk.Entry(order_frame, textvariable=self.price_var).grid(row=3, column=3, sticky="ew", padx=5, pady=5)

        tk.Label(order_frame, text="Hard Date/Soft Date*").grid(row=4, column=0, sticky=tk.E, padx=5, pady=5)
        self.date_type_cb = ttk.Combobox(
            order_frame,
            textvariable=self.date_type_var,
            state="readonly",
            values=["Hard Date", "Soft Date"],
            width=15
        )
        self.date_type_cb.grid(row=4, column=1, sticky="ew", padx=5, pady=5)
        self.date_type_cb.bind("<Key>", self.date_type_key_handler)

        tk.Label(order_frame, text="Referral").grid(row=4, column=2, sticky=tk.E, padx=5, pady=5)
        self.referral_cb = FilterableCombobox(order_frame, textvariable=self.referral_var)
        self.referral_cb.set_completion_list(get_companies())
        self.referral_cb.grid(row=4, column=3, sticky="ew", padx=5, pady=5)

    def build_materials_frame(self, master):
        materials_frame = tk.LabelFrame(master, text="Materials", padx=10, pady=10)
        materials_frame.pack(fill=tk.X, padx=5, pady=5)
        for i in range(4):
            materials_frame.grid_columnconfigure(i, weight=1)

        tk.Label(materials_frame, text="Material 1*").grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        self.material1_cb = FilterableCombobox(
    materials_frame,
    validate_callback=lambda e: self.material_selected(self.material1_var),
    textvariable=self.material1_var
)
        self.material1_cb.set_completion_list(self.get_union_material_list())
        self.material1_cb.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.material1_cb.bind("<<ComboboxSelected>>", lambda e: self.material_selected(self.material1_var))


        tk.Label(materials_frame, text="Material 2").grid(row=1, column=0, sticky=tk.E, padx=5, pady=5)
        self.material2_cb = FilterableCombobox(
    materials_frame,
    validate_callback=lambda e: self.material_selected(self.material2_var),
    textvariable=self.material2_var
)
        self.material2_cb.set_completion_list(self.get_union_material_list())
        self.material2_cb.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        self.material2_cb.bind("<<ComboboxSelected>>", lambda e: self.material_selected(self.material2_var))


        tk.Label(materials_frame, text="Material 3").grid(row=2, column=0, sticky=tk.E, padx=5, pady=5)
        self.material3_cb = FilterableCombobox(
    materials_frame,
    validate_callback=lambda e: self.material_selected(self.material3_var),
    textvariable=self.material3_var
)
        self.material3_cb.set_completion_list(self.get_union_material_list())
        self.material3_cb.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.material3_cb.bind("<<ComboboxSelected>>", lambda e: self.material_selected(self.material3_var))


        tk.Label(materials_frame, text="Material 4").grid(row=0, column=2, sticky=tk.E, padx=5, pady=5)
        self.material4_cb = FilterableCombobox(
    materials_frame,
    validate_callback=lambda e: self.material_selected(self.material4_var),
    textvariable=self.material4_var
)
        self.material4_cb.set_completion_list(self.get_union_material_list())
        self.material4_cb.grid(row=0, column=3, sticky="ew", padx=5, pady=5)
        self.material4_cb.bind("<<ComboboxSelected>>", lambda e: self.material_selected(self.material4_var))


        tk.Label(materials_frame, text="Material 5").grid(row=1, column=2, sticky=tk.E, padx=5, pady=5)
        self.material5_cb = FilterableCombobox(
    materials_frame,
    validate_callback=lambda e: self.material_selected(self.material5_var),
    textvariable=self.material5_var
)
        self.material5_cb.set_completion_list(self.get_union_material_list())
        self.material5_cb.grid(row=1, column=3, sticky="ew", padx=5, pady=5)
        self.material5_cb.bind("<<ComboboxSelected>>", lambda e: self.material_selected(self.material5_var))


        tk.Label(materials_frame, text="Back Material").grid(row=2, column=2, sticky=tk.E, padx=5, pady=5)
        self.back_material_cb = FilterableCombobox(
    materials_frame,
    validate_callback=lambda e: self.material_selected(self.back_material_var),
    textvariable=self.back_material_var
)
        self.back_material_cb.set_completion_list(self.get_union_material_list())
        self.back_material_cb.grid(row=2, column=3, sticky="ew", padx=5, pady=5)
        self.back_material_cb.bind("<<ComboboxSelected>>", lambda e: self.material_selected(self.back_material_var))



    def build_additional_info_frame(self, master):
        additional_frame = tk.LabelFrame(master, text="Additional Info", padx=10, pady=10)
        additional_frame.pack(fill=tk.X, padx=5, pady=5)
        for i in range(5):
            additional_frame.grid_columnconfigure(i, weight=1)

            # Fur Color
        tk.Label(additional_frame, text="Fur Color*") \
        .grid(row=0, column=0, sticky=tk.E, padx=5, pady=5)
        self.fur_color_cb = FilterableCombobox(
        additional_frame,
        validate_callback=self.fur_color_selected,
        textvariable=self.fur_color_var
    )
        self.fur_color_cb.set_completion_list(get_fur_colors())
        self.fur_color_cb.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        # still hook dropdown‐select if desired
        self.fur_color_cb.bind("<<ComboboxSelected>>", self.fur_color_selected)




        # Backing Type
        tk.Label(additional_frame, text="Backing Type*").grid(row=0, column=2, sticky=tk.E, padx=5, pady=5)
        self.backing_type_cb = FilterableCombobox(additional_frame, textvariable=self.backing_type_var)
        self.backing_type_cb.set_completion_list(["Cut Away", "Tear Away"])
        self.backing_type_cb.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        # Notes
        tk.Label(additional_frame, text="Notes").grid(row=1, column=0, sticky=tk.NE, padx=5, pady=5)
        tk.Entry(additional_frame, textvariable=self.notes_var) \
            .grid(row=1, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        # Production File
        tk.Label(additional_frame, text="Production File").grid(row=2, column=0, sticky=tk.E, padx=5, pady=5)
        tk.Button(additional_frame, text="Browse", command=self.upload_prod_file) \
            .grid(row=2, column=1, sticky="w", padx=5, pady=5)
        self.prod_file_label = tk.Label(additional_frame, text="", fg="blue")
        self.prod_file_label.grid(row=2, column=2, sticky="w", padx=5, pady=5)
        self.prod_file_close = tk.Button(additional_frame, text="X", command=self.clear_prod_file, fg="red")
        self.prod_file_close.grid(row=2, column=3, sticky="w", padx=5, pady=5)

        # Print File(s)
        tk.Label(additional_frame, text="Print File(s)").grid(row=3, column=0, sticky=tk.E, padx=5, pady=5)
        tk.Button(additional_frame, text="Browse", command=self.upload_print_files) \
            .grid(row=3, column=1, sticky="w", padx=5, pady=5)
        self.print_file_label = tk.Label(additional_frame, text="", fg="blue")
        self.print_file_label.grid(row=3, column=2, sticky="w", padx=5, pady=5)
        self.print_file_close = tk.Button(additional_frame, text="X", command=self.clear_print_files, fg="red")
        self.print_file_close.grid(row=3, column=3, sticky="w", padx=5, pady=5)

        # Submit Button
        tk.Button(additional_frame, text="Submit Order", command=self.submit_order) \
            .grid(row=4, column=0, columnspan=5, pady=10)



    def build_file_previews(self, master):
        # Create preview frames for Production and Print files with fixed size 250x250.
        self.prod_preview_frame = tk.LabelFrame(master, text="Production File Preview", padx=5, pady=5, width=250, height=250)
        self.prod_preview_frame.grid_propagate(False)
        self.prod_preview_frame.pack(padx=5, pady=5)
        self.print_preview_frame = tk.LabelFrame(master, text="Print File Preview", padx=5, pady=5, width=250, height=250)
        self.print_preview_frame.grid_propagate(False)
        self.print_preview_frame.pack(padx=5, pady=5)

    def update_previews(self, container, filepaths):
        # Clear container first.
        for widget in container.winfo_children():
            widget.destroy()
        n = len(filepaths)
        if n == 0:
            label = tk.Label(container, text="No File Selected", fg="black")
            label.pack(expand=True)
            return
        n = min(n, 4)
        container.config(width=250, height=250)
        if n == 1:
            rows, cols = 1, 1
        elif n == 2:
            rows, cols = 2, 1
        else:
            rows, cols = 2, 2
        canvas_width = 250 // cols
        canvas_height = 250 // rows
        for col in range(cols):
            container.grid_columnconfigure(col, weight=1)
        for row in range(rows):
            container.grid_rowconfigure(row, weight=1)
        for i in range(n):
            r = i // cols
            c = i % cols
            canvas = tk.Canvas(container, bg="white")
            canvas.grid(row=r, column=c, sticky="nsew")
            canvas.config(width=canvas_width, height=canvas_height)
            img = update_image_preview(filepaths[i], canvas, canvas_width, canvas_height)
            if img:
                canvas.create_image(canvas_width//2, canvas_height//2, image=img)
                canvas.image = img

    def update_prod_previews(self, filepaths):
        self.update_previews(self.prod_preview_frame, filepaths)

    def update_print_previews(self, filepaths):
        self.update_previews(self.print_preview_frame, filepaths)

    # ----------------- Modified Functions for Multiple File Selection -----------------
    # inside class OrderEntryApp, at the same indent as the old method
    def upload_prod_file(self):
        new_filepaths = filedialog.askopenfilenames(title="Select Production File(s)")
        if not new_filepaths:
            return
        self.prod_file_paths.extend(new_filepaths)

        # truncate to 20 chars
        display_names = []
        for fp in self.prod_file_paths:
            name = os.path.basename(fp)
            if len(name) > 20:
                name = name[:17] + "..."
            display_names.append(name)

        self.prod_file_label.config(text=", ".join(display_names))
        messagebox.showinfo("Selected", ", ".join(display_names))

        # set design name from first file
        truncated = display_names[0]                  # e.g. "averylongfilena..."
        design_base = os.path.splitext(truncated)[0]  # strip any extension
        self.design_var.set(design_base)

        self.update_prod_previews(self.prod_file_paths)


    def upload_print_files(self):
        new_filepaths = filedialog.askopenfilenames(title="Select Print File(s)")
        if not new_filepaths:
            return
        self.print_file_paths.extend(new_filepaths)

        # truncate to 20 chars
        display_names = []
        for fp in self.print_file_paths:
            name = os.path.basename(fp)
            if len(name) > 20:
                name = name[:17] + "..."
            display_names.append(name)

        self.print_file_label.config(text=", ".join(display_names))
        messagebox.showinfo("Selected", ", ".join(display_names))

        self.update_print_previews(self.print_file_paths)

    def _on_company_write(self, *args):
        # schedule a focus‐aware check once the event loop is idle
        self.after_idle(self._check_company_focus)

    def _check_company_focus(self):
        # if the combobox no longer has focus, run the pop‐up
        if self.focus_get() != self.company_cb:
            self.company_selected(None)

    def _on_fur_write(self, *args):
        self.after_idle(self._check_fur_focus)

    def _check_fur_focus(self):
        if self.focus_get() != self.fur_color_cb:
            self.fur_color_selected(None)


    # -------------------------------------------------------------------------------------

    def clear_prod_file(self):
        self.prod_file_paths = []
        self.prod_file_label.config(text="")
        for widget in self.prod_preview_frame.winfo_children():
            widget.destroy()
        label = tk.Label(self.prod_preview_frame, text="No File Selected", fg="black")
        label.pack(expand=True)

    def clear_print_files(self):
        self.print_file_paths = []
        self.print_file_label.config(text="")
        for widget in self.print_preview_frame.winfo_children():
            widget.destroy()
        label = tk.Label(self.print_preview_frame, text="No File Selected", fg="black")
        label.pack(expand=True)

    def company_selected(self, event):
        name = self.company_var.get().strip()
        # if it’s empty or already known, nothing to do
        if not name or name in get_companies():
            return

        resp = messagebox.askyesno("New Company", f"'{name}' not found. Add it now?")
        if not resp:
            # cancel → clear and abort
            self.company_var.set("")
            raise Exception("Please provide a valid Company Name.")

        dlg = NewCompanyDialog(self.master)
        if not dlg.result:
            # user hit Cancel or left fields blank
            self.company_var.set("")
            raise Exception("All company fields are required.")
        
        # 1️⃣ actually write the new row…
        create_new_company(dlg.result)

        # 2️⃣ clear & re-populate the combobox cache
        clear_companies_cache()
        self.company_cb.set_completion_list(get_companies())

        # 3️⃣ set it to what the user entered
        self.company_var.set(dlg.result["Company Name"])



    def update_companies(self, new_company):
        self.company_cb.set_completion_list(get_companies())
        if new_company:
            self.company_var.set(new_company)

    def product_selected(self, event):
        # stub left intentionally—no pop-up for Product currently
        pass

    def update_products(self, new_product):
        self.product_cb.set_completion_list(get_products())
        if new_product:
            self.product_var.set(new_product)

    def material_selected(self, var):
        mat = var.get().strip()
        union = self.get_union_material_list()
        if not mat or mat in union:
            return

        resp = messagebox.askyesno("New Material", f"'{mat}' not found.  Add it now?")
        if resp:
            dlg = NewMaterialDialog(self.master, mat)
            if dlg.material:
                create_new_material_inventory(
                    dlg.material, dlg.unit, dlg.min_inv, dlg.reorder, dlg.price
                )
                clear_materials_cache()
                clear_fur_colors_cache()
                self.refresh_material_dropdowns()
                var.set(mat)
        else:
            var.set("")
            raise Exception(f"Please provide a valid value for material '{mat}'.")


    def fur_color_selected(self, event):
        fur = self.fur_color_var.get().strip()
        if not fur or fur in get_fur_colors():
            return

        resp = messagebox.askyesno("New Fur Color", f"'{fur}' not found. Add it now?")
        if not resp:
            self.fur_color_var.set("")
            raise Exception("Please provide a valid Fur Color.")

        dlg = NewMaterialDialog(self.master, fur, title="New Fur Color")
        # Note: NewMaterialDialog always gives you dlg.material==fur
        # but stores min_inv, reorder, price, unit on the dialog object.
        if not (dlg.min_inv and dlg.reorder and dlg.price):
            self.fur_color_cb.configure(background="yellow")
            raise Exception("All fields (Min Inv, Reorder, Price) are required.")

        # write it into the sheet (no unit column for fur)
        create_new_fur_color_inventory(
            dlg.material,
            dlg.min_inv,
            dlg.reorder,
            dlg.price
        )

        # clear & repopulate
        clear_fur_colors_cache()
        self.fur_color_cb.set_completion_list(get_fur_colors())

        # and set the var so the popup won’t re-fire
        self.fur_color_var.set(dlg.material)
        self.fur_color_cb.configure(background="white")



    def get_union_material_list(self):
        return sorted(list(set(get_materials()).union(set(get_fur_colors()))))

    def refresh_material_dropdowns(self):
        union_list = self.get_union_material_list()
        for cb in [
            self.material1_cb, self.material2_cb, self.material3_cb,
            self.material4_cb, self.material5_cb, self.back_material_cb
        ]:
            cb.set_completion_list(union_list)

    def validate_order_materials(self):
        valid = True
        required = ["Material 1"]
        if self.product_var.get() in ["Driver Full", "Hybrid Full", "Fairway Full"]:
            required.append("Back Material")
        fields = [
            ("Material 1", self.material1_var, self.material1_cb, "Material 1" in required),
            ("Material 2", self.material2_var, self.material2_cb, False),
            ("Material 3", self.material3_var, self.material3_cb, False),
            ("Material 4", self.material4_var, self.material4_cb, False),
            ("Material 5", self.material5_var, self.material5_cb, False),
            ("Back Material", self.back_material_var, self.back_material_cb, "Back Material" in required)
        ]
        for fname, var, widget, req in fields:
            mat = var.get().strip()
            if req and not mat:
                widget.configure(background="yellow")
                valid = False
            else:
                widget.configure(background="white")
        return valid

    def validate_order_fur_color(self):
        valid = True
        fur = self.fur_color_var.get().strip()
        if not fur:
            self.fur_color_cb.configure(background="yellow")
            valid = False
        else:
            self.fur_color_cb.configure(background="white")
        return valid

    def process_new_materials(self):
        material_fields = [
            ("Material 1", self.material1_var, self.material1_cb),
            ("Material 2", self.material2_var, self.material2_cb),
            ("Material 3", self.material3_var, self.material3_cb),
            ("Material 4", self.material4_var, self.material4_cb),
            ("Material 5", self.material5_var, self.material5_cb),
            ("Back Material", self.back_material_var, self.back_material_cb)
        ]
        union_list = self.get_union_material_list()
        for field_name, var, widget in material_fields:
            mat = var.get().strip()
            if mat and mat not in union_list:
                resp = messagebox.askyesno("New Entry", f"'{mat}' not found for {field_name}. Add it?")
                if resp:
                    dlg = NewMaterialDialog(tk._get_default_root(), mat)
                    if not (dlg.unit and dlg.min_inv and dlg.reorder and dlg.price):
                        widget.configure(background="yellow")
                        raise Exception(f"Missing info for {field_name}.")
                    create_new_material_inventory(
                        mat, dlg.unit, dlg.min_inv, dlg.reorder, dlg.price
                    )
                    clear_materials_cache()
                    union_list = self.get_union_material_list()
                    widget.configure(background="white")
                else:
                    widget.configure(background="yellow")
                    raise Exception(f"Please provide a valid value for {field_name}.")

    def process_new_fur_color(self):
        fur = self.fur_color_var.get().strip()
        fur_list = get_fur_colors()
        if fur and fur not in fur_list:
            resp = messagebox.askyesno("New Entry", f"Fur color '{fur}' not found. Add it?")
            if resp:
                dlg = NewMaterialDialog(tk._get_default_root(), fur, title="New Fur Color Setup")
                if not (dlg.unit and dlg.min_inv and dlg.reorder and dlg.price):
                    self.fur_color_cb.configure(background="yellow")
                    raise Exception("Missing info for Fur Color.")
                create_new_fur_color_inventory(
                    fur,
                    dlg.min_inv,
                    dlg.reorder,
                    dlg.price
                )

                clear_fur_colors_cache()
                self.fur_color_cb.configure(background="white")
            else:
                self.fur_color_cb.configure(background="yellow")
                raise Exception("Please provide a valid Fur Color.")


    def submit_order(self):
        # ── force new-entry pop-ups for any typed-but-unrecognized values ──
        self.company_selected(None)
        self.process_new_materials()
        self.process_new_fur_color()


        # ------------------ show the bottom progress bar ------------------
        self.show_loading()
        try:
            self.loading_progress.config(value=10)
            self.update_idletasks()

            # ---- parse due date ----
            due_date_str = self.due_date_var.get().strip()
            if not due_date_str:
                raise Exception("Due Date is required.")
            try:
                current_year = datetime.date.today().year
                due_date = datetime.datetime.strptime(due_date_str, '%m/%d').date().replace(year=current_year)
                if due_date < datetime.date.today():
                    due_date = due_date.replace(year=current_year + 1)
            except ValueError:
                raise Exception("Invalid due date format. Use m/dd or mm/dd.")
            if due_date.weekday() >= 5:
                next_business = get_next_business_day(due_date)
                resp = messagebox.askyesno(
                    "Weekend Detected",
                    f"Due date ({due_date.strftime('%m/%d/%Y')}) falls on a weekend.\n"
                    f"Use next business day ({next_business.strftime('%m/%d/%Y')})?"
                )
                if resp:
                    due_date = next_business
                self.due_date_var.set(due_date.strftime("%m/%d"))

            self.loading_progress.config(value=20)
            self.update_idletasks()

            # ---- required fields ----
            for label, var in [
                ("Company Name", self.company_var.get()),
                ("Design Name", self.design_var.get()),
                ("Quantity", self.quantity_var.get()),
                ("Product", self.product_var.get()),
                ("Due Date", self.due_date_var.get()),
                ("Price", self.price_var.get()),
                ("Backing Type", self.backing_type_var.get()),
                ("Hard/Soft Date", self.date_type_var.get()),
            ]:
                if not var:
                    raise Exception(f"Provide a valid {label}.")

            # ---- still validate materials/fur for coloring, but do NOT log them here ----
            if not self.validate_order_materials() or not self.validate_order_fur_color():
                raise Exception("Some required material or fur fields are missing.")

            self.loading_progress.config(value=30)
            self.update_idletasks()

            # ---- gather order data ----
            order_data = {
                "Company Name": self.company_var.get(),
                "Design Name": self.design_var.get().strip(),
                "Due Date": self.due_date_var.get().strip(),
                "Quantity": self.quantity_var.get().strip(),
                "Product": self.product_var.get(),
                "Price": self.price_var.get(),
                "Material 1": self.material1_var.get(),
                "Material 2": self.material2_var.get(),
                "Material 3": self.material3_var.get(),
                "Material 4": self.material4_var.get(),
                "Material 5": self.material5_var.get(),
                "Back Material": self.back_material_var.get(),
                "Fur Color": self.fur_color_var.get(),
                "Hard Date/Soft Date": self.date_type_var.get(),
                "Notes": self.notes_var.get().strip(),
            }
            mapping = {
                "Company Name": "Company Name",
                "Design Name":  "Design",
                "Due Date":     "Due Date",
                "Quantity":     "Quantity",
                "Product":      "Product",
                "Price":        "Price",
                "Material 1":   "Material1",
                "Material 2":   "Material2",
                "Material 3":   "Material3",
                "Material 4":   "Material4",
                "Material 5":   "Material5",
                "Back Material":"Back Material",
                "Fur Color":    "Fur Color",
                "Hard Date/Soft Date": "Hard Date/Soft Date",
                "Notes":        "Notes"
            }

            self.loading_progress.config(value=40)
            self.update_idletasks()

            # ---- new row + folder ----
            prod_ws = open_sheet().worksheet("Production Orders")
            next_row = len(prod_ws.col_values(1)) + 1
            folder_id, _ = create_drive_folder(str(next_row))
            self.loading_progress.config(value=50)
            self.update_idletasks()
            # ---- upload production files (first‐file link only) ----
            prod_file_link = ""
            if self.prod_file_paths:
                # upload first and keep its link
                _, prod_file_link = upload_file_to_drive(self.prod_file_paths[0], folder_id)
                # upload the rest silently
                for fp in self.prod_file_paths[1:]:
                    upload_file_to_drive(fp, folder_id)
            self.loading_progress.config(value=60)
            self.update_idletasks()
            # ---- upload print files (folder link) ----
            print_file_link = ""
            if self.print_file_paths:
                print_folder_id, print_folder_link = create_drive_folder("Print Files", parent_id=folder_id)
                for fp in self.print_file_paths:
                    upload_file_to_drive(fp, print_folder_id)
            self.loading_progress.config(value=65)
            self.update_idletasks()

            # ---- write to sheet ----
            headers = prod_ws.row_values(1)

            # Print column
            pc = find_exact_header_index(headers, "Print")
            if pc is not None:
                prod_ws.update_cell(next_row, pc+1, "PRINT" if self.print_file_paths else "NO")
            self.loading_progress.config(value=70)
            self.update_idletasks()

            # Reenter Order
            rio = find_exact_header_index(headers, "Reenter Order")
            if rio is not None:
                cell = f"{column_letter(rio+1)}{next_row}"
                rule = DataValidationRule(BooleanCondition("BOOLEAN", []), showCustomUi=True)
                set_data_validation_for_cell_range(prod_ws, cell, rule)
                prod_ws.update_cell(next_row, rio+1, "FALSE")
            self.loading_progress.config(value=75)
            self.update_idletasks()

            # Main fields
            for field, col_name in mapping.items():
                idx = find_exact_header_index(headers, col_name)
                if idx is not None:
                    prod_ws.update_cell(next_row, idx+1, order_data[field])
            self.loading_progress.config(value=80)
            self.update_idletasks()

            # Order #
            oni = find_exact_header_index(headers, "Order #")
            if oni is not None:
                if next_row == 2:
                    num = 1
                else:
                    prev = prod_ws.cell(next_row-1, oni+1).value
                    try: num = int(prev) + 1
                    except: num = 1
                prod_ws.update_cell(next_row, oni+1, num)

            # Timestamp
            now = datetime.datetime.now()
            ts = f"{now.month}/{now.day}/{now.year} {now.hour:02d}:{now.minute:02d}:{now.second:02d}"
            di = find_exact_header_index(headers, "Date")
            if di is not None:
                prod_ws.update_cell(next_row, di+1, ts)
            self.loading_progress.config(value=85)
            self.update_idletasks()

            # Image & Print Files links
            ic = find_exact_header_index(headers, "Image")
            if ic is not None:
                prod_ws.update_cell(next_row, ic+1, prod_file_link)
            pfc = find_exact_header_index(headers, "Print Files")
            if pfc is not None:
                prod_ws.update_cell(next_row, pfc+1, print_file_link)
            self.loading_progress.config(value=90)
            self.update_idletasks()

            # Copy formulas down (Stage, Ship Date, Preview, Stitch Count)
            for key in ["Stage", "Ship Date", "Preview", "Stitch Count"]:
                idx = find_exact_header_index(headers, key)
                if idx is not None:
                    formula = prod_ws.cell(2, idx+1, value_render_option='FORMULA').value
                    if formula:
                        newf = re.sub(r'(\$?[A-Z]+\$?)2\b', lambda m: f"{m.group(1)}{next_row}", formula)
                        prod_ws.update_cell(next_row, idx+1, newf)

            # EMB Backing
            eb = find_exact_header_index(headers, "EMB Backing")
            if eb is not None:
                prod_ws.update_cell(next_row, eb+1, self.backing_type_var.get())

            self.loading_progress.config(value=100)
            self.update_idletasks()

            messagebox.showinfo("Success", "Order submitted successfully!")
            self.clear_prod_file()
            self.clear_print_files()
            self.clear_fields()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            # ------------------ always hide at end ------------------
            self.hide_loading()

    def _on_focus_change(self, event):
        """
        On *every* widget focus, check which widget *lost* focus last.
        If it was one of our comboboxes, run its pop-up logic.
        """
        prev = self._last_widget

        # COMPANY
        if prev == self.company_cb:
            self.company_selected(None)

        # MATERIALS
        elif prev in (
            self.material1_cb, self.material2_cb, self.material3_cb,
            self.material4_cb, self.material5_cb, self.back_material_cb
        ):
            var_map = {
                self.material1_cb: self.material1_var,
                self.material2_cb: self.material2_var,
                self.material3_cb: self.material3_var,
                self.material4_cb: self.material4_var,
                self.material5_cb: self.material5_var,
                self.back_material_cb: self.back_material_var
            }
            self.material_selected(var_map[prev])

        # FUR COLOR
        elif prev == self.fur_color_cb:
            self.fur_color_selected(None)

        # remember for next time
        self._last_widget = event.widget


    def clear_fields(self):
        self.company_var.set("")
        self.referral_var.set("")
        self.design_var.set("")
        self.quantity_var.set("")
        self.product_var.set("")
        self.due_date_var.set("")
        self.price_var.set("")
        self.date_type_var.set("")
        self.material1_var.set("")
        self.material2_var.set("")
        self.material3_var.set("")
        self.material4_var.set("")
        self.material5_var.set("")
        self.back_material_var.set("")
        self.fur_color_var.set("")
        self.backing_type_var.set("")
        self.notes_var.set("")
        self.prod_file_paths = []
        self.print_file_paths = []
        self.clear_prod_file()
        self.clear_print_files()
        self.company_cb.set_completion_list(get_companies())
        self.product_cb.set_completion_list(get_products())
        self.material1_cb.set_completion_list(get_union_material_list())
        self.material2_cb.set_completion_list(get_union_material_list())
        self.material3_cb.set_completion_list(get_union_material_list())
        self.material4_cb.set_completion_list(get_union_material_list())
        self.material5_cb.set_completion_list(get_union_material_list())
        self.back_material_cb.set_completion_list(get_union_material_list())
        self.fur_color_cb.set_completion_list(get_fur_colors())



# ===================== Inventory Tab =====================

class InventoryTab(tk.Frame):
    def o_r_key_handler(self, event):
        ch = event.char.lower()
        if ch == 'o':
            event.widget.set("Ordered")
        elif ch == 'r':
            event.widget.set("Received")
    
    def __init__(self, master):
        super().__init__(master)
        self.loading_frame = tk.Frame(self)
        self.loading_label = tk.Label(self.loading_frame, text="Submitting order, please wait...")
        self.loading_progress = ttk.Progressbar(self.loading_frame, mode='determinate', maximum=100, value=0)
        self.loading_label.pack(side="left", padx=5, pady=5)
        self.loading_progress.pack(side="left", padx=5, pady=5)

# ensure row 0 is the main form, row 1 is the progress bar
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=1)

        self.loading_frame.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.loading_frame.grid_remove()
        container = tk.Frame(self)
        container.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        # Thread Inventory Section
        thread_frame = tk.LabelFrame(container, text="Thread Inventory (15 rows)", padx=10, pady=10)
        thread_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        tk.Label(thread_frame, text="Thread Color").grid(row=0, column=0, padx=5, pady=5)
        tk.Label(thread_frame, text="O/R").grid(row=0, column=1, padx=5, pady=5)
        tk.Label(thread_frame, text="Quantity (# of Cones)").grid(row=0, column=2, padx=5, pady=5)
        self.thread_color_entries = []
        self.thread_or_entries = []
        self.thread_quantity_entries = []
        for i in range(15):
            thread_cb = FilterableCombobox(thread_frame, width=30)
            thread_cb.grid(row=i+1, column=0, padx=5, pady=2)
            thread_cb.set_completion_list(get_threads_inventory())
            thread_or_cb = ttk.Combobox(thread_frame, values=["Ordered", "Received"], state="readonly", width=10)
            thread_or_cb.grid(row=i+1, column=1, padx=5, pady=2)
            thread_or_cb.set("")
            thread_or_cb.bind("<Key>", self.o_r_key_handler)
            qty_entry = tk.Entry(thread_frame, width=10)
            qty_entry.grid(row=i+1, column=2, padx=5, pady=2)
            self.thread_color_entries.append(thread_cb)
            self.thread_or_entries.append(thread_or_cb)
            self.thread_quantity_entries.append(qty_entry)
        def update_thread_entries():
            self.show_loading()
            try:
                for i in range(15):
                    thread_color = self.thread_color_entries[i].get().strip()
                    thread_or = self.thread_or_entries[i].get().strip()
                    thread_qty = self.thread_quantity_entries[i].get().strip()
                    if thread_color and thread_qty and thread_or:
                        result = process_thread_entry(thread_color, thread_qty)
                        if not result:
                            messagebox.showinfo("Stopped", f"Stopped processing at row {i+1} due to cancellation or error.")
                            break
                        now = datetime.datetime.now()
                        date_stamp = f"{now.month}/{now.day}/{now.year} {now.hour:02d}:{now.minute:02d}:{now.second:02d}"
                        post_thread_data(thread_color, thread_qty, thread_or, date_stamp)
                        self.loading_progress.step(5)
                messagebox.showinfo("Success", "Thread inventory update complete.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")
            finally:
                self.clear_thread_entries()
                if ORDERED_TAB:
                    ORDERED_TAB.refresh()
                self.hide_loading()
        thread_submit = tk.Button(thread_frame, text="Submit Threads", command=update_thread_entries)
        thread_submit.grid(row=16, column=0, columnspan=3, pady=5)
        # Material Inventory Section
        material_frame = tk.LabelFrame(container, text="Material Inventory (15 rows)", padx=10, pady=10)
        material_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        tk.Label(material_frame, text="Material").grid(row=0, column=0, padx=5, pady=5)
        tk.Label(material_frame, text="O/R").grid(row=0, column=1, padx=5, pady=5)
        tk.Label(material_frame, text="Quantity").grid(row=0, column=2, padx=5, pady=5)
        self.material_entries = []
        self.material_or_entries = []
        self.material_quantity_entries = []
        for i in range(15):
            material_cb = FilterableCombobox(material_frame, width=30)
            material_cb.grid(row=i+1, column=0, padx=5, pady=2)
            material_cb.set_completion_list(get_union_material_list())
            material_or_cb = ttk.Combobox(material_frame, values=["Ordered", "Received"], state="readonly", width=10)
            material_or_cb.grid(row=i+1, column=1, padx=5, pady=2)
            material_or_cb.set("")
            material_or_cb.bind("<Key>", self.o_r_key_handler)
            qty_entry = tk.Entry(material_frame, width=10)
            qty_entry.grid(row=i+1, column=2, padx=5, pady=2)
            self.material_entries.append(material_cb)
            self.material_or_entries.append(material_or_cb)
            self.material_quantity_entries.append(qty_entry)
        material_submit = tk.Button(material_frame, text="Submit Materials", command=self.update_material_entries)
        material_submit.grid(row=16, column=0, columnspan=3, pady=5)

    def show_loading(self):
        self.loading_progress.config(value=0)
        self.loading_frame.grid()
        self.update_idletasks()

    def hide_loading(self):
        self.loading_frame.grid_remove()

    def update_material_entries(self):
        self.show_loading()
        try:
            for i in range(15):
                material_name = self.material_entries[i].get().strip()
                material_or = self.material_or_entries[i].get().strip()
                material_qty = self.material_quantity_entries[i].get().strip()
                if material_name and material_qty and material_or:
                    result = process_material_entry(material_name, material_qty, material_or)
                    if result is None:
                        messagebox.showinfo("Stopped", f"Stopped processing at row {i+1} due to cancellation or error.")
                        break
                    self.loading_progress.step(5)
            messagebox.showinfo("Success", "Material inventory update complete.")
            self.clear_material_entries()
            if ORDERED_TAB:
                ORDERED_TAB.refresh()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            self.hide_loading()

    def clear_thread_entries(self):
        for entry in self.thread_color_entries:
            entry.set("")
        for entry in self.thread_or_entries:
            entry.set("")
        for entry in self.thread_quantity_entries:
            entry.delete(0, tk.END)

    def clear_material_entries(self):
        for entry in self.material_entries:
            entry.set("")
        for entry in self.material_or_entries:
            entry.set("")
        for entry in self.material_quantity_entries:
            entry.delete(0, tk.END)

# ===================== Main Application =====================

def get_union_material_list():
    return sorted(list(set(get_materials()).union(set(get_fur_colors()))))

OrderEntryApp.get_union_material_list = lambda self: get_union_material_list()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        root.title("JR and Co. - Order and Inventory")
        root.geometry("900x650")
        style = ttk.Style()
        style.theme_use('default')
        style.configure("TNotebook.Tab", borderwidth=2, padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", "gray40")], foreground=[("selected", "white")])
        main_container = tk.Frame(root)
        main_container.pack(expand=True)
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        notebook = ttk.Notebook(main_container)
        notebook.grid(row=0, column=0, sticky="nsew")
        order_frame = tk.Frame(notebook)
        order_frame.grid_rowconfigure(0, weight=1)
        order_frame.grid_columnconfigure(0, weight=1)
        order_app = OrderEntryApp(order_frame)
        order_app.pack(expand=True, fill=tk.BOTH)
        notebook.add(order_frame, text="Order Entry")
        inventory_frame = tk.Frame(notebook)
        inventory_frame.grid_rowconfigure(0, weight=1)
        inventory_frame.grid_columnconfigure(0, weight=1)
        inv_app = InventoryTab(inventory_frame)
        inv_app.pack(expand=True, fill=tk.BOTH)
        notebook.add(inventory_frame, text="Inventory")
        ordered_frame = tk.Frame(notebook)
        ordered_frame.grid_rowconfigure(0, weight=1)
        ordered_frame.grid_columnconfigure(0, weight=1)
        ordered_tab = InventoryOrderedTab(ordered_frame)
        ordered_tab.pack(expand=True, fill=tk.BOTH)
        notebook.add(ordered_frame, text="Inventory Ordered")
        ORDERED_TAB = ordered_tab
        root.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("An error occurred. Press Enter to exit...")
