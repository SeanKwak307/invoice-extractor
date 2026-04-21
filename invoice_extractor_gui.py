#!/usr/bin/env python3
"""
Invoice PDF Extractor - GUI Application
Works on Windows and macOS
"""

import os
import re
import sys
import threading
import pandas as pd
import pdfplumber
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Try to import tkinter (should be built-in with Python)
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk, scrolledtext
    GUI_AVAILABLE = True
except ImportError as e:
    GUI_AVAILABLE = False
    print(f"Tkinter not available: {e}")

def extract_invoice_data_corrected(pdf_path):
    """Extract data with proper charges/allowances linking"""
    all_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                item_pattern = r'^(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\w+)\s+(.+?)\s+([\d.]+)\s+([\d.-]+)$'
                match = re.match(item_pattern, line)
                
                if match:
                    item_num = match.group(1)
                    ordered = match.group(2)
                    shipped = match.group(3)
                    pack = match.group(4)
                    unit = match.group(5)
                    description = match.group(6).strip()
                    price = match.group(7)
                    amount = float(match.group(8))
                    
                    charges = 0
                    if i + 1 < len(lines) and '*** Charges/Allowances ***' in lines[i + 1]:
                        charges_match = re.search(r'([\d.]+)-?', lines[i + 1])
                        if charges_match:
                            charges_value = charges_match.group(1)
                            charges = float(charges_value)
                            if '-' in lines[i + 1] and not charges_match.group(0).startswith('-'):
                                charges = -charges
                        i += 1
                    
                    final_amount = amount + charges
                    
                    all_data.append({
                        'Item_Number': item_num,
                        'Description': description,
                        'Ordered': int(ordered),
                        'Shipped': int(shipped),
                        'Pack': int(pack),
                        'Unit': unit,
                        'Unit_Price': float(price),
                        'Subtotal': amount,
                        'Adjustment': charges,
                        'Final_Amount': final_amount,
                        'Source_PDF': os.path.basename(pdf_path),
                        'Page_Number': page_num
                    })
                i += 1
    return all_data

def process_pdfs(pdf_folder_path, progress_callback=None):
    """Process all PDFs and return DataFrame"""
    all_data = []
    pdf_files = list(Path(pdf_folder_path).glob('*.pdf'))
    total = len(pdf_files)
    
    for idx, pdf_file in enumerate(pdf_files):
        if progress_callback:
            progress_callback(idx + 1, total, pdf_file.name)
        try:
            data = extract_invoice_data_corrected(pdf_file)
            if data:
                all_data.extend(data)
        except Exception as e:
            print(f"Error processing {pdf_file.name}: {e}")
    
    if all_data:
        df = pd.DataFrame(all_data)
        df['Invoice_ID'] = df['Source_PDF'].str.replace('.pdf', '', case=False)
        return df
    return pd.DataFrame()

class InvoiceExtractorGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Invoice PDF Extractor")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        self.input_folder = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.output_file = tk.StringVar(value="consolidated_invoices.xlsx")
        
        self.setup_ui()
    
    def setup_ui(self):
        title = tk.Label(self.root, text="Invoice PDF Extractor", font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        folder_frame = tk.LabelFrame(self.root, text="Step 1: Select PDF Folder", padx=10, pady=5)
        folder_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Entry(folder_frame, textvariable=self.input_folder, width=60).pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(folder_frame, text="Browse", command=self.browse_folder).pack(side="right", padx=5)
        
        output_frame = tk.LabelFrame(self.root, text="Step 2: Output Excel File", padx=10, pady=5)
        output_frame.pack(pady=10, padx=20, fill="x")
        
        tk.Entry(output_frame, textvariable=self.output_file, width=60).pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(output_frame, text="Browse", command=self.browse_output).pack(side="right", padx=5)
        
        progress_frame = tk.LabelFrame(self.root, text="Progress", padx=10, pady=5)
        progress_frame.pack(pady=10, padx=20, fill="x")
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=5, pady=5)
        
        self.status_label = tk.Label(progress_frame, text="Ready", fg="blue")
        self.status_label.pack(pady=5)
        
        log_frame = tk.LabelFrame(self.root, text="Processing Log", padx=10, pady=5)
        log_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        
        self.start_button = tk.Button(self.root, text="Start Extraction", command=self.start_extraction,
                                      bg="green", fg="white", font=("Arial", 12, "bold"), padx=20, pady=5)
        self.start_button.pack(pady=10)
        
        self.bottom_status = tk.Label(self.root, text="Select a folder containing PDF invoices", fg="gray")
        self.bottom_status.pack(side="bottom", pady=5)
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with Invoice PDFs")
        if folder:
            self.input_folder.set(folder)
    
    def browse_output(self):
        file = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.output_file.set(file)
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_progress(self, current, total, filename):
        percent = (current / total) * 100
        self.progress_var.set(percent)
        self.status_label.config(text=f"Processing: {filename} ({current}/{total})")
        self.log(f"  {current}/{total}: {filename}")
        self.root.update()
    
    def run_extraction(self):
        try:
            input_folder = self.input_folder.get().strip()
            output_file = self.output_file.get().strip()
            
            if not input_folder:
                self.log("Error: No input folder selected")
                return
            if not os.path.exists(input_folder):
                self.log(f"Error: Folder '{input_folder}' does not exist")
                return
            if not output_file:
                output_file = "consolidated_invoices.xlsx"
            if not output_file.endswith('.xlsx'):
                output_file += '.xlsx'
            
            self.log(f"Reading PDFs from: {input_folder}")
            self.log(f"Output will be saved to: {output_file}")
            
            df = process_pdfs(input_folder, self.update_progress)
            
            if df.empty:
                self.log("No data extracted. Please check your PDF files.")
                return
            
            self.log(f"Total items extracted: {len(df)}")
            
            column_order = ['Invoice_ID', 'Item_Number', 'Description', 'Ordered', 'Shipped',
                            'Pack', 'Unit', 'Unit_Price', 'Subtotal', 'Adjustment', 'Final_Amount']
            df = df[[c for c in column_order if c in df.columns]]
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Invoice_Items', index=False)
                
                summary = df.groupby('Invoice_ID').agg({
                    'Item_Number': 'count',
                    'Shipped': 'sum',
                    'Final_Amount': 'sum'
                }).round(2)
                summary.columns = ['Total_Items', 'Total_Shipped', 'Total_Amount']
                summary.reset_index().to_excel(writer, sheet_name='Summary', index=False)
                
                product_df = df[~df['Description'].str.contains('CHARGES/ALLOWANCES', na=False)]
                if not product_df.empty:
                    pivot = product_df.groupby('Description').agg({
                        'Shipped': 'sum',
                        'Final_Amount': 'sum'
                    }).sort_values('Final_Amount', ascending=False).head(20).round(2)
                    pivot.reset_index().to_excel(writer, sheet_name='Top_Products', index=False)
            
            self.log(f"\nSUCCESS! Excel file saved to: {output_file}")
            self.log(f"Summary: {df['Invoice_ID'].nunique()} invoices, {len(df)} items, ${df['Final_Amount'].sum():,.2f} total")
            
            messagebox.showinfo("Success", f"Extraction complete!\n\n{df['Invoice_ID'].nunique()} invoices processed\n{len(df)} line items extracted\n\nSaved to:\n{output_file}")
            
        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        finally:
            self.start_button.config(state="normal")
            self.progress_var.set(0)
            self.status_label.config(text="Done")
    
    def start_extraction(self):
        self.start_button.config(state="disabled")
        self.log_text.delete(1.0, tk.END)
        self.log("Starting extraction...")
        thread = threading.Thread(target=self.run_extraction)
        thread.daemon = True
        thread.start()
    
    def run(self):
        self.root.mainloop()

def main():
    if not GUI_AVAILABLE:
        print("Tkinter not available. Please install python3-tk.")
        print("On Ubuntu/Debian: sudo apt-get install python3-tk")
        print("On macOS: tkinter comes with Python, but you may need to install python-tk via brew")
        sys.exit(1)
    
    app = InvoiceExtractorGUI()
    app.run()

if __name__ == "__main__":
    main()
