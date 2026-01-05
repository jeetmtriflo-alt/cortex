import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
import os
from pathlib import Path


class FileComparisonTool:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV vs XLSX Hostname Comparison Tool")
        self.root.geometry("800x600")
        
        # File paths
        self.csv_file_path = None
        self.xlsx_file_path = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Title
        title_label = tk.Label(
            self.root, 
            text="CSV vs XLSX Hostname Comparison Tool",
            font=("Arial", 16, "bold"),
            pady=10
        )
        title_label.pack()
        
        # Instructions
        instructions = tk.Label(
            self.root,
            text="Drag and drop your CSV file and XLSX file below, or click 'Browse' to select files",
            font=("Arial", 10),
            pady=5
        )
        instructions.pack()
        
        # CSV File Section
        csv_frame = tk.LabelFrame(self.root, text="CSV File (Source)", padx=10, pady=10)
        csv_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.csv_label = tk.Label(
            csv_frame,
            text="No CSV file selected",
            bg="lightgray",
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=5,
            pady=5
        )
        self.csv_label.pack(fill=tk.X, side=tk.LEFT, expand=True)
        self.csv_label.drop_target_register(DND_FILES)
        self.csv_label.dnd_bind('<<Drop>>', self.on_csv_drop)
        
        csv_browse_btn = tk.Button(csv_frame, text="Browse", command=self.browse_csv)
        csv_browse_btn.pack(side=tk.RIGHT, padx=5)
        
        # XLSX File Section
        xlsx_frame = tk.LabelFrame(self.root, text="XLSX File (Reference)", padx=10, pady=10)
        xlsx_frame.pack(fill=tk.X, padx=20, pady=10)
        
        self.xlsx_label = tk.Label(
            xlsx_frame,
            text="No XLSX file selected",
            bg="lightgray",
            relief=tk.SUNKEN,
            anchor=tk.W,
            padx=5,
            pady=5
        )
        self.xlsx_label.pack(fill=tk.X, side=tk.LEFT, expand=True)
        self.xlsx_label.drop_target_register(DND_FILES)
        self.xlsx_label.dnd_bind('<<Drop>>', self.on_xlsx_drop)
        
        xlsx_browse_btn = tk.Button(xlsx_frame, text="Browse", command=self.browse_xlsx)
        xlsx_browse_btn.pack(side=tk.RIGHT, padx=5)
        
        # Column Selection (optional)
        column_frame = tk.Frame(self.root)
        column_frame.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(column_frame, text="CSV Column (default: Column A/Index 0):").pack(side=tk.LEFT, padx=5)
        self.csv_column_var = tk.StringVar(value="0")
        csv_column_entry = tk.Entry(column_frame, textvariable=self.csv_column_var, width=10)
        csv_column_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(column_frame, text="XLSX Column (default: Column A/Index 0):").pack(side=tk.LEFT, padx=5)
        self.xlsx_column_var = tk.StringVar(value="0")
        xlsx_column_entry = tk.Entry(column_frame, textvariable=self.xlsx_column_var, width=10)
        xlsx_column_entry.pack(side=tk.LEFT, padx=5)
        
        # Compare Button
        compare_btn = tk.Button(
            self.root,
            text="Compare Files",
            command=self.compare_files,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10
        )
        compare_btn.pack(pady=20)
        
        # Results Section
        results_frame = tk.LabelFrame(self.root, text="Results", padx=10, pady=10)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Scrollbar for results
        scrollbar = tk.Scrollbar(results_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.results_text = tk.Text(
            results_frame,
            yscrollcommand=scrollbar.set,
            wrap=tk.WORD,
            font=("Consolas", 10)
        )
        self.results_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.results_text.yview)
        
        # Export Button
        export_btn = tk.Button(
            self.root,
            text="Export Results to CSV",
            command=self.export_results,
            bg="#2196F3",
            fg="white",
            font=("Arial", 10),
            padx=15,
            pady=5
        )
        export_btn.pack(pady=10)
        
        # Store unique hostnames for export
        self.unique_hostnames = []
        
    def on_csv_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0].strip('{}').strip('"').strip("'")
            if os.path.exists(file_path) and file_path.lower().endswith('.csv'):
                self.csv_file_path = file_path
                self.csv_label.config(text=f"CSV: {os.path.basename(file_path)}", bg="lightgreen")
            else:
                messagebox.showerror("Error", "Please drop a valid CSV file")
    
    def on_xlsx_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0].strip('{}').strip('"').strip("'")
            if os.path.exists(file_path) and file_path.lower().endswith(('.xlsx', '.xls')):
                self.xlsx_file_path = file_path
                self.xlsx_label.config(text=f"XLSX: {os.path.basename(file_path)}", bg="lightgreen")
            else:
                messagebox.showerror("Error", "Please drop a valid XLSX file")
    
    def browse_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.csv_file_path = file_path
            self.csv_label.config(text=f"CSV: {os.path.basename(file_path)}", bg="lightgreen")
    
    def browse_xlsx(self):
        file_path = filedialog.askopenfilename(
            title="Select XLSX File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.xlsx_file_path = file_path
            self.xlsx_label.config(text=f"XLSX: {os.path.basename(file_path)}", bg="lightgreen")
    
    def get_column_index(self, column_str):
        """Convert column string to index (supports both numeric and Excel column letters)"""
        try:
            # Try numeric first
            return int(column_str)
        except ValueError:
            # Try Excel column letter (A=0, B=1, etc.)
            column_str = column_str.upper()
            result = 0
            for char in column_str:
                result = result * 26 + (ord(char) - ord('A') + 1)
            return result - 1
    
    def compare_files(self):
        if not self.csv_file_path or not self.xlsx_file_path:
            messagebox.showerror("Error", "Please select both CSV and XLSX files")
            return
        
        try:
            # Get column indices
            csv_col = self.get_column_index(self.csv_column_var.get())
            xlsx_col = self.get_column_index(self.xlsx_column_var.get())
            
            # Read CSV file
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, "Reading CSV file...\n")
            self.root.update()
            
            csv_df = pd.read_csv(self.csv_file_path)
            csv_hostnames = csv_df.iloc[:, csv_col].astype(str).str.strip().dropna().unique()
            
            self.results_text.insert(tk.END, f"Found {len(csv_hostnames)} unique hostnames in CSV\n")
            self.root.update()
            
            # Read XLSX file
            self.results_text.insert(tk.END, "Reading XLSX file...\n")
            self.root.update()
            
            xlsx_df = pd.read_excel(self.xlsx_file_path)
            xlsx_hostnames = set(xlsx_df.iloc[:, xlsx_col].astype(str).str.strip().dropna().unique())
            
            self.results_text.insert(tk.END, f"Found {len(xlsx_hostnames)} unique hostnames in XLSX\n\n")
            self.root.update()
            
            # Find unique hostnames in CSV that are NOT in XLSX
            self.results_text.insert(tk.END, "Comparing hostnames...\n")
            self.root.update()
            
            unique_in_csv = [hostname for hostname in csv_hostnames if hostname not in xlsx_hostnames]
            self.unique_hostnames = unique_in_csv
            
            # Display results
            self.results_text.insert(tk.END, "=" * 60 + "\n")
            self.results_text.insert(tk.END, f"UNIQUE HOSTNAMES IN CSV (NOT IN XLSX):\n")
            self.results_text.insert(tk.END, "=" * 60 + "\n\n")
            
            if unique_in_csv:
                self.results_text.insert(tk.END, f"Total unique hostnames: {len(unique_in_csv)}\n\n")
                for i, hostname in enumerate(unique_in_csv, 1):
                    self.results_text.insert(tk.END, f"{i}. {hostname}\n")
            else:
                self.results_text.insert(tk.END, "No unique hostnames found. All CSV hostnames exist in XLSX file.\n")
            
            self.results_text.insert(tk.END, "\n" + "=" * 60 + "\n")
            self.results_text.insert(tk.END, f"Comparison completed successfully!\n")
            
            messagebox.showinfo("Success", f"Comparison completed!\nFound {len(unique_in_csv)} unique hostnames.")
            
        except Exception as e:
            error_msg = f"Error during comparison: {str(e)}"
            self.results_text.insert(tk.END, f"\nERROR: {error_msg}\n")
            messagebox.showerror("Error", error_msg)
    
    def export_results(self):
        if not self.unique_hostnames:
            messagebox.showwarning("Warning", "No results to export. Please run comparison first.")
            return
        
        try:
            file_path = filedialog.asksaveasfilename(
                title="Save Results",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if file_path:
                df = pd.DataFrame(self.unique_hostnames, columns=["Unique_Hostnames"])
                df.to_csv(file_path, index=False)
                messagebox.showinfo("Success", f"Results exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export results: {str(e)}")


def main():
    root = TkinterDnD.Tk()
    app = FileComparisonTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()

