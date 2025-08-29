import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from datetime import datetime
import pandas as pd
import json

# Add paths
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
src_dir = os.path.join(project_root, 'src')
sys.path.insert(0, project_root)
sys.path.insert(0, src_dir)

try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False

try:
    from src.coda_extractor import CodaTimesheetExtractor
    from src.data_processor import TimesheetProcessor
    from config.config import Config
except ImportError as e:
    print(f"Error importing modules: {e}")
    sys.exit(1)

class TimesheetExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Coda Timesheet Extractor - Enhanced")
        self.root.geometry("1000x700")
        
        # Variables
        self.api_token = tk.StringVar()
        self.doc_id = tk.StringVar()
        self.table_id = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        self.max_rows_var = tk.StringVar(value="")
        
        # Data
        self.current_df = None
        self.current_metrics = None
        
        os.makedirs('logs', exist_ok=True)
        os.makedirs('data/raw', exist_ok=True)
        os.makedirs('data/processed', exist_ok=True)
        
        self.load_config()
        self.create_widgets()
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Title
        ttk.Label(main_frame, text="Coda Timesheet Extractor", font=('Arial', 16, 'bold')).grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Config
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(config_frame, text="API Token:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.api_token, show="*", width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        ttk.Label(config_frame, text="Document ID:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.doc_id, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        ttk.Label(config_frame, text="Table ID:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.table_id, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        ttk.Label(config_frame, text="Max Rows:").grid(row=3, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.max_rows_var, width=20).grid(row=3, column=1, sticky=tk.W, pady=2, padx=(10, 0))
        
        config_btn_frame = ttk.Frame(config_frame)
        config_btn_frame.grid(row=4, column=0, columnspan=2, pady=10)
        ttk.Button(config_btn_frame, text="Load from .env", command=self.load_from_env).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_btn_frame, text="Save Config", command=self.save_config).pack(side=tk.LEFT, padx=5)
        
        # Actions
        action_frame = ttk.LabelFrame(main_frame, text="Actions", padding="10")
        action_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(action_frame, text="List Documents", command=self.list_documents).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="List Tables", command=self.list_tables).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Extract Data", command=self.extract_data).pack(side=tk.LEFT, padx=10)
        
        # Progress
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status_text).grid(row=0, column=0, sticky=tk.W)
        
        # Results
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="5")
        results_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Data tab
        self.data_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_frame, text="Data")
        self.create_data_tab()
        
        # Metrics tab
        self.metrics_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.metrics_frame, text="Metrics")
        self.create_metrics_tab()
        
        # Log tab
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="Log")
        self.create_log_tab()
        
        # Export
        export_frame = ttk.Frame(main_frame)
        export_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        self.export_csv_btn = ttk.Button(export_frame, text="Export CSV", command=self.export_csv, state='disabled')
        self.export_csv_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_excel_btn = ttk.Button(export_frame, text="Export Excel", command=self.export_excel, state='disabled')
        self.export_excel_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(export_frame, text="Show Metrics", command=self.show_metrics, state='disabled').pack(side=tk.LEFT, padx=5)
    
    def create_data_tab(self):
        self.data_frame.columnconfigure(0, weight=1)
        self.data_frame.rowconfigure(0, weight=1)
        
        if HAS_TKSHEET:
            self.sheet = Sheet(self.data_frame, height=400, width=900)
            self.sheet.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
            self.sheet.enable_bindings("single_select", "column_width_resize", "copy", "select_all")
        else:
            tree_frame = ttk.Frame(self.data_frame)
            tree_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
            tree_frame.columnconfigure(0, weight=1)
            tree_frame.rowconfigure(0, weight=1)
            
            self.tree = ttk.Treeview(tree_frame)
            self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
            v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        # Summary
        self.summary_label = ttk.Label(self.data_frame, text="No data loaded")
        self.summary_label.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    
    def create_metrics_tab(self):
        self.metrics_frame.columnconfigure(0, weight=1)
        self.metrics_frame.rowconfigure(0, weight=1)
        
        self.metrics_text = tk.Text(self.metrics_frame, height=20, wrap=tk.WORD)
        self.metrics_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(self.metrics_frame, orient="vertical", command=self.metrics_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.metrics_text.configure(yscrollcommand=scrollbar.set)
    
    def create_log_tab(self):
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(self.log_frame, height=15, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        log_scrollbar = ttk.Scrollbar(self.log_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        ttk.Button(self.log_frame, text="Clear Log", command=lambda: self.log_text.delete(1.0, tk.END)).grid(row=1, column=0, pady=5)
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message):
        self.status_text.set(message)
        self.root.update_idletasks()
    
    def load_config(self):
        if os.path.exists("gui_config.json"):
            try:
                with open("gui_config.json", 'r') as f:
                    config = json.load(f)
                    self.doc_id.set(config.get('doc_id', ''))
                    self.table_id.set(config.get('table_id', ''))
                    self.max_rows_var.set(config.get('max_rows', ''))
            except:
                pass
    
    def save_config(self):
        config = {'doc_id': self.doc_id.get(), 'table_id': self.table_id.get(), 'max_rows': self.max_rows_var.get()}
        try:
            with open("gui_config.json", 'w') as f:
                json.dump(config, f, indent=2)
            messagebox.showinfo("Success", "Configuration saved!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save: {e}")
    
    def load_from_env(self):
        if not os.path.exists(".env"):
            messagebox.showerror("Error", ".env file not found!")
            return
        
        try:
            with open(".env", 'r') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith('CODA_API_TOKEN='):
                        self.api_token.set(line.split('=', 1)[1])
                    elif line.startswith('CODA_DOC_ID='):
                        self.doc_id.set(line.split('=', 1)[1])
                    elif line.startswith('CODA_TABLE_ID='):
                        self.table_id.set(line.split('=', 1)[1])
            messagebox.showinfo("Success", "Configuration loaded!")
            self.log_message("Configuration loaded from .env")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load .env: {e}")
    
    def validate_inputs(self):
        if not self.api_token.get():
            messagebox.showerror("Error", "API Token required!")
            return False
        if not self.doc_id.get():
            messagebox.showerror("Error", "Document ID required!")
            return False
        if not self.table_id.get():
            messagebox.showerror("Error", "Table ID required!")
            return False
        return True
    
    def _create_extractor(self):
        # Set env vars temporarily
        os.environ['CODA_API_TOKEN'] = self.api_token.get()
        os.environ['CODA_DOC_ID'] = self.doc_id.get() or 'dummy'
        os.environ['CODA_TABLE_ID'] = self.table_id.get() or 'dummy'
        
        extractor = CodaTimesheetExtractor()
        extractor.api_token = self.api_token.get()
        extractor.headers = {"Authorization": f"Bearer {self.api_token.get()}", "Content-Type": "application/json"}
        return extractor
    
    def list_documents(self):
        if not self.api_token.get():
            messagebox.showerror("Error", "API Token required!")
            return
        threading.Thread(target=self._list_documents_thread).start()
    
    def _list_documents_thread(self):
        self.progress.start()
        self.update_status("Fetching documents...")
        try:
            extractor = self._create_extractor()
            docs = extractor.get_documents()
            self.show_selection_window("Documents", docs.get('items', []), lambda x: self.doc_id.set(x))
            self.log_message(f"Found {len(docs.get('items', []))} documents")
            self.update_status("Documents fetched")
        except Exception as e:
            self.log_message(f"Error: {e}")
            messagebox.showerror("Error", f"Could not fetch documents: {e}")
        finally:
            self.progress.stop()
    
    def list_tables(self):
        if not self.api_token.get() or not self.doc_id.get():
            messagebox.showerror("Error", "API Token and Document ID required!")
            return
        threading.Thread(target=self._list_tables_thread).start()
    
    def _list_tables_thread(self):
        self.progress.start()
        self.update_status("Fetching tables...")
        try:
            extractor = self._create_extractor()
            tables = extractor.get_tables(self.doc_id.get())
            self.show_selection_window("Tables", tables.get('items', []), lambda x: self.table_id.set(x))
            self.log_message(f"Found {len(tables.get('items', []))} tables")
            self.update_status("Tables fetched")
        except Exception as e:
            self.log_message(f"Error: {e}")
            messagebox.showerror("Error", f"Could not fetch tables: {e}")
        finally:
            self.progress.stop()
    
    def show_selection_window(self, title, items, callback):
        window = tk.Toplevel(self.root)
        window.title(f"Available {title}")
        window.geometry("600x400")
        
        frame = ttk.Frame(window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(frame, columns=("name", "id"), show="headings")
        tree.heading("name", text="Name")
        tree.heading("id", text="ID")
        tree.column("name", width=300)
        tree.column("id", width=200)
        
        for item in items:
            tree.insert("", tk.END, values=(item['name'], item['id']))
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        def on_select(event):
            if tree.selection():
                selection = tree.selection()[0]
                item_id = tree.item(selection)['values'][1]
                callback(item_id)
                window.destroy()
        
        tree.bind("<Double-1>", on_select)
        ttk.Label(frame, text="Double-click to select").pack(pady=5)
    
    def extract_data(self):
        if not self.validate_inputs():
            return
        threading.Thread(target=self._extract_data_thread).start()
    
    def _extract_data_thread(self):
        self.progress.start()
        self.update_status("Extracting data...")
        
        try:
            extractor = self._create_extractor()
            self.log_message("Starting extraction...")
            
            max_rows = None
            if self.max_rows_var.get().strip():
                try:
                    max_rows = int(self.max_rows_var.get().strip())
                except:
                    self.log_message("Invalid max rows, extracting all")
            
            raw_data = extractor.get_timesheet_data(self.doc_id.get(), self.table_id.get(), max_rows=max_rows)
            self.log_message(f"Extracted {len(raw_data.get('items', []))} rows")
            
            processor = TimesheetProcessor()
            df = processor.process_raw_data(raw_data)
            df_cleaned = processor.clean_timesheet_data(df)
            metrics = processor.calculate_timesheet_metrics(df_cleaned)
            
            self.current_df = df_cleaned
            self.current_metrics = metrics
            
            self.display_data(df_cleaned)
            self.display_metrics(metrics)
            
            summary = processor.generate_summary(df_cleaned)
            summary_text = f"Rows: {summary['total_rows']}, Columns: {len(summary['columns'])}"
            if summary.get('total_hours'):
                summary_text += f", Total Hours: {summary['total_hours']:.1f}"
            if summary.get('date_range'):
                summary_text += f", Date Range: {summary['date_range']}"
            
            self.summary_label.config(text=summary_text)
            self.export_csv_btn.config(state='normal')
            self.export_excel_btn.config(state='normal')
            
            self.log_message("Extraction completed!")
            self.update_status(f"Extracted {len(df_cleaned)} rows")
            messagebox.showinfo("Success", f"Extracted {len(df_cleaned)} rows with proper column names!")
            
        except Exception as e:
            self.log_message(f"Error: {e}")
            messagebox.showerror("Error", f"Extraction failed: {e}")
        finally:
            self.progress.stop()
    
    def display_data(self, df):
        if HAS_TKSHEET:
            if df.empty:
                return
            self.sheet.set_sheet_data([])
            self.sheet.headers(list(df.columns))
            data = df.values.tolist()
            for i, row in enumerate(data):
                for j, cell in enumerate(row):
                    data[i][j] = "" if pd.isna(cell) else str(cell)
            self.sheet.set_sheet_data(data)
            self.sheet.set_all_column_widths()
        else:
            for item in self.tree.get_children():
                self.tree.delete(item)
            if df.empty:
                return
            columns = list(df.columns)
            self.tree["columns"] = columns
            self.tree["show"] = "headings"
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100, minwidth=50)
            for _, row in df.iterrows():
                values = [str(row[col]) if pd.notna(row[col]) else "" for col in columns]
                self.tree.insert("", tk.END, values=values)
    
    def display_metrics(self, metrics):
        self.metrics_text.delete(1.0, tk.END)
        text = "TIMESHEET METRICS\n" + "="*30 + "\n\n"
        
        if 'total_hours' in metrics:
            text += f"Total Hours: {metrics['total_hours']:.2f}\n"
        if 'average_daily_hours' in metrics:
            text += f"Average Daily: {metrics['average_daily_hours']:.2f}\n"
        if 'max_daily_hours' in metrics:
            text += f"Max Daily: {metrics['max_daily_hours']:.2f}\n"
        if 'overtime_days' in metrics:
            text += f"Overtime Days: {metrics['overtime_days']}\n"
        
        if 'weekly_totals' in metrics and metrics['weekly_totals']:
            text += "\nWEEKLY BREAKDOWN:\n"
            for week, hours in metrics['weekly_totals'].items():
                text += f"Week {week}: {hours:.2f} hours\n"
        
        if 'project_breakdown' in metrics and metrics['project_breakdown']:
            text += "\nPROJECT BREAKDOWN:\n"
            for project, hours in metrics['project_breakdown'].items():
                text += f"{project}: {hours:.2f} hours\n"
        
        self.metrics_text.insert(tk.END, text)
    
    def show_metrics(self):
        if self.current_metrics:
            self.notebook.select(self.metrics_frame)
        else:
            messagebox.showinfo("Info", "No metrics available")
    
    def export_csv(self):
        if self.current_df is None:
            messagebox.showerror("Error", "No data to export!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialname=f"timesheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if filename:
            try:
                self.current_df.to_csv(filename, index=False)
                self.log_message(f"Exported to: {filename}")
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {e}")
    
    def export_excel(self):
        if self.current_df is None:
            messagebox.showerror("Error", "No data to export!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialname=f"timesheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                self.current_df.to_excel(filename, index=False)
                self.log_message(f"Exported to: {filename}")
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {e}")

def main():
    root = tk.Tk()
    app = TimesheetExtractorGUI(root)
    
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()