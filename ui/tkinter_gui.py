import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from datetime import datetime
import pandas as pd
import json

# Add the project root and src to Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
src_dir = os.path.join(project_root, 'src')
sys.path.insert(0, project_root)
sys.path.insert(0, src_dir)

# Try to import tksheet (optional advanced grid component)
try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False
    print("tksheet not available - using basic treeview instead")
    print("To get advanced grid features, install: pip install tksheet")

try:
    from src.coda_extractor import CodaTimesheetExtractor
    from src.data_processor import TimesheetProcessor
    from config.config import Config
except ImportError as e:
    print(f"Error: Could not import required modules: {e}")
    print("Make sure you're running from the project root directory")
    print("Current working directory:", os.getcwd())
    print("Script location:", __file__)
    sys.exit(1)

class TimesheetExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Coda Timesheet Extractor")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Variables
        self.api_token = tk.StringVar()
        self.doc_id = tk.StringVar()
        self.table_id = tk.StringVar()
        self.status_text = tk.StringVar(value="Ready")
        
        # Current data
        self.current_df = None
        self.filtered_df = None
        
        # Create necessary directories
        os.makedirs('logs', exist_ok=True)
        os.makedirs('data/raw', exist_ok=True)
        os.makedirs('data/processed', exist_ok=True)
        
        # Load saved config if exists
        self.load_config()
        
        # Create UI
        self.create_widgets()
        
        # Center window
        self.center_window()
    
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)  # Results area expands
        
        # Title
        title_label = ttk.Label(main_frame, text="Coda Timesheet Extractor", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Configuration section
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        config_frame.columnconfigure(1, weight=1)
        
        # API Token
        ttk.Label(config_frame, text="API Token:").grid(row=0, column=0, sticky=tk.W, pady=2)
        token_entry = ttk.Entry(config_frame, textvariable=self.api_token, show="*", width=50)
        token_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # Document ID
        ttk.Label(config_frame, text="Document ID:").grid(row=1, column=0, sticky=tk.W, pady=2)
        doc_entry = ttk.Entry(config_frame, textvariable=self.doc_id, width=50)
        doc_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # Table ID
        ttk.Label(config_frame, text="Table ID:").grid(row=2, column=0, sticky=tk.W, pady=2)
        table_entry = ttk.Entry(config_frame, textvariable=self.table_id, width=50)
        table_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2, padx=(10, 0))
        
        # Config buttons
        config_btn_frame = ttk.Frame(config_frame)
        config_btn_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(config_btn_frame, text="Load from .env", 
                  command=self.load_from_env).pack(side=tk.LEFT, padx=5)
        ttk.Button(config_btn_frame, text="Save Config", 
                  command=self.save_config).pack(side=tk.LEFT, padx=5)
        
        # Action buttons
        action_frame = ttk.LabelFrame(main_frame, text="Actions", padding="10")
        action_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(action_frame, text="List Documents", 
                  command=self.list_documents).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="List Tables", 
                  command=self.list_tables).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Extract Data", 
                  command=self.extract_data).pack(side=tk.LEFT, padx=10)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status bar
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        status_frame.columnconfigure(0, weight=1)
        
        status_label = ttk.Label(status_frame, textvariable=self.status_text)
        status_label.grid(row=0, column=0, sticky=tk.W)
        
        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="5")
        results_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Data tab
        self.data_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_frame, text="Data")
        self.create_data_tab()
        
        # Log tab
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="Log")
        self.create_log_tab()
        
        # Export buttons
        export_frame = ttk.Frame(main_frame)
        export_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        self.export_csv_btn = ttk.Button(export_frame, text="Export CSV", 
                                        command=self.export_csv, state='disabled')
        self.export_csv_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_excel_btn = ttk.Button(export_frame, text="Export Excel", 
                                          command=self.export_excel, state='disabled')
        self.export_excel_btn.pack(side=tk.LEFT, padx=5)
    
    def create_data_tab(self):
        """Create the data viewing tab"""
        self.data_frame.columnconfigure(0, weight=1)
        self.data_frame.rowconfigure(0, weight=1)
        
        # Create grid frame
        grid_frame = ttk.Frame(self.data_frame)
        grid_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        grid_frame.columnconfigure(0, weight=1)
        grid_frame.rowconfigure(0, weight=1)
        
        if HAS_TKSHEET:
            # Use tksheet for Excel-like grid
            self.sheet = Sheet(grid_frame,
                              page_up_down_select_row=True,
                              column_width=120,
                              startup_select=(0, 1, "rows"),
                              headers=[],
                              height=400,
                              width=700)
            self.sheet.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Enable grid features
            self.sheet.enable_bindings("single_select",
                                      "row_select", 
                                      "column_width_resize",
                                      "arrowkeys",
                                      "right_click_popup_menu",
                                      "rc_select",
                                      "copy",
                                      "select_all")
        else:
            # Fallback to treeview if tksheet not available
            self.tree = ttk.Treeview(grid_frame)
            self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Scrollbars for treeview
            v_scrollbar = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
            v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            self.tree.configure(yscrollcommand=v_scrollbar.set)
            
            h_scrollbar = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree.xview)
            h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
            self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Summary frame
        summary_frame = ttk.Frame(self.data_frame)
        summary_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        self.summary_label = ttk.Label(summary_frame, text="No data loaded")
        self.summary_label.grid(row=0, column=0, sticky=tk.W)
    
    def create_log_tab(self):
        """Create the log tab"""
        self.log_frame.columnconfigure(0, weight=1)
        self.log_frame.rowconfigure(0, weight=1)
        
        # Text widget with scrollbar
        log_text_frame = ttk.Frame(self.log_frame)
        log_text_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        log_text_frame.columnconfigure(0, weight=1)
        log_text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_text_frame, height=15, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        log_scrollbar = ttk.Scrollbar(log_text_frame, orient="vertical", command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        # Clear log button
        ttk.Button(self.log_frame, text="Clear Log", 
                  command=lambda: self.log_text.delete(1.0, tk.END)).grid(row=1, column=0, pady=5)
    
    def log_message(self, message):
        """Add message to log tab"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Update status bar"""
        self.status_text.set(message)
        self.root.update_idletasks()
    
    def load_config(self):
        """Load configuration from file"""
        config_file = "gui_config.json"
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
                    self.doc_id.set(config.get('doc_id', ''))
                    self.table_id.set(config.get('table_id', ''))
            except Exception as e:
                print(f"Could not load config: {e}")
    
    def save_config(self):
        """Save current configuration"""
        config = {
            'doc_id': self.doc_id.get(),
            'table_id': self.table_id.get()
        }
        try:
            with open("gui_config.json", 'w') as f:
                json.dump(config, f, indent=2)
            messagebox.showinfo("Success", "Configuration saved!")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save configuration: {e}")
    
    def load_from_env(self):
        """Load configuration from .env file"""
        env_file = ".env"
        if not os.path.exists(env_file):
            messagebox.showerror("Error", ".env file not found!")
            return
        
        try:
            with open(env_file, 'r') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith('CODA_API_TOKEN='):
                        self.api_token.set(line.split('=', 1)[1])
                    elif line.startswith('CODA_DOC_ID='):
                        self.doc_id.set(line.split('=', 1)[1])
                    elif line.startswith('CODA_TABLE_ID='):
                        self.table_id.set(line.split('=', 1)[1])
            
            messagebox.showinfo("Success", "Configuration loaded from .env file!")
            self.log_message("Configuration loaded from .env file")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not load .env file: {e}")
    
    def validate_inputs(self):
        """Validate that required inputs are provided"""
        if not self.api_token.get():
            messagebox.showerror("Error", "API Token is required!")
            return False
        if not self.doc_id.get():
            messagebox.showerror("Error", "Document ID is required!")
            return False
        if not self.table_id.get():
            messagebox.showerror("Error", "Table ID is required!")
            return False
        return True
    
    def _create_extractor(self):
        """Create an extractor with current GUI settings"""
        # Temporarily set environment variables for the extractor
        original_token = os.environ.get('CODA_API_TOKEN')
        original_doc = os.environ.get('CODA_DOC_ID') 
        original_table = os.environ.get('CODA_TABLE_ID')
        
        os.environ['CODA_API_TOKEN'] = self.api_token.get()
        os.environ['CODA_DOC_ID'] = self.doc_id.get() or 'dummy'
        os.environ['CODA_TABLE_ID'] = self.table_id.get() or 'dummy'
        
        try:
            extractor = CodaTimesheetExtractor()
            # Override with GUI values
            extractor.api_token = self.api_token.get()
            extractor.headers = {
                "Authorization": f"Bearer {self.api_token.get()}",
                "Content-Type": "application/json"
            }
            return extractor
        finally:
            # Restore original environment
            if original_token:
                os.environ['CODA_API_TOKEN'] = original_token
            else:
                os.environ.pop('CODA_API_TOKEN', None)
                
            if original_doc:
                os.environ['CODA_DOC_ID'] = original_doc
            else:
                os.environ.pop('CODA_DOC_ID', None)
                
            if original_table:
                os.environ['CODA_TABLE_ID'] = original_table
            else:
                os.environ.pop('CODA_TABLE_ID', None)
    
    def list_documents(self):
        """List available documents"""
        if not self.api_token.get():
            messagebox.showerror("Error", "API Token is required!")
            return
        
        thread = threading.Thread(target=self._list_documents_thread)
        thread.start()
    
    def _list_documents_thread(self):
        """Background thread for listing documents"""
        self.progress.start()
        self.update_status("Fetching documents...")
        
        try:
            extractor = self._create_extractor()
            docs = extractor.get_documents()
            
            # Show results in a new window
            self.show_documents_window(docs)
            
            self.log_message(f"Found {len(docs.get('items', []))} documents")
            self.update_status("Documents fetched successfully")
            
        except Exception as e:
            self.log_message(f"Error fetching documents: {e}")
            self.update_status("Error fetching documents")
            messagebox.showerror("Error", f"Could not fetch documents: {e}")
        finally:
            self.progress.stop()
    
    def show_documents_window(self, docs):
        """Show documents in a popup window"""
        docs_window = tk.Toplevel(self.root)
        docs_window.title("Available Documents")
        docs_window.geometry("600x400")
        
        # Create treeview for documents
        frame = ttk.Frame(docs_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(frame, columns=("name", "id"), show="headings")
        tree.heading("name", text="Document Name")
        tree.heading("id", text="Document ID")
        tree.column("name", width=300)
        tree.column("id", width=200)
        
        # Add documents to tree
        for doc in docs.get('items', []):
            tree.insert("", tk.END, values=(doc['name'], doc['id']))
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        # Double-click to select
        def on_select(event):
            if tree.selection():
                selection = tree.selection()[0]
                doc_id = tree.item(selection)['values'][1]
                self.doc_id.set(doc_id)
                docs_window.destroy()
        
        tree.bind("<Double-1>", on_select)
        
        ttk.Label(frame, text="Double-click a document to select it").pack(pady=5)
    
    def list_tables(self):
        """List tables in the selected document"""
        if not self.api_token.get() or not self.doc_id.get():
            messagebox.showerror("Error", "API Token and Document ID are required!")
            return
        
        thread = threading.Thread(target=self._list_tables_thread)
        thread.start()
    
    def _list_tables_thread(self):
        """Background thread for listing tables"""
        self.progress.start()
        self.update_status("Fetching tables...")
        
        try:
            extractor = self._create_extractor()
            tables = extractor.get_tables(self.doc_id.get())
            
            # Show results in a new window
            self.show_tables_window(tables)
            
            self.log_message(f"Found {len(tables.get('items', []))} tables")
            self.update_status("Tables fetched successfully")
            
        except Exception as e:
            self.log_message(f"Error fetching tables: {e}")
            self.update_status("Error fetching tables")
            messagebox.showerror("Error", f"Could not fetch tables: {e}")
        finally:
            self.progress.stop()
    
    def show_tables_window(self, tables):
        """Show tables in a popup window"""
        tables_window = tk.Toplevel(self.root)
        tables_window.title("Available Tables")
        tables_window.geometry("600x400")
        
        frame = ttk.Frame(tables_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(frame, columns=("name", "id"), show="headings")
        tree.heading("name", text="Table Name")
        tree.heading("id", text="Table ID")
        tree.column("name", width=300)
        tree.column("id", width=200)
        
        for table in tables.get('items', []):
            tree.insert("", tk.END, values=(table['name'], table['id']))
        
        tree.pack(fill=tk.BOTH, expand=True)
        
        def on_select(event):
            if tree.selection():
                selection = tree.selection()[0]
                table_id = tree.item(selection)['values'][1]
                self.table_id.set(table_id)
                tables_window.destroy()
        
        tree.bind("<Double-1>", on_select)
        
        ttk.Label(frame, text="Double-click a table to select it").pack(pady=5)
    
    def extract_data(self):
        """Extract timesheet data"""
        if not self.validate_inputs():
            return
        
        thread = threading.Thread(target=self._extract_data_thread)
        thread.start()
    
    def _extract_data_thread(self):
        """Background thread for data extraction"""
        self.progress.start()
        self.update_status("Extracting data...")
        
        try:
            extractor = self._create_extractor()
            
            self.log_message("Starting data extraction...")
            
            # Extract data
            raw_data = extractor.get_timesheet_data(self.doc_id.get(), self.table_id.get())
            self.log_message(f"Extracted {len(raw_data.get('items', []))} rows")
            
            # Process data
            processor = TimesheetProcessor()
            df = processor.process_raw_data(raw_data)
            df_cleaned = processor.clean_timesheet_data(df)
            
            # Store data
            self.current_df = df_cleaned
            
            # Update UI
            self.display_data(df_cleaned)
            
            # Generate summary
            summary = processor.generate_summary(df_cleaned)
            summary_text = f"Rows: {summary['total_rows']}, Columns: {len(summary['columns'])}"
            if summary.get('total_hours'):
                summary_text += f", Total Hours: {summary['total_hours']:.1f}"
            if summary.get('date_range'):
                summary_text += f", Date Range: {summary['date_range']}"
            
            self.summary_label.config(text=summary_text)
            
            # Enable export buttons
            self.export_csv_btn.config(state='normal')
            self.export_excel_btn.config(state='normal')
            
            self.log_message("Data extraction completed successfully!")
            self.update_status(f"Extracted {len(df_cleaned)} rows successfully")
            
            messagebox.showinfo("Success", f"Successfully extracted {len(df_cleaned)} rows!")
            
        except Exception as e:
            self.log_message(f"Error during extraction: {e}")
            self.update_status("Extraction failed")
            messagebox.showerror("Error", f"Extraction failed: {e}")
        finally:
            self.progress.stop()
    
    def display_data(self, df):
        """Display data in the grid or treeview"""
        if HAS_TKSHEET:
            self.display_data_sheet(df)
        else:
            self.display_data_tree(df)
    
    def display_data_sheet(self, df):
        """Display data in tksheet grid"""
        if df.empty:
            return
        
        # Clear existing data
        self.sheet.set_sheet_data([])
        
        # Set headers
        headers = list(df.columns)
        self.sheet.headers(headers)
        
        # Convert DataFrame to list of lists
        data = df.values.tolist()
        
        # Handle NaN values
        for i, row in enumerate(data):
            for j, cell in enumerate(row):
                if pd.isna(cell):
                    data[i][j] = ""
                else:
                    data[i][j] = str(cell)
        
        # Set data
        self.sheet.set_sheet_data(data)
        
        # Auto-resize columns
        self.sheet.set_all_column_widths()
    
    def display_data_tree(self, df):
        """Display data in treeview (fallback)"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if df.empty:
            return
        
        # Configure columns
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"
        
        # Set column headings and widths
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, minwidth=50)
        
        # Insert data
        for index, row in df.iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in columns]
            self.tree.insert("", tk.END, values=values)
    
    def export_csv(self):
        """Export data to CSV"""
        if self.current_df is None:
            messagebox.showerror("Error", "No data to export!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialname=f"timesheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        
        if filename:
            try:
                self.current_df.to_csv(filename, index=False)
                self.log_message(f"Data exported to: {filename}")
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not export CSV: {e}")
    
    def export_excel(self):
        """Export data to Excel"""
        if self.current_df is None:
            messagebox.showerror("Error", "No data to export!")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialname=f"timesheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if filename:
            try:
                self.current_df.to_excel(filename, index=False)
                self.log_message(f"Data exported to: {filename}")
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not export Excel: {e}")


def main():
    """Main function to run the GUI"""
    root = tk.Tk()
    app = TimesheetExtractorGUI(root)
    
    # Handle window closing
    def on_closing():
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Start the GUI
    root.mainloop()


if __name__ == "__main__":
    main()