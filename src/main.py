import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from processor import process_files, find_job_pos
import threading
from datetime import datetime

class ProdSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ProdSync - Garments Intelligence System")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700)
        
        # Variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.job_entry = tk.StringVar()
        self.sheets_list = []
        self.df2 = None  # Store Data Sheet 2 for job lookup
        
        # Colors - Clean modern look
        self.bg_color = "#ffffff"
        self.primary_color = "#1e293b"
        self.secondary_color = "#2563eb"
        self.accent_color = "#16a34a"
        self.text_color = "#334155"
        self.border_color = "#e2e8f0"
        self.console_bg = "#0f172a"
        self.console_fg = "#e2e8f0"
        
        # Configure root window
        self.root.configure(bg=self.bg_color)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main container with padding
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header Section
        header_frame = tk.Frame(main_container, bg=self.bg_color)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Title
        title_frame = tk.Frame(header_frame, bg=self.bg_color)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        title_label = tk.Label(title_frame, 
                              text="ProdSync",
                              font=('Segoe UI', 24, 'bold'),
                              fg=self.primary_color,
                              bg=self.bg_color)
        title_label.pack(anchor='w')
        
        subtitle_label = tk.Label(title_frame, 
                                 text="Garments Intelligence System",
                                 font=('Segoe UI', 11),
                                 fg=self.secondary_color,
                                 bg=self.bg_color)
        subtitle_label.pack(anchor='w')
        
        # Developer info
        dev_frame = tk.Frame(header_frame, bg=self.bg_color)
        dev_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        dev_name = tk.Label(dev_frame,
                           text="Developed by Rakib Hasan",
                           font=('Segoe UI', 9, 'italic'),
                           fg=self.text_color,
                           bg=self.bg_color)
        dev_name.pack(anchor='e')
        
        copyright_label = tk.Label(dev_frame,
                                  text="© 2026 All Rights Reserved",
                                  font=('Segoe UI', 8),
                                  fg='#94a3b8',
                                  bg=self.bg_color)
        copyright_label.pack(anchor='e')
        
        # Separator
        separator = tk.Frame(main_container, height=2, bg=self.border_color)
        separator.pack(fill=tk.X, pady=(0, 15))
        
        # File Selection Section
        file_frame = tk.Frame(main_container, bg=self.bg_color, highlightbackground=self.border_color, highlightthickness=1)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Section header
        file_header = tk.Frame(file_frame, bg=self.bg_color)
        file_header.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        tk.Label(file_header,
                text="📁 File Selection",
                font=('Segoe UI', 12, 'bold'),
                fg=self.primary_color,
                bg=self.bg_color).pack(anchor='w')
        
        # Content
        content_frame = tk.Frame(file_frame, bg=self.bg_color)
        content_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        # File 1 selection
        file1_row = tk.Frame(content_frame, bg=self.bg_color)
        file1_row.pack(fill=tk.X, pady=5)
        
        tk.Label(file1_row, 
                text="Buyer Orders File:",
                font=('Segoe UI', 10),
                width=18,
                anchor='w',
                fg=self.text_color,
                bg=self.bg_color).pack(side=tk.LEFT)
        
        file1_entry = tk.Entry(file1_row, textvariable=self.file1_path, 
                              font=('Segoe UI', 9),
                              bg='#f8fafc',
                              fg=self.text_color,
                              relief='solid',
                              borderwidth=1)
        file1_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        tk.Button(file1_row, 
                 text="Browse",
                 font=('Segoe UI', 9),
                 bg=self.secondary_color,
                 fg='white',
                 relief='flat',
                 padx=15,
                 pady=2,
                 cursor='hand2',
                 command=self.select_file1).pack(side=tk.LEFT)
        
        # File 2 selection
        file2_row = tk.Frame(content_frame, bg=self.bg_color)
        file2_row.pack(fill=tk.X, pady=5)
        
        tk.Label(file2_row, 
                text="Production Data File:",
                font=('Segoe UI', 10),
                width=18,
                anchor='w',
                fg=self.text_color,
                bg=self.bg_color).pack(side=tk.LEFT)
        
        file2_entry = tk.Entry(file2_row, textvariable=self.file2_path, 
                              font=('Segoe UI', 9),
                              bg='#f8fafc',
                              fg=self.text_color,
                              relief='solid',
                              borderwidth=1)
        file2_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        tk.Button(file2_row, 
                 text="Browse",
                 font=('Segoe UI', 9),
                 bg=self.secondary_color,
                 fg='white',
                 relief='flat',
                 padx=15,
                 pady=2,
                 cursor='hand2',
                 command=self.select_file2).pack(side=tk.LEFT)
        
        # Sheet selection
        sheet_row = tk.Frame(content_frame, bg=self.bg_color)
        sheet_row.pack(fill=tk.X, pady=(10, 5))
        
        tk.Label(sheet_row, 
                text="Select Buyer Sheet:",
                font=('Segoe UI', 10),
                width=18,
                anchor='w',
                fg=self.text_color,
                bg=self.bg_color).pack(side=tk.LEFT)
        
        self.sheet_combo = ttk.Combobox(sheet_row, textvariable=self.selected_sheet,
                                        values=[], width=40, state='readonly',
                                        font=('Segoe UI', 9))
        self.sheet_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(sheet_row, 
                 text="📂 Load Sheets",
                 font=('Segoe UI', 9),
                 bg='#f1f5f9',
                 fg=self.text_color,
                 relief='flat',
                 padx=15,
                 pady=2,
                 cursor='hand2',
                 command=self.load_sheets).pack(side=tk.LEFT)
        
        # Job Lookup Section
        job_frame = tk.Frame(main_container, bg=self.bg_color, highlightbackground=self.border_color, highlightthickness=1)
        job_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Section header
        job_header = tk.Frame(job_frame, bg=self.bg_color)
        job_header.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        tk.Label(job_header,
                text="🔍 Job Lookup",
                font=('Segoe UI', 12, 'bold'),
                fg=self.primary_color,
                bg=self.bg_color).pack(anchor='w')
        
        # Content
        job_content = tk.Frame(job_frame, bg=self.bg_color)
        job_content.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        # Job entry
        job_entry_row = tk.Frame(job_content, bg=self.bg_color)
        job_entry_row.pack(fill=tk.X)
        
        tk.Label(job_entry_row, 
                text="Enter Job Number:",
                font=('Segoe UI', 10),
                width=18,
                anchor='w',
                fg=self.text_color,
                bg=self.bg_color).pack(side=tk.LEFT)
        
        self.job_entry_widget = tk.Entry(job_entry_row, textvariable=self.job_entry, 
                                        font=('Segoe UI', 11),
                                        bg='#f8fafc',
                                        fg=self.text_color,
                                        relief='solid',
                                        borderwidth=1)
        self.job_entry_widget.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        
        self.find_po_btn = tk.Button(job_entry_row, 
                                     text="🔎 Find POs",
                                     font=('Segoe UI', 10, 'bold'),
                                     bg=self.secondary_color,
                                     fg='white',
                                     relief='flat',
                                     padx=20,
                                     pady=5,
                                     cursor='hand2',
                                     command=self.find_pos, 
                                     state='disabled')
        self.find_po_btn.pack(side=tk.LEFT)
        
        # Example text
        example_label = tk.Label(job_content, 
                                text="Examples: 196, 240, 263, SGL-25-00196",
                                font=('Segoe UI', 9, 'italic'),
                                fg='#94a3b8',
                                bg=self.bg_color)
        example_label.pack(anchor='w', padx=(145, 0), pady=(5, 0))
        
        # Results Section
        results_frame = tk.Frame(main_container, bg=self.bg_color, highlightbackground=self.border_color, highlightthickness=1)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Section header
        results_header = tk.Frame(results_frame, bg=self.bg_color)
        results_header.pack(fill=tk.X, padx=15, pady=(10, 5))
        
        tk.Label(results_header,
                text="📊 Matching POs",
                font=('Segoe UI', 12, 'bold'),
                fg=self.primary_color,
                bg=self.bg_color).pack(side=tk.LEFT)
        
        # Clear button
        self.clear_btn = tk.Button(results_header, 
                                   text="🗑️ Clear",
                                   font=('Segoe UI', 9),
                                   bg='#f1f5f9',
                                   fg=self.text_color,
                                   relief='flat',
                                   padx=10,
                                   pady=2,
                                   cursor='hand2',
                                   command=self.clear_results)
        self.clear_btn.pack(side=tk.RIGHT)
        
        # Treeview with scrollbars
        tree_container = tk.Frame(results_frame, bg=self.bg_color)
        tree_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        
        # Create treeview
        columns = ('Order No', 'Style', 'Color', 'Qty', 'Ship Date')
        self.po_tree = ttk.Treeview(tree_container, columns=columns, 
                                     show='headings', height=8,
                                     selectmode='browse')
        
        # Define headings
        for col in columns:
            self.po_tree.heading(col, text=col, anchor='w')
            self.po_tree.column(col, width=120, anchor='w')
        
        # Add scrollbars
        vsb = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.po_tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.po_tree.xview)
        self.po_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Grid layout
        self.po_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Action Buttons
        action_frame = tk.Frame(main_container, bg=self.bg_color)
        action_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.process_btn = tk.Button(action_frame, 
                                     text="⚙️ Process Full File",
                                     font=('Segoe UI', 11, 'bold'),
                                     bg=self.accent_color,
                                     fg='white',
                                     relief='flat',
                                     padx=25,
                                     pady=8,
                                     cursor='hand2',
                                     command=self.process_files, 
                                     state='disabled')
        self.process_btn.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(action_frame, mode='indeterminate',
                                        length=200)
        self.progress.pack(side=tk.LEFT, padx=(15, 0))
        
        # Console Panel (Now Visible)
        console_frame = tk.Frame(main_container, bg=self.console_bg, highlightbackground=self.border_color, highlightthickness=1)
        console_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Console header
        console_header = tk.Frame(console_frame, bg=self.console_bg)
        console_header.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        tk.Label(console_header,
                text="📋 Console",
                font=('Segoe UI', 10, 'bold'),
                fg='#ffffff',
                bg=self.console_bg).pack(side=tk.LEFT)
        
        # Clear console button
        tk.Button(console_header,
                 text="Clear",
                 font=('Segoe UI', 8),
                 bg='#334155',
                 fg='#ffffff',
                 relief='flat',
                 padx=10,
                 cursor='hand2',
                 command=self.clear_console).pack(side=tk.RIGHT)
        
        # Console text with scrollbar
        console_container = tk.Frame(console_frame, bg=self.console_bg)
        console_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # Create console text widget
        self.console_text = tk.Text(console_container, 
                                    height=8,
                                    font=('Consolas', 10),
                                    wrap=tk.WORD,
                                    bg=self.console_bg,
                                    fg=self.console_fg,
                                    relief='flat',
                                    borderwidth=0)
        self.console_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Console scrollbar
        console_scroll = ttk.Scrollbar(console_container, orient=tk.VERTICAL, 
                                       command=self.console_text.yview)
        console_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.console_text.config(yscrollcommand=console_scroll.set)
        
        # Initial console message
        self.log_to_console("ProdSync - Garments Intelligence System initialized")
        self.log_to_console("Ready to process files...")
        
        # Footer
        footer_frame = tk.Frame(main_container, bg=self.bg_color)
        footer_frame.pack(fill=tk.X)
        
        tk.Label(footer_frame,
                text="Version 1.0.0",
                font=('Segoe UI', 8),
                fg='#94a3b8',
                bg=self.bg_color).pack(side=tk.RIGHT)
        
    def log_to_console(self, message, message_type="info"):
        """Add timestamped message to console with color coding"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Configure tags for different message types
        self.console_text.tag_configure("info", foreground=self.console_fg)
        self.console_text.tag_configure("success", foreground="#4ade80")
        self.console_text.tag_configure("error", foreground="#f87171")
        self.console_text.tag_configure("warning", foreground="#fbbf24")
        
        self.console_text.insert(tk.END, f"[{timestamp}] ", "info")
        self.console_text.insert(tk.END, f"{message}\n", message_type)
        self.console_text.see(tk.END)
        self.root.update()
        
    def clear_console(self):
        """Clear the console"""
        self.console_text.delete(1.0, tk.END)
        self.log_to_console("Console cleared")
        
    def select_file1(self):
        filename = filedialog.askopenfilename(
            title="Select Buyer Orders File",
            filetypes=[("Excel files", "*.xlsx *.xls"), 
                      ("CSV files", "*.csv"), 
                      ("All files", "*.*")]
        )
        if filename:
            self.file1_path.set(filename)
            self.update_buttons()
            self.log_to_console(f"✓ Selected buyer file: {os.path.basename(filename)}", "success")
            
    def select_file2(self):
        filename = filedialog.askopenfilename(
            title="Select Production Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), 
                      ("CSV files", "*.csv"), 
                      ("All files", "*.*")]
        )
        if filename:
            self.file2_path.set(filename)
            # Load Data Sheet 2 for job lookup
            self.load_data_sheet2()
            self.update_buttons()
            self.log_to_console(f"✓ Selected production file: {os.path.basename(filename)}", "success")
            
    def load_data_sheet2(self):
        """Load Data Sheet 2 for job lookup functionality"""
        try:
            self.log_to_console("Loading production data for job lookup...", "info")
            
            file_ext2 = os.path.splitext(self.file2_path.get())[1].lower()
            
            if file_ext2 == '.csv':
                self.df2 = pd.read_csv(self.file2_path.get())
            else:
                self.df2 = pd.read_excel(self.file2_path.get())
            
            self.log_to_console(f"✅ Loaded {len(self.df2)} rows from production data", "success")
            
        except Exception as e:
            self.log_to_console(f"❌ Error loading production data: {str(e)}", "error")
            messagebox.showerror("Error", f"Failed to load production data: {str(e)}")
            
    def load_sheets(self):
        if not self.file1_path.get():
            messagebox.showerror("Error", "Please select buyer orders file first")
            return
            
        try:
            self.log_to_console("Loading sheets from buyer file...", "info")
            
            # Get file extension
            file_ext = os.path.splitext(self.file1_path.get())[1].lower()
            
            if file_ext == '.csv':
                self.sheets_list = ['Sheet1']  # CSV has only one sheet
            else:
                # Read Excel file to get sheet names
                xl = pd.ExcelFile(self.file1_path.get())
                self.sheets_list = xl.sheet_names
            
            self.sheet_combo['values'] = self.sheets_list
            if self.sheets_list:
                self.sheet_combo.current(0)
                self.selected_sheet.set(self.sheets_list[0])
                
            self.log_to_console(f"✅ Found {len(self.sheets_list)} sheets in buyer file", "success")
            self.update_buttons()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheets: {str(e)}")
            self.log_to_console(f"❌ Error: {str(e)}", "error")
            
    def find_pos(self):
        """Find POs for entered job number"""
        if self.df2 is None:
            messagebox.showerror("Error", "Please load production data first")
            return
            
        job_input = self.job_entry.get().strip()
        if not job_input:
            messagebox.showwarning("Warning", "Please enter a job number")
            return
            
        self.log_to_console(f"Searching for job: {job_input}", "info")
        
        # Clear previous results
        for item in self.po_tree.get_children():
            self.po_tree.delete(item)
        
        try:
            # Call the find_job_pos function from processor
            results = find_job_pos(self.df2, job_input)
            
            if results.empty:
                self.log_to_console(f"⚠️ No POs found for job: {job_input}", "warning")
                messagebox.showinfo("No Results", f"No POs found for job: {job_input}")
                return
            
            # Display results in treeview
            for _, row in results.iterrows():
                values = []
                for col in ['Order No', 'Style Name', 'Item Name', 'Order Qty.', 'Ship Date']:
                    if col in row.index:
                        # Format numbers nicely
                        val = row[col]
                        if col == 'Order Qty.' and pd.notna(val):
                            try:
                                val = f"{int(float(val)):,}"
                            except:
                                pass
                        values.append(val)
                    else:
                        values.append('')
                
                self.po_tree.insert('', tk.END, values=values)
            
            self.log_to_console(f"✅ Found {len(results)} POs for job: {job_input}", "success")
            
        except Exception as e:
            self.log_to_console(f"❌ Error finding POs: {str(e)}", "error")
            messagebox.showerror("Error", f"Failed to find POs: {str(e)}")
            import traceback
            traceback.print_exc()
            
    def clear_results(self):
        """Clear the PO results treeview"""
        for item in self.po_tree.get_children():
            self.po_tree.delete(item)
        self.log_to_console("Results cleared", "info")
            
    def update_buttons(self):
        """Update button states based on loaded files"""
        if self.file2_path.get() and self.df2 is not None:
            self.find_po_btn.config(state='normal')
        
        if (self.file1_path.get() and self.file2_path.get() and 
            self.selected_sheet.get()):
            self.process_btn.config(state='normal')
        else:
            self.process_btn.config(state='disabled')
            
    def process_files(self):
        # Ask for output file
        output_file = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not output_file:
            return
            
        # Disable buttons and start progress
        self.process_btn.config(state='disabled')
        self.find_po_btn.config(state='disabled')
        self.progress.start()
        
        # Run processing in separate thread
        thread = threading.Thread(target=self.run_processing, 
                                 args=(output_file,))
        thread.daemon = True
        thread.start()
        
    def run_processing(self, output_file):
        try:
            self.log_to_console("🚀 Starting file processing...", "info")
            self.log_to_console(f"📁 File 1: {os.path.basename(self.file1_path.get())}", "info")
            self.log_to_console(f"📁 File 2: {os.path.basename(self.file2_path.get())}", "info")
            self.log_to_console(f"📄 Selected sheet: {self.selected_sheet.get()}", "info")
            
            # Call the processor function
            result = process_files(
                file1_path=self.file1_path.get(),
                file2_path=self.file2_path.get(),
                sheet_name=self.selected_sheet.get(),
                output_path=output_file,
                status_callback=self.log_to_console
            )
            
            if result:
                self.root.after(0, self.processing_complete, output_file)
            else:
                self.root.after(0, self.processing_failed)
                
        except Exception as e:
            self.root.after(0, self.show_error, str(e))
            
    def processing_complete(self, output_file):
        self.progress.stop()
        self.process_btn.config(state='normal')
        self.find_po_btn.config(state='normal')
        self.log_to_console("✅ Processing completed successfully!", "success")
        messagebox.showinfo("Success", 
                           f"File processed successfully!\nSaved to:\n{output_file}")
        
    def processing_failed(self):
        self.progress.stop()
        self.process_btn.config(state='normal')
        self.find_po_btn.config(state='normal')
        self.log_to_console("❌ Processing failed. Check the console for details.", "error")
        
    def show_error(self, error_msg):
        self.progress.stop()
        self.process_btn.config(state='normal')
        self.find_po_btn.config(state='normal')
        self.log_to_console(f"❌ Error: {error_msg}", "error")
        messagebox.showerror("Error", f"Processing failed:\n{error_msg}")

def main():
    root = tk.Tk()
    app = ProdSyncApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
