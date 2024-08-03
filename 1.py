import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re

class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Filter and Export Tool")
        self.root.geometry("400x400")
        
        self.file_path = ""
        self.df = None
        self.export_var = tk.StringVar(self.root)
        self.filter_var = tk.StringVar(self.root)
        self.apply_filter = tk.BooleanVar(self.root, value=False)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Label for file selection
        self.label_file = tk.Label(self.root, text="Select Excel/CSV file:")
        self.label_file.pack(pady=10)
        
        # Button to select file
        self.btn_browse = tk.Button(self.root, text="Browse", command=self.browse_file)
        self.btn_browse.pack()
        
        # Label to display selected file
        self.file_label = tk.Label(self.root, text="Selected file: None")
        self.file_label.pack(pady=10)
        
        # Dropdown for export selection
        self.label_export = tk.Label(self.root, text="Select data to export:")
        self.label_export.pack(pady=10)
        
        self.export_dropdown = tk.OptionMenu(self.root, self.export_var, "")
        self.export_dropdown.pack()
        
        # Checkbox for applying filter
        self.apply_filter_checkbox = tk.Checkbutton(self.root, text="Apply Filter", variable=self.apply_filter, command=self.toggle_filter)
        self.apply_filter_checkbox.pack(pady=5)
        
        # Dropdown for filter selection
        self.filter_label = tk.Label(self.root, text="Select filter column:")
        self.filter_label.pack()
        
        self.filter_dropdown = tk.OptionMenu(self.root, self.filter_var, "")
        self.filter_dropdown.pack()
        self.filter_dropdown.config(state=tk.DISABLED)  # Initially disabled
        
        # Export button
        self.btn_export = tk.Button(self.root, text="Export", command=self.export_data)
        self.btn_export.pack(pady=20)
        
        # Center the window on screen
        self.center_window()
        
    def center_window(self):
        # Function to center the window on the screen
        window_width = self.root.winfo_reqwidth()
        window_height = self.root.winfo_reqheight()
        
        position_right = int(self.root.winfo_screenwidth()/2 - window_width/2)
        position_down = int(self.root.winfo_screenheight()/2 - window_height/2)
        
        self.root.geometry("+{}+{}".format(position_right, position_down))
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx"), ("CSV Files", "*.csv")])
        
        if self.file_path:
            self.load_data()
            self.update_dropdowns()
            self.file_label.config(text=f"Selected file: {self.file_path}")
        else:
            messagebox.showwarning("Warning", "No file selected.")
    
    def load_data(self):
        try:
            # Use pandas to read the selected file into a DataFrame
            if self.file_path.endswith('.csv'):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading file: {str(e)}")
            self.file_path = ""
            self.df = None
    
    def update_dropdowns(self):
        if self.df is not None:
            # Get column names from the first row of the DataFrame
            columns = list(self.df.columns)
            
            # Update export dropdown with column names
            self.export_var.set("")
            menu_export = self.export_dropdown["menu"]
            menu_export.delete(0, "end")
            for column in columns:
                menu_export.add_command(label=column, command=lambda value=column: self.export_var.set(value))
            
            # Update filter dropdown with column names (excluding export column)
            self.filter_var.set("")
            menu_filter = self.filter_dropdown["menu"]
            menu_filter.delete(0, "end")
            for column in columns:
                if column != self.export_var.get():  # Exclude export column from filter options
                    menu_filter.add_command(label=column, command=lambda value=column: self.filter_var.set(value))
                
            # Enable filter dropdown if apply_filter is checked
            self.toggle_filter()
        else:
            messagebox.showwarning("Warning", "No data loaded.")
    
    def toggle_filter(self):
        if self.apply_filter.get():
            self.filter_dropdown.config(state=tk.NORMAL)
        else:
            self.filter_dropdown.config(state=tk.DISABLED)
    
    def export_data(self):
        if self.df is None:
            messagebox.showwarning("Warning", "No data loaded.")
            return
        
        export_selection = self.export_var.get()
        filter_selection = self.filter_var.get()
        
        try:
            if export_selection:
                if self.apply_filter.get() and filter_selection:
                    # Apply filter and export filtered data to corresponding files
                    unique_values = self.df[filter_selection].unique()
                    
                    for value in unique_values:
                        filtered_data = self.df[self.df[filter_selection] == value][export_selection].tolist()
                        filename = f"{value}_{export_selection}.txt"
                        
                        # Replace invalid characters in filename
                        filename = re.sub(r'[\/?]', '_', filename)
                        
                        with open(filename, "w") as f:
                            for data in filtered_data:
                                f.write(f"{data}\n")
                        
                    messagebox.showinfo("Export Successful", "Data exported successfully.")
                else:
                    # Export all data to a single file if no filter is applied
                    all_data = self.df[export_selection].tolist()
                    filename = f"all_{export_selection}.txt"
                    
                    # Replace invalid characters in filename
                    filename = re.sub(r'[\/?]', '_', filename)
                    
                    with open(filename, "w") as f:
                        for data in all_data:
                            f.write(f"{data}\n")
                    
                    messagebox.showinfo("Export Successful", "Data exported successfully.")
            else:
                messagebox.showwarning("Warning", "Please select data to export.")
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting data: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()
