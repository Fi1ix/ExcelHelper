import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class TableProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Table Processor")

        # 设置窗口的大小和位置
        window_width = 600
        window_height = 400
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 创建界面组件
        self.label_file = tk.Label(self.root, text="导入的表格：", font=("Arial", 16))
        self.label_file.pack(pady=10)

        self.label_filename = tk.Label(self.root, text="", font=("Arial", 14))
        self.label_filename.pack()

        self.button_import = tk.Button(self.root, text="选择表格文件", command=self.import_table, font=("Arial", 14))
        self.button_import.pack(pady=10)

        self.label_txt = tk.Label(self.root, text="导入的txt文件：", font=("Arial", 16))
        self.label_txt.pack(pady=10)

        self.label_txt_filename = tk.Label(self.root, text="", font=("Arial", 14))
        self.label_txt_filename.pack()

        self.button_import_txt = tk.Button(self.root, text="选择txt文件", command=self.import_txt, font=("Arial", 14))
        self.button_import_txt.pack(pady=10)

        self.button_process = tk.Button(self.root, text="确定", command=self.process_data, font=("Arial", 16))
        self.button_process.pack(pady=20)

        self.label_status = tk.Label(self.root, text="", font=("Arial", 14))
        self.label_status.pack()

        self.selected_table = None
        self.txt_lines = []

    def import_table(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if filename:
            try:
                self.selected_table = pd.read_excel(filename)  # Assuming it's an Excel file
                self.label_filename.config(text=filename)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import table: {e}")

    def import_txt(self):
        filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as file:
                    self.txt_lines = [line.strip() for line in file.readlines()]
                self.label_txt_filename.config(text=filename)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import txt file: {e}")

    def process_data(self):
        if self.selected_table is None:
            messagebox.showerror("Error", "Please import a table first.")
            return
        if not self.txt_lines:
            messagebox.showerror("Error", "Please import a txt file first.")
            return

        deleted_indices = []
        for txt_line in self.txt_lines:
            # Check if any cell in the table contains txt_line as a substring
            rows_to_delete = self.selected_table[self.selected_table.apply(lambda row: row.astype(str).str.contains(txt_line).any(), axis=1)].index
            deleted_indices.extend(rows_to_delete)

        # Remove duplicates from indices list
        deleted_indices = list(set(deleted_indices))

        # Delete rows from selected_table
        self.selected_table.drop(deleted_indices, inplace=True)

        # Reset index to avoid empty rows
        self.selected_table.reset_index(drop=True, inplace=True)

        # Save processed table to a new file
        output_filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_filename:
            try:
                self.selected_table.to_excel(output_filename, index=False)
                self.label_status.config(text=f"Deleted {len(deleted_indices)} rows. Saved to {output_filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save output file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TableProcessorApp(root)
    root.mainloop()
