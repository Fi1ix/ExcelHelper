import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("表格处理工具")
        self.master.geometry("600x400")
        self.center_window()  # 居中显示窗口
        self.pack()
        self.create_widgets()
        self.file_path_excel = None
        self.file_path_txt = None

    def center_window(self):
        # 居中显示窗口
        window_width = self.master.winfo_reqwidth()
        window_height = self.master.winfo_reqheight()
        position_right = int(self.master.winfo_screenwidth() / 2 - window_width / 2)
        position_down = int(self.master.winfo_screenheight() / 2 - window_height / 2)
        self.master.geometry(f"+{position_right}+{position_down}")

    def create_widgets(self):
        # Label and Entry for 增加列名
        self.label_input1 = tk.Label(self, text="增加列名:")
        self.label_input1.grid(row=0, column=0, padx=10, pady=10)
        self.entry_input1 = tk.Entry(self)
        self.entry_input1.grid(row=0, column=1, padx=10, pady=10)

        # Label and Entry for 增加数据
        self.label_input2 = tk.Label(self, text="增加数据:")
        self.label_input2.grid(row=1, column=0, padx=10, pady=10)
        self.entry_input2 = tk.Entry(self)
        self.entry_input2.grid(row=1, column=1, padx=10, pady=10)

        # Label and Entry for 要对比的列
        self.label_input3 = tk.Label(self, text="要对比的列:")
        self.label_input3.grid(row=2, column=0, padx=10, pady=10)
        self.entry_input3 = tk.Entry(self)
        self.entry_input3.grid(row=2, column=1, padx=10, pady=10)

        # Button to import excel table
        self.import_excel_button = tk.Button(self, text="导入Excel表格", command=self.import_excel)
        self.import_excel_button.grid(row=3, column=0, padx=10, pady=10)

        # Button to import txt table
        self.import_txt_button = tk.Button(self, text="导入Txt表格", command=self.import_txt)
        self.import_txt_button.grid(row=3, column=1, padx=10, pady=10)

        # Button to process data
        self.process_button = tk.Button(self, text="确定", command=self.process_data)
        self.process_button.grid(row=4, columnspan=2, padx=10, pady=10)

        # Text widget to display file paths
        self.text_display = tk.Text(self, height=6, width=50)
        self.text_display.grid(row=5, columnspan=2, padx=10, pady=10)

    def import_excel(self):
        self.file_path_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.display_file_path(self.file_path_excel)

    def import_txt(self):
        self.file_path_txt = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        self.display_file_path(self.file_path_txt)

    def display_file_path(self, file_path):
        self.text_display.delete(1.0, tk.END)
        if file_path:
            self.text_display.insert(tk.END, f"导入文件：{file_path}\n")
        else:
            self.text_display.insert(tk.END, "未选择文件\n")

    def process_data(self):
        add_column_name = self.entry_input1.get()
        add_data = self.entry_input2.get()
        compare_column = self.entry_input3.get()

        if not (add_column_name and add_data and compare_column):
            messagebox.showerror("错误", "请输入完整信息")
            return

        if not (self.file_path_excel and self.file_path_txt):
            messagebox.showerror("错误", "请先导入Excel表格和Txt表格")
            return

        try:
            df_excel = pd.read_excel(self.file_path_excel)
            df_txt = pd.read_csv(self.file_path_txt, sep='\t' if self.file_path_txt.endswith('.txt') else ',')

            if compare_column not in df_excel.columns:
                messagebox.showerror("错误", f"表格中不存在列 '{compare_column}'")
                return

            df_excel[add_column_name] = ''

            for index, row in df_txt.iterrows():
                value_to_find = row[0]  # 假设txt文件中第一列是要查找的值
                mask = df_excel[compare_column] == value_to_find
                if not mask.any():
                    continue
                df_excel.loc[mask, add_column_name] = add_data

            output_file = os.path.splitext(self.file_path_excel)[0] + '_processed.xlsx'
            df_excel.to_excel(output_file, index=False)
            messagebox.showinfo("完成", f"处理完成！导出文件：{output_file}")

        except Exception as e:
            messagebox.showerror("错误", f"处理数据出错: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()
