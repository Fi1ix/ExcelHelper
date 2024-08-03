import tkinter as tk
from tkinter import ttk
import subprocess

class ToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Main Window")
        self.center_window()

        # 创建 Notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill=tk.BOTH)

        # 添加第一个标签页 Detail
        self.tab1 = tk.Frame(self.notebook)
        self.notebook.add(self.tab1, text='Detail')
        self.create_detail_widgets()

        # 添加第二个标签页 导出数据
        self.tab2 = tk.Frame(self.notebook)
        self.notebook.add(self.tab2, text='导出数据')

        # 添加第三个标签页 批量写入数据
        self.tab3 = tk.Frame(self.notebook)
        self.notebook.add(self.tab3, text='批量写入数据')

        # 绑定标签页切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def center_window(self):
        # 获取屏幕宽高
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 计算窗口大小和位置
        window_width = 800  # 自定义窗口宽度
        window_height = 600  # 自定义窗口高度

        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')

    def create_detail_widgets(self):
        # 在第一个标签页（Detail）创建信息展示部件
        label = tk.Label(self.tab1, text="Version Test\n\n"
                                         "Tab1     用来导出数据，可以根据名称进行过滤\n"
                                         "Tab2     用来写入数据，根据已有的txt标进原来的表格",
                         justify='center', font=('Helvetica', 12), padx=20, pady=20)
        label.pack(expand=True)

    def on_tab_changed(self, event):
        # 获取当前选中的标签页索引
        current_tab = self.notebook.index("current")

        # 根据标签页索引执行不同的操作
        if current_tab == 1:
            # 如果切换到第二个标签页，运行 1.py 程序
            subprocess.Popen(['python', '1.py'])
        elif current_tab == 2:
            # 如果切换到第三个标签页，运行 2.py 程序
            subprocess.Popen(['python', '2.py'])

if __name__ == "__main__":
    root = tk.Tk()
    app = ToolsApp(root)
    root.mainloop()
