import sys
import threading, time, requests, json, datetime
import tkinter as tk
from tkinter import filedialog

class honmeDialog:

    def __init__(self):
        self.root = tk.Tk()
        self.root.resizable(0, 0)
        self.root.attributes('-topmost', True)
        self.root.title("选择文件")
        self.root.geometry("600x500+500+100")

        self.tip = tk.Label(self.root, wraplength="400", text="文件未选择", fg="red", font=("宋体", 15))
        self.tip.place(x=90, y=18)
        # 当期支出明细原表
        tk.Label(self.root, text="当期支出明细原表：", font=("宋体", 15)).place(x=90, y=88)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select_1)
        self.select_button.place(x=290, y=87)

        # 当期日期
        tk.Label(self.root, text="当期日期：", font=("宋体", 15)).place(x=90, y=138)
        self.user_entry = tk.Entry(self.root, font=("宋体", 20), width=17)
        self.user_entry.place(x=290, y=137)

        # 去年同期支出明细原表
        tk.Label(self.root, text="同期支出明细原表：", font=("宋体", 15)).place(x=90, y=188)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select_2)
        self.select_button.place(x=290, y=187)

        # 去年同期日期
        tk.Label(self.root, text="同期日期：", font=("宋体", 15)).place(x=90, y=238)
        self.user_entry = tk.Entry(self.root, font=("宋体", 20), width=17)
        self.user_entry.place(x=290, y=237)

        # 当期劳产率表
        tk.Label(self.root, text="当期劳产率原表：", font=("宋体", 15)).place(x=90, y=288)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select_3)
        self.select_button.place(x=290, y=287)

        # 同期劳产率表
        tk.Label(self.root, text="当期劳产率原表：", font=("宋体", 15)).place(x=90, y=338)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select_4)
        self.select_button.place(x=290, y=337)


        # 确认和取消按钮
        self.ok_button = tk.Button(self.root, text="确定", width="10", font=("宋体", 15), command=self.on_ok)
        self.ok_button.place(x=290, y=430)

        # 结果
        self.resultFlag = False
        self.filePath_1 = ""
        self.filePath_2 = ""
        self.filePath_3 = ""
        self.filePath_4 = ""

    def on_ok(self):
        self.resultFlag = True
        self.on_close()

    def on_cancel(self):
        self.resultFlag = False
        self.on_close()

    def on_close(self):
        self.root.destroy()

    def show(self):
        self.root.mainloop()

    def on_select_1(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.tip.config(text=f"路径:{file_path}")
            self.filePath_1 = file_path

    def on_select_2(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.tip.config(text=f"路径:{file_path}")
            self.filePath_2 = file_path

    def on_select_3(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.tip.config(text=f"路径:{file_path}")
            self.filePath_3 = file_path

    def on_select_4(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.tip.config(text=f"路径:{file_path}")
            self.filePath_4 = file_path


dialog = honmeDialog()
dialog.show()
# revar1 = dialog.resultFlag
# revar2 = dialog.filePath