import sys
import openpyxl
import tkinter as tk
from tkinter import filedialog


class honmeDialog:

    def __init__(self):
        self.root = tk.Tk()
        self.root.resizable(0, 0)
        self.root.attributes('-topmost', True)
        self.root.title("选择文件")
        self.root.geometry("600x500+500+100")

        self.tip = tk.Label(self.root, wraplength="400", text="", fg="red", font=("宋体", 15))
        self.tip.place(x=50, y=10)
        self.msg = tk.Label(self.root, wraplength="400", text="", fg="green", font=("宋体", 15))
        self.msg.place(x=90, y=30)
        # 选择证书统计表
        tk.Label(self.root, text="请选择台账", font=("宋体", 15)).place(x=90, y=188)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select_1)
        self.select_button.place(x=330, y=187)

        # 生成文件路径
        # tk.Label(self.root, text="生成文件路径：", font=("宋体", 15)).place(x=90, y=260)
        # self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15),
        #                                command=self.on_select_dir)
        # self.select_button.place(x=330, y=260)

        # 确认按钮
        self.ok_button = tk.Button(self.root, text="确认", width="10", font=("宋体", 15), command=self.on_ok)
        self.ok_button.place(x=330, y=350)

        # 结果
        self.filePath_1 = ""
        self.fileDir = ""

        self.tempResultMap = {}
        self.result = []
        self.titleIndexMap = {9: "经济金融类", 10: "国际贸易融资类", 11: "财务审计类",
                              12: "风险管理类", 13: "法律类", 14: "计算机类",
                              15: "人力资源类", 16: "其他", 17: "职称类证书"}

    def on_ok(self):
        if (self.filePath_1 == ""):
            self.tip.config(text=f"请选择选择证书统计表")
            return

        self.fill_data()
        self.on_close()

    def on_cancel(self):
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



    # -------------------------------------------------------------

    def fill_data(self):
        # 输出文件路径（新的Excel表格）
        print(self.filePath_1)




dialog = honmeDialog()
dialog.show()
