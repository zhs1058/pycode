import sys, os, threading
import tkinter as tk
from tkinter import filedialog
from tkinter import scrolledtext
from tkinter import  ttk
import requests, json


class honmeDialog:

    def __init__(self, customName):
        self.root = tk.Tk()
        self.root.resizable(0, 0)
        self.root.attributes('-topmost', True)
        self.root.title("邮E连智能外呼")
        self.root.geometry("500x580+500+300")
        # 对象数组
        self.options = [{"name": "随璐", "phone": "18800108853"}, {"name": "闫思洁", "phone": "18810266815"},
                        {"name": "郭青霞", "phone": "13810982832"}, {"name": "柳贞姬", "phone": "13681576706"},
                        {"name": "张美佳", "phone": "18810916089"}, {"name": "邓海廷", "phone": "15901432113"}]
        self.tip = tk.Label(self.root, wraplength="400", fg="red", font=("宋体", 15))
        self.tip.place(x=90, y=18)
        self.msg = tk.Label(self.root, wraplength="400", fg="green", font=("宋体", 15))
        self.msg.place(x=90, y=18)

        # tk.Label(self.root, text="请选择合同文件：", font=("宋体", 15)).place(x=90, y=88)

        # 按钮
        # self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select)
        # self.select_button.place(x=290, y=87)

        # 客户名称
        tk.Label(self.root, text="客户名称", font=("宋体", 15)).place(x=70, y=80)
        self.custom_entry = tk.Entry(self.root, font=("宋体", 15), width=27)
        self.custom_entry.place(x=180, y=80)

        self.custom_entry.insert(tk.END, f"{customName}")

        # 呼叫电话
        tk.Label(self.root, text="呼叫电话", font=("宋体", 15)).place(x=70, y=130)
        self.guhu_entry = tk.Entry(self.root, font=("宋体", 15), width=27)
        self.guhu_entry.place(x=180, y=130)

        # 呼叫日期
        tk.Label(self.root, text="呼叫日期", font=("宋体", 15)).place(x=70, y=180)
        tk.Label(self.root, text="日期格式(20240101)", fg="grey", font=("宋体", 15)).place(x=30, y=200)
        self.guhuhang_entry = tk.Entry(self.root, font=("宋体", 15), width=27)
        self.guhuhang_entry.place(x=180, y=180)

        # 下拉列表
        tk.Label(self.root, text="人员", font=("宋体", 15)).place(x=70, y=250)
        self.user_combo = ttk.Combobox(self.root, values=[item.get("name") for item in self.options], width=19)
        self.user_combo.place(x=180, y=250)
        self.user_combo.bind("<<ComboboxSelected>>", self.on_select_combo)

        # 话术
        tk.Label(self.root, text="话术", font=("宋体", 15)).place(x=70, y=300)
        # self.user_entry = tk.Entry(self.root, font=("宋体", 20), width=19)
        self.speak_entry = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, font=("宋体", 15), width=27, height=7)
        self.speak_entry.place(x=180, y=300)

        # 确认和取消按钮
        self.ok_button = tk.Button(self.root, text="确认", width="10", font=("宋体", 15), command=self.on_ok)
        self.ok_button.place(x=350, y=490)

        # 结果
        self.resultFlag = False
        self.filePath = ""
        self.name = ""
        self.phone = ""

    def on_ok(self):
        # 检查手机号
        if self.guhu_entry.get() == '': 
            self.tip.config(text=f"请填写手机号信息")
            return
        # 检查日期
        if self.guhuhang_entry.get() == '':
            self.tip.config(text=f"请填写日期")
            return
        # 检查日期格式
        if len(self.guhuhang_entry.get()) != 8:
            self.tip.config(text=f"请检查日期格式")
            return
        # 检查话术
        if len(self.speak_entry.get("1.0", tk.END)) == 1:
            self.tip.config(text=f"请选择人员生成话术")
            return

        self.resultFlag = True
        api_url = ("http://21.16.7.46:9087/yyzt/intellect-outbound/outbound-task/detail")
        data = {'templateId': '29491202934b40999e221c84a2ce28de', 'audioNumbers': '0', 'channel': '1'}
        data['targetCustTel'] = self.guhu_entry.get()
        data['targetCustName'] = self.custom_entry.get()
        data['createTellerNo'] = self.phone
        data['createTellerName'] = self.name
        data['firstCallDate'] = self.guhuhang_entry.get()
        header = {'Content-Type': 'application/json'}
        self.send_get_request(api_url, data, header)
        # self.on_close()

    def send_get_request(self, url, params=None, headers=None):
        print(params)
        paramsJSON = json.dumps(params).encode("utf-8")

        try:
            # 发送post请求
            response = requests.post(url, data=paramsJSON, headers=headers)
            if response.status_code == 200:
                print(response.json())
                if response.json().get('code') == '200':
                    self.tip.config(text=f"")
                    self.msg.config(text=f"发送成功！")
                else:
                    message = response.json().get('message')
                    self.tip.config(text=f"发送失败！原因：{message}")
                    self.msg.config(text=f"")

            else:
                message = response.json().get('message')
                self.tip.config(text=f"请求失败！原因：{message}")
                self.msg.config(text=f"")
        except Exception as e:
            print(f"发生异常: {e}")

    def on_cancel(self):
        self.resultFlag = False
        self.on_close()

    def on_close(self):
        self.root.destroy()

    def on_select_combo(self, e):
        self.name = self.user_combo.get()
        custom = self.custom_entry.get()
        for item in self.options:
            if item.get("name") == self.name:
                self.phone = item.get("phone")

        self.speak_entry.delete(1.0, tk.END)
        self.speak_entry.insert(tk.END,
                                f"老师您好，打扰下，{custom}这个客户现在做线上供应链业务，CRM在您的名下，麻烦帮忙维护下信息，谢谢！如有问题请咨询{self.name},电话:{self.phone}")

    def show(self):
        self.root.mainloop()

    def on_select(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.tip.config(text=f"路径:{file_path}")
            self.filePath = file_path


dialog = honmeDialog('我的客户')
dialog.show()
revar1 = dialog.resultFlag
revar2 = dialog.filePath