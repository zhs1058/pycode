import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from datetime import datetime

def on_select():
    if calendar_frame.winfo_ismapped():
        calendar_frame.pack_forget()
    else:
        calendar_frame.pack(pady=20)

def on_date_selected():
    selected_date = cal.get_date()
    formatted_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y%m%d")
    result_label.config(text=f"选择的日期: {formatted_date}")
    calendar_frame.pack_forget()

# 创建主窗口
root = tk.Tk()
root.title("日期选择器")

# 创建下拉框
combo_var = tk.StringVar()
combo = ttk.Combobox(root, textvariable=combo_var, values=["年", "月", "日"])
combo.set("年")
combo.pack()

# 创建日历框架
calendar_frame = tk.Frame(root)

# 创建日历
cal = Calendar(calendar_frame, selectmode="day", year=2024, month=1, day=1)
cal.pack()

# 创建按钮，点击后显示/隐藏日历框架
select_button = ttk.Button(root, text="选择日期", command=on_select)
select_button.pack(pady=10)

# 创建确认按钮，点击后获取选择的日期并隐藏日历框架
confirm_button = ttk.Button(calendar_frame, text="确认", command=on_date_selected)
confirm_button.pack(pady=10)

# 显示选择的日期结果
result_label = ttk.Label(root, text="")
result_label.pack(pady=10)

# 运行主循环
root.mainloop()
