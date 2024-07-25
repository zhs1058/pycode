import os
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

# 获取当前脚本的所在路径
current_dir = os.path.dirname(__file__)

# 图片的相对路径
image_path = os.path.join(current_dir, "images", "happy.png")

# 创建主窗口
root = tk.Tk()
root.title("宝子喜欢哥子嘛？")

# 显示的文本内容
message = "宝子喜欢哥子嘛？"

# 显示文本
text_label = tk.Label(root, text=message, padx=20, pady=20)
text_label.pack()

# 函数：显示图片
def show_image(image_path):
    image = Image.open(image_path)
    photo = ImageTk.PhotoImage(image)
    image_label = tk.Label(root, image=photo)
    image_label.image = photo  # 保持引用以避免图像被垃圾回收
    image_label.pack()

# 函数：按钮点击事件
def button_click(answer):
    if answer == "喜欢" or answer == "对呀":
        show_image(image_path)
    else:
        messagebox.showinfo("提示", "请回答喜欢或对呀。")

# 创建按钮
button_like = tk.Button(root, text="喜欢", command=lambda: button_click("喜欢"))
button_like.pack(pady=10)
# button_like.place(x=10, y=50)

button_yes = tk.Button(root, text="对呀", command=lambda: button_click("对呀"))
button_yes.pack(pady=10)
# button_yes.place(x=100, y=50)



# 运行主循环
root.mainloop()
