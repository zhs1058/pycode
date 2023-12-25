import tkinter as tk
from tkinter import messagebox
import threading


class TimerLabel(tk.Label):
    def __init__(self, master=None, **kwargs):
        tk.Label.__init__(self, master, **kwargs)
        self.seconds_left = 0
        self.timer = None

    def start(self, seconds):
        self.seconds_left = seconds
        self.countdown()

    def stop(self):
        if self.timer:
            self.after_cancel(self.timer)
            self.timer = None

    def countdown(self):
        if self.seconds_left <= 0:
            self.config(text="")
        else:
            self.config(text=str(self.seconds_left))
            self.seconds_left -= 1
            self.timer = self.after(1000, self.countdown)


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("验证码倒计时")

        self.timer_label = TimerLabel(self.root, font=("Arial", 20))
        self.timer_label.pack(pady=20)

        self.send_button = tk.Button(self.root, text="发送验证码", command=self.send)
        self.send_button.pack()

        self.root.mainloop()

    def send(self):
        self.send_button.config(state=tk.DISABLED)
        threading.Thread(target=self.send_code).start()

    def send_code(self):
        # 这里写发送验证码的代码
        # 发送成功后启动计时器
        self.timer_label.start(60)
        messagebox.showinfo("提示", "验证码已发送，请注意查收！")
        self.send_button.config(state=tk.NORMAL)


if __name__ == "__main__":
    app = App()
