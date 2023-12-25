import win32gui

while True:
    # 窗口标题
    window_title = "82.157.165.106:8082 - 远程桌面连接"

    # 查找窗口句柄
    hwnd = win32gui.FindWindow(None, window_title)

    if hwnd == 0:
        print(f"{window_title} 不存在")
    else:
        print(f"{window_title} 存在，句柄为 {hwnd}")


