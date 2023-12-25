from win32gui import *


while True:
    title = set()

    def foo(hwnd, mouse):
        if IsWindow(hwnd) and IsWindowEnabled(hwnd) and IsWindowVisible(hwnd):
            if GetWindowText(hwnd).find('82.157.165.106:8082') != -1:
                title.add(hwnd)


    EnumWindows(foo, 0)
    if title.pop() != 0:
        print('窗口存在')
    else:
        print('窗口不存在')


