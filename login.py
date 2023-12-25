import threading
import time
import tkinter as tk
from tkinter import *
# 导入国密算法sm4包
from gmssl import sm4, sm3, func
import requests, json, datetime


class LoginDialog:

    def __init__(self, title, keyWord):
        self.root = tk.Tk()
        self.root.resizable(0, 0)
        self.keyWord = keyWord
        self.butFlag = False
        #self.root.iconbitmap('D:\download\psbc1.ico')
        self.root.title(title)
        self.root.geometry("500x230+500+300")
        self.root.tips = StringVar()
        self.root.butText = StringVar()
        self.root.butText.set("获取验证码")
        self.timerCount = threading.Timer(1.0, self.countdown)
        self.timerTips = threading.Timer(0.5, lambda: self.root.tips.set(''))

        self.root.message = tk.Message(self.root, width="300", textvariable=self.root.tips, fg="red", font=("宋体", 15))
        self.root.message.place(x=190, y=20)
        # 柜员号输入框
        tk.Label(self.root, text="柜员号：", font=("宋体", 20)).place(x=70, y=50)
        self.user_entry = tk.Entry(self.root, font=("宋体", 20), width=17)
        self.user_entry.place(x=190, y=55)

        # 验证码输入框和获取验证码按钮
        tk.Label(self.root, text="验证码：", font=("宋体", 20)).place(x=70, y=90)
        self.code_entry = tk.Entry(self.root, font=("宋体", 20), width=8)
        self.code_entry.place(x=190, y=95)
        self.code_button = tk.Button(self.root, width="10", font=("宋体", 15), textvariable=self.root.butText, command=self.get_code)
        self.code_button.place(x=325, y=93)
        #
        # # 登录和取消按钮
        self.ok_button = tk.Button(self.root, text="登录", width="10", font=("宋体", 15), command=self.on_ok)
        self.ok_button.place(x=190, y=150)
        self.cancel_button = tk.Button(self.root, text="取消", width="10", font=("宋体", 15), command=self.on_cancel)
        self.cancel_button.place(x=325, y=150)

        # 登录结果
        self.result = None
        #是否成功获取登录结果标识
        self.resultFlag = False

    def showTips(self, msg):
        self.root.tips.set(msg)
        self.timerTips = threading.Timer(3.0, lambda : self.root.tips.set(''))
        self.timerTips.start()
    def get_code(self):
        if self.butFlag:
            self.showTips('请一分钟后再试！')
            return
        if self.user_entry.get() == '':
            self.showTips('请输入柜员号！')
            return
        if len(self.user_entry.get()) != 11:
            self.showTips('请输入正确的柜员号！')
            return
        requestData = {'transCode': 100000, 'timestamp': self.getTime(), 'counterId': self.user_entry.get()}
        responseResult = self.sendMsg(requestData)
        if responseResult == 'error':

            self.showTips('参数可能被篡改，请联系开发人员！')
        elif responseResult['code'] == 0:
            self.showTips('验证码已经发送')
            self.butFlag = True
            self.timerCount = threading.Timer(1.0, self.countdown)
            self.timerCount.start()


        else:
            self.showTips(responseResult['message'])

    def countdown(self):

        t = 60
        while t > 0:
            self.root.butText.set(t)
            t -= 1
            time.sleep(1)

        self.root.butText.set('获取验证码')
        self.butFlag = False

    def on_ok(self):
        requestData = {'transCode': 200000, 'timestamp': self.getTime(), 'counterId': self.user_entry.get(),  'verifyCode': self.code_entry.get()}
        responseResult = self.sendMsg(requestData)
        if responseResult['code'] != 0:
            self.showTips(responseResult['message'])
            return
        else:
            self.result = responseResult['data']
            if self.timerTips.is_alive():
                self.timerTips.cancel()
            if self.timerCount.is_alive():
                self.timerCount.cancel()
                self.timerCount.join(timeout=1)
            self.root.destroy()
            self.resultFlag = True


    def on_cancel(self):
        self.result = "取消登录"  # 取消登录时设置结果为字符串"取消登录"
        if self.timerTips.is_alive():
            self.timerTips.cancel()
        if self.timerCount.is_alive():
            self.timerCount.cancel()
            self.timerCount.join(timeout=1)
        self.root.destroy()

    def show(self):
        self.root.mainloop()

    def sm3_encode(self, data):
        return sm3.sm3_hash(func.bytes_to_list(data.encode('utf-8')))

    # 国密sm4加密
    def sm4_encode(self, data):
        sm4Alg = sm4.CryptSM4()  # 实例化sm4
        sm4Alg.set_key(self.keyWord.encode(), sm4.SM4_ENCRYPT)  # 设置密钥
        dateStr = str(data)
        enRes = sm4Alg.crypt_ecb(dateStr.encode())  # 开始加密,bytes类型，ecb模式
        enHexStr = enRes.hex()
        return enHexStr  # 返回十六进制值

    # 国密sm4解密
    def sm4_decode(self, data):
        sm4Alg = sm4.CryptSM4()  # 实例化sm4
        sm4Alg.set_key(self.keyWord.encode(), sm4.SM4_DECRYPT)  # 设置密钥
        deRes = sm4Alg.crypt_ecb(bytes.fromhex(data))  # 开始解密。十六进制类型,ecb模式
        deHexStr = deRes.decode()
        return deHexStr

    def sendMsg(self, data):
        dataSort = {}
        result = {}
        for key, value in sorted(data.items(), key=lambda x: x[0]):
            dataSort[key] = value

        dataStr = json.dumps(dataSort)

        result['data'] = self.sm4_encode(data=dataStr)
        result['sign'] = self.sm3_encode(data='psbcWind' + dataStr)
        return self.doRequest(data=result)

    def doRequest(self, data):
        url = 'https://www.moonstorm.top:8443/wind/windService'
        headers = {
            'PFPJ-GatewayCode': 'corpwx',
            'PFPJ-SourceSysKey': 'e8103990e255450ca1f64638296831a3',
            'Content-type': 'application/json'
        }
        response = requests.post(url=url, data=json.dumps(data), headers=headers)
        # 解密
        decStr = self.sm4_decode(response.json()['data'])
        mySing = self.sm3_encode('psbcWind' + decStr)
        if (mySing == response.json()['sign']):
            return json.loads(decStr)
        else:
            return 'error'

    def getTime(self):
        now = datetime.datetime.now()
        return now.strftime("%Y-%m-%d %H:%M:%S")


keyword = "psbcbjqywxcommon"
dialog = LoginDialog("登录云桌面", keyword)
dialog.show()
print(dialog.result)

