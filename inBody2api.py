from __future__ import print_function
from os import times
import tkinter as tk
from tkinter import filedialog
import base64
import json
import ctypes, sys
from typing_extensions import final
import requests
import time
import os
import threading

class windows:
    def __init__(self):
        self.getAdmin()
        self.root = tk.Tk()
        self.root.title('圖片傳送器')
        self.root.geometry('500x300')
        self.file_path = ''
        self.repone ={}
        self.getData()

        self.text1 = tk.Label(self.root,text='圖片傳送器',font=('Arial', 18))
        self.text1.place(relx=0.4,rely=0.1)

        self.text2 = tk.Label(self.root,text='預設路徑:',font=('Arial', 18))
        self.text2.place(relx=0.1,rely=0.3)

        try:
            if self.data['path'] != '':
                self.text3 = tk.Label(self.root,text=self.data['path'],font=('Arial', 12))
                self.file_path = self.data['path']
            else:
                raise TypeError("Error")
        except:
            self.file_path = filedialog.askdirectory()
            self.data['path'] = self.file_path
            self.setData()
            self.getData()
            self.text3 = tk.Label(self.root,text=self.data['path'],font=('Arial', 12))
        self.text3.place(relx=0.4,rely=0.3)

        self.text4 = tk.Label(self.root,text='',font=('Arial', 18))
        self.text4.place(relx=0.4,rely=0.5)

        self.okBtn = tk.Button(self.root,font=('Arial', 18),text='開始傳送',command=self.btnClick)
        self.okBtn.place(relx=0.3,rely=0.8)

        self.changeBtn = tk.Button(self.root,font=('Arial', 18),text='重選目錄',command=self.btn2Click)
        self.changeBtn.place(relx=0.6,rely=0.8)

        self.root.mainloop()
    
    # 函數定義區
    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def getAdmin(self):
        if self.is_admin():
           pass
        else:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    
    def setData(self):
        try:
            with open('temp.json','w') as f:
                f.write(json.dumps(self.data))
        except:
            self.getAdmin()
            self.setData()

    def getData(self):
        try:
            with open('temp.json','r') as f:
                self.data = json.load(f)
        except:
            self.data = {}
    
    def sendToAPI(self,file):
        print(1)
        send_data ={'token':'dad5ece51af754e71e3333fc909b6b3b32893e826a29960d335d1378704237db','type':'jpg','image':self.imgToBase64(file)}
        print(2)
        r = requests.post('https://v3.eternalwtech.com/api/upload-image',data=send_data)
        self.repone = r.json()
        self.outputFile(self.repone['data'])
        time.sleep(1)
    
    def imgToBase64(self,file_path):
        with open(file_path,'rb') as f:
            image = f.read()
            image_base64 = base64.encodebytes(image)
            return image_base64

    def btnClick(self):
        self.okBtn.configure(state='disable')
        self.changeBtn.configure(state='disable')
        t = threading.Thread(target=self.startSend)
        t.start()
    
    def outputFile(self,data):
        with open('output.txt','a') as f:
            f.write(data+'\n')
    
    def btn2Click(self):
        self.file_path = filedialog.askdirectory()
        self.data['path'] = self.file_path
        self.setData()
        self.getData()
        self.text3.configure(text=self.data['path'])

    def startSend(self):
        fileList = os.listdir(self.file_path)
        i = 1
        for file in fileList:
            self.text4.configure(text='第'+str(i)+'張傳送中...')
            finalPath = self.file_path+'\\'+file
            self.sendToAPI(finalPath)
            os.remove(finalPath)
            i+=1
        self.text4.configure(text='~傳送完成~')
        self.okBtn.configure(state='normal')
        self.changeBtn.configure(state='normal')


if __name__ == '__main__':
    f = windows()
