from __future__ import print_function
from os import times
import tkinter as tk
import tkinter
from tkinter import filedialog
from tkinter import messagebox
import base64
import json
import ctypes, sys
from typing_extensions import final
import requests
import time
import os
import threading
import csv
import win32com.client

class windows:
    def __init__(self):
        #self.getAdmin()
        self.root = tk.Tk()
        self.root.title('ACCUNIQ')
        self.root.geometry('500x300')
        self.file_path = ''
        self.executepath = os.path.expanduser("~") + "\\Desktop\\ACCUNIQ"
        self.repone ={}
        self.DataList = []
        self.getData()
        self.stop = False

        self.text1 = tk.Label(self.root,text='ACCUNIQ',font=('Arial', 18))
        self.text1.place(relx=0.35,rely=0.1)

        self.text2 = tk.Label(self.root,text='預設路徑:',font=('Arial', 18))
        self.text2.place(relx=0.1,rely=0.3)
        
        # 建立排程檢查
        try:
            if self.data['task'] != True:
                #self.getAdmin()
                print(os.getcwd())
                #os.system("schtasks /create /tn ACCUNIQControl /tr "+ os.getcwd() +"\\ACCUNIQ(24101).exe /sc ONSTART /ru System")
                self.createShortCut(os.path.expanduser("~") + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup",os.getcwd() + "\\ACCUNIQ_24101.exe")
        except:
            self.data['task'] = True
            self.setData()
            self.getData()
            #self.getAdmin()
            #os.system("schtasks /create /tn ACCUNIQControl /tr "+ os.getcwd() +"\\ACCUNIQ(24101).exe /sc ONSTART /ru System")
            self.createShortCut(os.path.expanduser("~") + "\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup",os.getcwd() + "\\ACCUNIQ_24101.exe")
            
        # 建立路徑檢查
        try:
            if self.data['path'] != '':
                self.text3 = tk.Label(self.root,text=self.data['path'],font=('Arial', 12))
                self.file_path = self.data['path']
            else:
                raise TypeError("Error")
        except:
            try:
                self.file_path = os.path.expanduser("~") + "\\Desktop\\accuniqApi"
            except:
                pass
            self.data['path'] = self.file_path
            self.setData()
            self.getData()
        
            self.text3 = tk.Label(self.root,text=self.data['path'],font=('Arial', 12))
        
        self.text3.place(relx=0.4,rely=0.3)

        self.text4 = tk.Label(self.root,text='',font=('Arial', 18))
        self.text4.place(relx=0.4,rely=0.5)

        self.okBtn = tk.Button(self.root,font=('Arial', 18),text='開始監控',command=self.btnClick)
        self.okBtn.place(relx=0.1,rely=0.8)

        self.changeBtn = tk.Button(self.root,font=('Arial', 18),text='重選目錄',command=self.btn2Click)
        self.changeBtn.place(relx=0.4,rely=0.8)

        self.chenalBtn = tk.Button(self.root,font=('Arial', 18),text='停止監控',command=self.btn3Click)
        self.chenalBtn.place(relx=0.7,rely=0.8)
        self.chenalBtn.configure(state='disable')

        self.root.protocol('WM_DELETE_WINDOW',self.stopApp)

        self.root.tk.call('package', 'require', 'Winico')
        icon = self.root.tk.call('winico', 'createfrom', self.executepath + '\\avvji-bcalw-001.ico')  # 这里说一下xxx.ico 就是你要显示在托盘上的小图标
        self.root.tk.call('winico', 'taskbar', 'add', icon,
            '-callback', (self.root.register(self.menu_func), '%m', '%x', '%y'),
            '-pos', 0,
            '-text', u'accuniq發送器')
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label=u'顯示畫面', command=self.root.deiconify)
        self.menu.add_command(label=u'退出', command=self.stopApp)
        self.root.withdraw()
        self.btnClick()
        self.root.mainloop()
    
    # 函數定義區
    def menu_func(self,event, x, y):
        if event == 'WM_RBUTTONDOWN':    # 监听右击事件
            self.menu.tk_popup(x, y) #弹出菜单
        if event == 'WM_LBUTTONDOWN':   #左 事件  还有其他的如 WM_LBUTTONDBLCLK 左双击
            self.root.deiconify() #显示主页面

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
            os._exit(0)
    
    def setData(self):
        try:
            with open(self.executepath + '\\temp.json','w') as f:
                f.write(json.dumps(self.data))
        except:
            messagebox.showerror("錯誤提示","請確認任您的軟體資料夾路徑為\n" + self.executepath + "\n，詳情請見說明文件!!!")
            os._exit(0)
            

    def getData(self):
        try:
            with open(self.executepath + '\\temp.json','r') as f:
                self.data = json.load(f)
        except:
            self.data = {}
    
    def sendToAPI(self,file,fileType,arrayData=''):
        send_data ={
            'Token':'dad5ece51af754e71e3333fc909b6b3b32893e826a29960d335d1378704237db',
            'Type':fileType,
            'Image':self.imgToBase64(file),
            'GymCode+RandomNo':arrayData[0],
            'ID':arrayData[1],
            'MeasureDate':arrayData[2],
            'MeasureTime':arrayData[3],
            'Height':arrayData[4],
            'Weight':arrayData[5],
            'Age':arrayData[6],
            'Gender':arrayData[7],
            'BodyMoisture':arrayData[17],
            'Fat':arrayData[19],
            'Protein':arrayData[20],
            'BFM':arrayData[26],
            'SMM':arrayData[31],
            'BMI':arrayData[37],
            'FatPercentage':arrayData[41],
            'WHR':arrayData[45],
            'VisceralLevel':arrayData[49],
            'VisceralArea':arrayData[51],
            'LeftArm':arrayData[55],
            'RightArm':arrayData[58],
            'LeftLeg':arrayData[61],
            'RightLeg':arrayData[64],
            'Body':arrayData[67],
            'FatLeftArmWeight':arrayData[70],
            'FatRightArmWeight':arrayData[73],
            'FatLeftLegWeight':arrayData[76],
            'FatRightLegWeight':arrayData[79],
            'FatBodyWeight':arrayData[82],
            'BodyType':arrayData[85],
            'PhysicalAge':arrayData[86],
            'BMR':arrayData[87],
            'TDEE':arrayData[88],
            'VisceralWeight':arrayData[90],
            'DegreeOfObesity':arrayData[91],
            'AC':arrayData[92],
            'FitnessScore':arrayData[93],
            'EdemaIndex':arrayData[96],
            'WeightControl':arrayData[97],
            'WeightControl':arrayData[98],
            'WeightControl':arrayData[99],
            'EdemaIndex':arrayData[100],
            'MuscleLeftArmWeight':arrayData[105],
            'MuscleRightArmWeight':arrayData[108],
            'MuscleLeftLegWeight':arrayData[111],
            'MuscleRightLegWeight':arrayData[114],
            'MuscleBodyWieght':arrayData[117],
            }
        print(send_data)
        r = requests.post('https://accuniq.eternalwtech.com/api/AccuniqRecord',data=send_data)
        self.repone = r.json()
        print(self.repone)
        try:
            self.outputFile(self.repone['data']['img'])
        except:
            self.text4.configure(text='<<<上傳失敗>>>')
        time.sleep(1)
    
    def imgToBase64(self,file_path):
        with open(file_path,'rb') as f:
            image = f.read()
            image_base64 = base64.encodebytes(image)
            return image_base64

    def btnClick(self):
        self.okBtn.configure(state='disable')
        self.changeBtn.configure(state='disable')
        self.chenalBtn.configure(state='normal')
        t = threading.Thread(target=self.startSend)
        t.start()
    
    def outputFile(self,data):
        with open(self.executepath + '\\output.txt','a') as f:
            f.write(data+'\n')
    
    def btn2Click(self):
        self.file_path = filedialog.askdirectory()
        self.data['path'] = self.file_path
        self.setData()
        self.getData()
        self.text3.configure(text=self.data['path'])

    def startSend(self):
        while True:
            time.sleep(1)
            try:
                fileList = os.listdir(self.file_path)
            except FileNotFoundError:
                os.makedirs(os.path.expanduser("~") + "\\Desktop\\accuniqApi")
                continue
            csvFileList = []
            ImageFileList = []
            for file in fileList:
                self.text4.configure(text='檢測到新檔，辨識中...')
                self.text4.place(relx=0.25,rely=0.5)
                if (file.find('jpg') != -1 or file.find('png') != -1):
                    ImageFileList.append(self.file_path + '\\' + file)
                elif (file.find('csv') != -1):
                    csvFileList.append(self.file_path + '\\' +file)
            for idx in range(len(ImageFileList)):
                QFile = ImageFileList[idx].split(".")[0].replace("_1","") + ".csv"
                if (QFile) in csvFileList :
                    self.text4.configure(text='辨識已完成，傳送中...')
                    self.text4.place(relx=0.25,rely=0.5)
                    self.getCSV(QFile)
                    try:
                        self.sendToAPI(ImageFileList[idx],ImageFileList[idx].split('.')[1],self.DataList)
                        os.remove(QFile)
                        os.remove(ImageFileList[idx])
                    except requests.exceptions.ConnectionError:
                        messagebox.showerror("網路錯誤", "請檢查您的網路連線是否正常!!!")
                        continue
            csvFileList = None
            ImageFileList = None
            self.text4.configure(text='~系統就緒~')
            self.text4.place(relx=0.4,rely=0.5)
            if self.stop == True:
                break
        self.stop = False
    
    def btn3Click(self):
        self.stop = True
        self.okBtn.configure(state='normal')
        self.changeBtn.configure(state='normal')
        self.chenalBtn.configure(state='disable')
    
    def stopApp(self):
        ans = messagebox.askyesno("結束確認","請問是否結束應用?\n是:立即停止並退出\n否:繼續執行並最小化")
        if ans == True:
            os._exit(0)
        else:
            self.root.withdraw()
    
    def getCSV(self,csvname):
        with open(csvname,'r+',encoding='big5') as csvFile:
            csvData = csv.reader(csvFile)
            self.DataList = list(csvData)[0]
            if(self.DataList[7] == '男'):
                self.DataList[7] = 1
            elif(self.DataList[7] == '女'):
                self.DataList[7] = 2
            #print(self.DataList)

    def createShortCut(self,ShortCutpath,targetpath):
        execute = win32com.client.Dispatch("wscript.shell")
        shortcut = execute.CreateShortCut(ShortCutpath + "\\ACCUNIQ_24101.lnk")
        shortcut.Targetpath= targetpath
        shortcut.save()


if __name__ == '__main__':
    f = windows()
