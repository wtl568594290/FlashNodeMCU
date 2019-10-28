import os
import subprocess
import time
import winsound
from tkinter import StringVar, Label, Tk, Button, filedialog, messagebox
from tkinter.ttk import Combobox

import serial


class ConfigWindow:
    def __init__(self, master):
        self.master = master
        self.initWidgets()

    def initWidgets(self):
        self.master.title('FlashNodeMCU Config')
        self.buttonPort = Button(self.master, text='刷新串口', command=self.refreshPort)
        self.buttonPort.grid(row=0, column=0)
        self.portVar = StringVar()
        self.cb = Combobox(self.master, state='readonly', textvariable=self.portVar)
        self.cb['state'] = 'readonly'
        self.cb.grid(row=0, column=1)
        self.buttonBin = Button(self.master, text='选择bin文件', command=self.selectBin)
        self.buttonBin.grid(row=1, column=0, pady=2)
        self.binVar = StringVar()
        self.labelBin = Label(self.master, textvariable=self.binVar)
        self.labelBin.grid(row=1, column=1)
        self.buttonLua = Button(self.master, text='选择init.lua', command=self.selectLua)
        self.buttonLua.grid(row=2, column=0, pady=2)
        self.luaVar = StringVar()
        self.labelLua = Label(self.master, textvariable=self.luaVar)
        self.labelLua.grid(row=2, column=1)
        self.buttonMac = Button(self.master, text='选择MAC文件', command=self.selectMac)
        self.buttonMac.grid(row=3, column=0, pady=2)
        self.macVar = StringVar()
        self.labelMac = Label(self.master, textvariable=self.macVar)
        self.labelMac.grid(row=3, column=1)
        self.buttonSure = Button(self.master, text='开始烧录', command=self.startFlash)
        self.buttonSure.grid(row=4, columnspan=2, pady=10)
        self.refreshPort()
        self.master.mainloop()

    def startFlash(self):
        port = self.portVar.get()
        bin = self.binVar.get()
        lua = self.luaVar.get()
        mac = self.macVar.get()
        if port == '' or port is None:
            return
        if bin == '' or bin is None:
            return
        if lua == '' or lua is None:
            return
        if mac == '' or mac is None:
            return
        self.master.destroy()
        Flash(port, bin, lua, mac)

    def refreshPort(self):
        p = subprocess.Popen(
            'nodemcu-tool devices', shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            universal_newlines=True)
        ports = []
        while p.poll() is None:
            line = p.stdout.readline()
            if 'COM' in line:
                line = line.split(' ')
                for com in line:
                    if 'COM' in com:
                        ports.append(com)
                        break
        self.cb['values'] = ports
        if len(ports) > 0:
            self.cb.current(0)
        else:
            messagebox.showerror(title='错误', message='没有已经连接的nodemcu设备')

    def selectBin(self):
        bin = filedialog.askopenfile(title='选择bin文件', filetypes=[('固件', '.bin')])
        if bin is not None:
            self.binVar.set('"' + bin.name + '"')

    def selectLua(self):
        lua = filedialog.askopenfile(title='选择init.lua文件', filetypes=[('Lua文件', 'init.lua')])
        if lua is not None:
            self.luaVar.set('"' + lua.name + '"')

    def selectMac(self):
        mac = filedialog.askopenfile(title='选择存放Mac的文件', filetypes=[('文本文件', '.txt')])
        if mac is not None:
            self.macVar.set(mac.name)


class Flash:
    def __init__(self, port, bin, lua, macLocation):
        self.port = port
        self.bin = bin
        self.lua = lua
        self.macLocation = macLocation
        self.mac = ''
        self.startFlash()

    def startFlash(self):
        while True:
            os.system('cls')
            self.flashNodeMCU()
            self.verifyPort()

    def verifyPort(self):
        portExist = True
        ser = serial.Serial()
        ser.port = self.port
        ser.baudrate = '115200'
        self.printInfo('请拔出NodeMCU')
        while True:
            try:
                ser.open()
                if ser.is_open is True:
                    ser.close()
                    if not portExist:
                        self.printInfo('NodeMCU已插入，即将进行下一个')
                        time.sleep(1)
                        break
                    else:
                        try:
                            winsound.PlaySound('C:/Windows/Media/tada.wav', winsound.SND_FILENAME)
                        except Exception:
                            pass
            except serial.SerialException:
                if portExist:
                    portExist = False
                    self.printInfo('NodeMCU已拔出，请插入NodeMCU进行下一个,退出请按Ctrl+C')
            time.sleep(1)

    def flashNodeMCU(self, runStr='welcome'):
        self.printTitle('开始烧录')
        p = subprocess.Popen(
            'esptool.py -p ' + self.port + ' -b 1500000 -a hard_reset write_flash -e -ff 26m -fm dio -fs 4MB 0x00000 ' + self.bin,
            shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, universal_newlines=True)
        isFlashFinish = False
        while p.poll() is None:
            line = p.stdout.readline()
            if line != '' and line != '\n':
                if 'Writing' in line:
                    if '(100 %)' not in line:
                        print('\r' + line.replace('\n', ''), end='')
                    else:
                        print('\r' + line, end='')
                else:
                    print(line, end='')
                if 'MAC' in line:
                    self.mac = line
                if 'verified' in line:
                    isFlashFinish = True
        if not isFlashFinish:
            messagebox.showerror(title='错误', message='烧录出错')
            return
        self.printTitle('开始初始化文件系统')
        try:
            ser = serial.Serial(self.port, 115200, timeout=1)
            bSer = True
            while bSer:
                line = ser.readline()
                if line != b'':
                    try:
                        print(line.decode())
                    except Exception:
                        print('', end='')
                    if b'>' in line:
                        bSer = False
            ser.close()
        except serial.SerialException:
            messagebox.showerror(title='错误', message='打开串口等待Format出错')
        time.sleep(1)
        self.printTitle('开始上传')
        p = subprocess.Popen(
            'nodemcu-tool upload -p ' + self.port + ' ' + self.lua, shell=True, stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=True)
        isUploadFinish = False
        while p.poll() is None:
            line = p.stdout.readline()
            if line != '' and line != '\n':
                print(line, end='')
                if 'complete' in line:
                    isUploadFinish = True
        if not isUploadFinish:
            messagebox.showerror(title='错误', message='上传出错')
            return
        self.printTitle('开始运行')
        p = subprocess.Popen(
            'nodemcu-tool run -p ' + self.port + ' init.lua', shell=True, stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            universal_newlines=True)
        isRunSuccess = False
        while p.poll() is None:
            line = p.stdout.readline()
            if line != '' and line != '\n':
                if runStr in line:
                    tplt = '\n{0:{1}^80}\n'
                    print(tplt.format(line.replace('\n', ''), ' '))
                    isRunSuccess = True
                else:
                    print(line, end='')
        if isRunSuccess:
            mac = self.mac.split(':', 1)[1].strip().replace(':', '-').upper()
            self.printTitle('开始写入MAC地址到' + self.macLocation)
            try:
                with open(self.macLocation, 'a') as f:
                    f.write(mac + '\n')
                self.printInfo('写入MAC:' + mac + '成功')
            except Exception:
                messagebox.showerror(title='错误', message='mac地址写入出错')
                return
            self.printInfo('烧录成功')
        else:
            messagebox.showerror(title='错误', message='运行出错了！！！！')

    def printTitle(self, str):
        print('\n' + '*' * 80)
        length = len(str.encode('gbk'))
        if length <= 70:
            length = 70 - (length - len(str))
            tplt = '%s{0:{1}^%d}%s' % ('*' * 5, length, '*' * 5,)
            print(tplt.format(str, ' '))
        else:
            print(str)
        print('*' * 80 + '\n')

    def printInfo(self, str):
        tplt = '{0:{1}^%d}\n' % (80 - (len(str.encode('gbk')) - len(str)))
        print(tplt.format(str, '-'))


if __name__ == '__main__':
    ConfigWindow(Tk())
