import os
import re
import tkinter as tk
from tkinter import RIGHT, BOTH, Text, X, LEFT, RAISED, INSERT, END
from tkinter.ttk import Frame, Button, Label, Entry
import subprocess
from sys import exit
#from win32 import win32api
from win32 import win32process
from win32 import win32process
from win32 import win32gui

TXT_FRAME = ""
LOGIN = ""
PASWORD = ""


def callback(hwnd, pid):
  if win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
    # hide window
    win32gui.ShowWindow(hwnd, 0)


def close_btn():
    exit(0)


def run_btn():
    global TXT_FRAME
    txt = TXT_FRAME
    txt.delete(1.0, END)
    #txt.insert(INSERT, 'Добавление ВПН...\n')
    file = create_vpn(txt)
    txt.insert(INSERT, 'Настройка ВПН...\n')
    res_conf = vpn_config(file, txt)
    if res_conf:
        res_connect = vpn_connect(txt)
        if res_connect:
            txt.insert(INSERT, 'VPN подключение настроено\n')
        else:
            txt.insert(INSERT, 'VPN подключение не удалось настроить\n')
            txt.insert(INSERT, 'Обратитесь в тех.поддержку\n')
            #time.sleep(5)


def get_interface(stdout):
    return re.search(r"([\d]+)\.+([S][a-z]+)", stdout)


def vpn_connect(txt):
    global LOGIN, PASWORD
    #print(self.entry1)
    all_ok = True
    txt.insert(INSERT, "Подключаемся к ВПН...\n")
    completed = subprocess.run(["rasdial.exe", "Summit", LOGIN.get().strip(), PASWORD.get().strip()], capture_output=True)
    lll = list(str(completed.stdout.decode("cp866")).split("\n"))
    for l in lll:
        if l.lower() == "error":
            all_ok = False
        else:
            txt.insert(INSERT, "".join(l))
            txt.insert(INSERT, "\n")
    if completed.returncode == 0:
        route = subprocess.run(["route", "PRINT"], capture_output=True)
        interface = get_interface(str(route.stdout))
        txt.insert(INSERT, f"VPN подключен...\nНастройка маршрутов для сетевого интерфейса: {interface.group(1)}\n")
        #txt.insert(INSERT, f"Настройка маршрута для интерфейса: {interface.group(1)}\n")
        for i in [13, 30, 31, 32, 33, 34, 35]:
            #txt.insert(INSERT, f"route add 192.168.0.0 mask 255.255.0.0 10.200.{i}.1 metric 5 if {interface.group(1)} -p\n")
            res = subprocess.run(["route", f"add 192.168.0.0 mask 255.255.0.0 10.200.{i}.1 metric 5 if {interface.group(1)}"],
                                 capture_output=True)
            if res.returncode == 0:
                r = "Успех"
            else:
                r = "Неудача"
                all_ok = False
            txt.insert(INSERT, f"Настройка маршрута для сети {i}... {r}\n")
        txt.insert(INSERT, "Отключаем VPN...\n")
        subprocess.run(["rasdial.exe", "/DISCONNECT"], capture_output=True)
    return all_ok


def create_vpn(txt):
    # self.txt.insert(INSERT, 'Текстовое поле2')
    # self.txt.insert(INSERT, self.entry1.get())
    cmd = 'Add-VpnConnection ' \
          '-Force ' \
          '-Name "Summit" ' \
          '-TunnelType L2tp ' \
          '-ServerAddress 127.0.0.1 ' \
          '-AuthenticationMethod MSChapv2 ' \
          '-L2tpPsk "SuperPa$$word" ' \
          '-EncryptionLevel Required ' \
          '-RememberCredential'
    txt.insert(INSERT, 'Создание ВПН с именем Summit\n')
    completed = subprocess.run(["powershell.exe", "-Command", cmd], shell=True, start_new_session= True, capture_output=True)
    if completed.returncode != 0:
        if "already" in str(completed.stderr):
            txt.insert(INSERT, 'VPN подключение с именем Summit уже существует\n')
        else:
            txt.insert(INSERT, completed.stderr)
            txt.insert(INSERT, "\n")
    vpn_file = os.environ['APPDATA'] + r'\Microsoft\Network\Connections\Pbk\rasphone.pbk'
    return vpn_file


def open_file(vpn_file):
    with open(vpn_file, 'r', encoding='UTF-8') as file:
        data = file.read()
        return data


def save_file(vpn_file, new_data):
    data = "\n".join(new_data)
    with open(vpn_file, 'w', encoding='UTF-8') as file:
        file.write(data)


def str_gen(data):
    for i in data.split("\n"):
        yield i


def vpn_config(vpn_file, txt):
    find_vpn = False
    if os.path.exists(vpn_file):
        file = open_file(vpn_file)
        file_data = str_gen(file)
        loop = True
        new_data = []
        while loop:
            try:
                data = next(file_data)
                new_data.append(data)
            except StopIteration:
                txt.insert(INSERT, "Файл полностью прочитан\n")
                loop = False
                break
            if re.match(r'^\[Summit\]$', data):
                txt.insert(INSERT, "Найден нужный ВПН\n")
                txt.insert(INSERT, "читаем параметры....\n")
                find_vpn = True
                while loop:
                    try:
                        argument = next(file_data)
                    except StopIteration:
                        txt.insert(INSERT, "Файл полностью прочитан\n")
                        loop = False
                        break
                    if re.match(r'^\[.*\]$', argument):
                        txt.insert(INSERT, f"Найдено другой впн {argument}\n")
                        new_data.append(argument)
                        loop = False
                    else:
                        argument = argument.replace("PreferredHwFlow=0", "PreferredHwFlow=1")
                        argument = argument.replace("PreferredProtocol=0", "PreferredProtocol=1")
                        argument = argument.replace("PreferredCompression=0", "PreferredCompression=1")
                        argument = argument.replace("PreferredSpeaker=0", "PreferredSpeaker=1")
                        argument = argument.replace("IpPrioritizeRemote=1", "IpPrioritizeRemote=0")
                        argument = argument.replace("IpInterfaceMetric=0", "IpInterfaceMetric=5")
                        new_data.append(argument)
        txt.insert(INSERT, "Записываем файл.....\n")
        save_file(vpn_file, new_data)
    else:
        txt.insert(INSERT, "Файл VPN не найден.\n")
    return find_vpn

def main():
    global LOGIN, PASWORD, TXT_FRAME
    window = tk.Tk()
    window.geometry("500x500+500+400")

    window.title("Настройка ВПН")
    #window.pack(fill=BOTH, expand=True)

    frame1 = tk.Frame()
    frame1.pack(fill=X)

    lbl1 = Label(frame1, text="Введите логин для ВПН:", width=23)
    lbl1.pack(side=LEFT, padx=5, pady=5)

    entry1 = Entry(frame1)
    entry1.pack(fill=X, padx=5, expand=True)

    frame2 = Frame()
    frame2.pack(fill=X)

    lbl2 = Label(frame2, text="Введите пароль для ВПН:", width=23)
    lbl2.pack(side=LEFT, padx=5, pady=5)

    entry2 = Entry(frame2)
    entry2.pack(fill=X, padx=5, expand=True)

    frame3 = Frame()
    frame3.pack(fill=BOTH, expand=True)

    txt = Text(frame3)
    txt.pack(fill=BOTH, pady=5, padx=5, expand=True)
    txt.insert(INSERT, "Внимание! Для корректной работы необходимы административные права.\n")

    frame = Frame(relief=RAISED, borderwidth=1)
    frame.pack(fill=BOTH, expand=True)

    TXT_FRAME = txt
    LOGIN = entry1
    PASWORD = entry2

    closeButton = Button(text="Закрыть", command=close_btn)
    closeButton.pack(side=RIGHT, padx=5, pady=5)
    okButton = Button(text="Выполнить", command=run_btn)
    okButton.pack(side=RIGHT)
    window.mainloop()


if __name__ == '__main__':
    win32gui.EnumWindows(callback, os.getppid())
    main()
