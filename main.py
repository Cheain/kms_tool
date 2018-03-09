from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext
import threading
import json
import os
import winreg
import subprocess
import time

import server

label_font = ("Times New Roman", 12, "normal")
content_font = ("Times New Roman", 14, "normal")


class My_frame():
    def __init__(self):
        self._win_vbs = '%windir%\system32\slmgr.vbs'
        self.keypaths = [r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                         r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"]

        self.root = Tk()
        self.root.title('KMS激活')
        # self.root.geometry('600x800')
        self.root.resizable(False, False)

        self._get_gvlks()
        self.show_activate_frame()
        self.show_content_frame()

        self.office_location_iterator = self._search_office_vbs()

    def _test(self):
        pass

    def _get_gvlks(self):
        with open('GVLK.json', 'r') as f:
            g = json.load(f)
        self.gvlks = {'******  选择Windows版本  ******': None}
        for (windows_editions, gvlk_keys) in g.items():
            self.gvlks['******  {}  ******'.format(windows_editions)] = None
            for (windows_edition, gvlk_key) in gvlk_keys.items():
                self.gvlks[windows_edition] = gvlk_key
        # for (k, v) in self.gvlks.items():
        #     print(str(k).ljust(80), str(v).ljust(30))

    def _install_gvlk(self):
        gvlk_key = self.gvlks[self.windows_gvlk_list.get()]
        if gvlk_key is not None:
            status, output = subprocess.getstatusoutput('cscript {} /ipk {}'.format(self._win_vbs, gvlk_key))
            if status == 0:
                self._insert_content('{}\nWindows的GVLK密钥安装成功\n'.format('\n'.join(output.split('\n')[3:])))
            else:
                self._insert_content(
                    '{}\nWindows的GVLK密钥安装失败\n请选择正确的Windows版本\n'.format('\n'.join(output.split('\n')[3:])))

    def _activate_win(self):
        kms_server = self.kms_server_text.get(0.0, END).strip()

        status, output = subprocess.getstatusoutput('cscript "{}" /skms {}'.format(self._win_vbs, kms_server))

        self._insert_content('{}\n设置Windows的KMS服务器成功\n'.format('\n'.join(output.split('\n')[3:])))

        status, output = subprocess.getstatusoutput('cscript "{}" /ato'.format(self._win_vbs))
        if status == 0:
            self._insert_content('{}\n激活Windows成功\n'.format('\n'.join(output.split('\n')[3:])))
        else:
            self._insert_content('{}\n激活Windows失败\n'.format('\n'.join(output.split('\n')[3:])))

    def _open_kms_server(self):
        t = threading.Thread(target=server.main)
        t.start()
        time.sleep(1)
        if t.is_alive():
            self._insert_content('密钥管理服务成功开启在 127.0.0.1:1688\n')
        else:
            self._insert_content('密钥管理服务开启失败\n通常每个套接字地址(协议/网络地址/端口)只允许使用一次。\n')

    def _get_install_location(self, software_name, keypath):
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, keypath) as software_keys:
            i, result = 0, []
            while 1:
                try:
                    software_key = winreg.EnumKey(software_keys, i)
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, (keypath + '\\' + software_key)) as key:
                        try:
                            display_name, _ = winreg.QueryValueEx(key, 'DisplayName')
                            install_location, _ = winreg.QueryValueEx(key, 'InstallLocation')
                            display_version, _ = winreg.QueryValueEx(key, 'DisplayVersion')
                            # print(display_name, install_location)
                            if software_name in str(display_name).lower() and install_location != '':
                                # print(display_name, install_location)
                                display_version = str(display_version).split('.')[0]
                                result.append([display_name, install_location, display_version])
                        except Exception as e:
                            # print(e)
                            pass
                except Exception as e:
                    # print(e)
                    break
                i += 1
        return result

    def _search_office_vbs(self):
        activated = []
        for keypath in self.keypaths:
            for office_info in self._get_install_location('office', keypath):
                display_name, install_location, display_version = office_info[0], office_info[1], office_info[2]
                office_path = os.path.join(install_location, 'Office{}'.format(display_version))
                vbs_path = os.path.join(install_location, 'Office{}'.format(display_version), 'OSPP.VBS')
                if os.path.isfile(vbs_path) and install_location.strip('\\') not in activated:
                    # print(office_info)
                    activated.append(install_location.strip('\\'))
                    # self._insert_content('Office{}安装位置为{}\n'.format(office_info[2], office_path))
                    # self._activate_office(vbs_path)

                    self.office_location.delete(0.0, END)
                    self.office_location.insert(INSERT, office_path)
                    yield

    def _activate_office(self):
        office_path = self.office_location.get(0.0, END).strip()
        kms_server = self.kms_server_text.get(0.0, END).strip()
        vbs_path = os.path.join(office_path, 'OSPP.VBS')
        if os.path.isfile(vbs_path):
            self._insert_content('Office安装位置为{}\n'.format(office_path))
            status, output = subprocess.getstatusoutput('cscript "{}" /sethst:{}'.format(vbs_path, kms_server))
            self._insert_content('{}\n设置Office的KMS服务器成功\n'.format('\n'.join(output.split('\n')[3:])))

            status, output = subprocess.getstatusoutput('cscript "{}" /act'.format(vbs_path))
            if '<Product activation successful>' in output:
                self._insert_content('{}\n激活Office成功\n'.format('\n'.join(output.split('\n')[3:])))
            else:
                self._insert_content('{}\n激活Office失败\n'.format('\n'.join(output.split('\n')[3:])))
            status, output = subprocess.getstatusoutput('cscript "{}" /dstatus'.format(vbs_path))
            if status == 0:
                self._insert_content('{}\n以上为Office激活的详细信息\n'.format('\n'.join(output.split('\n')[3:])))
        else:
            self._insert_content('Office安装位置错误，请搜索位置或者手动输入。\n')

    def show_activate_frame(self):
        self.activate_frame = ttk.Frame(self.root, padding='3 3 10 10')
        self.activate_frame.grid(column=0, row=0)
        # self.switch_frame.columnconfigure(0, weight=1)
        # self.switch_frame.rowconfigure(0, weight=1)

        ttk.Button(self.activate_frame, text='开启本机KMS服务器', command=self._open_kms_server) \
            .grid(column=0, row=0, columnspan=3, sticky=(W, E))

        ttk.Label(self.activate_frame, text='KMS', font=label_font) \
            .grid(column=0, row=1, sticky=(E, N))
        self.kms_server_text = Text(self.activate_frame, height=1, width=30, font=content_font)
        self.kms_server_text.insert(0.0, '127.0.0.1')
        self.kms_server_text.grid(column=1, row=1, columnspan=2, sticky=(W, E))

        ttk.Label(self.activate_frame, text='Windows', font=label_font) \
            .grid(column=0, row=2, sticky=(E, N))
        self.windows_gvlk_list = ttk.Combobox(self.activate_frame, font=label_font, state='readonly',
                                              values=list(self.gvlks.keys()))
        self.windows_gvlk_list.current(0)
        self.windows_gvlk_list.grid(column=1, row=2, sticky=(W, E))
        ttk.Button(self.activate_frame, text='安装GVLK密钥', command=self._install_gvlk).grid(column=2, row=2)
        ttk.Button(self.activate_frame, text='激活Windows', command=self._activate_win). \
            grid(column=1, columnspan=2, row=3, sticky=(W, E))

        ttk.Label(self.activate_frame, text='Office', font=label_font).grid(column=0, row=4, rowspan=2, sticky=(E, N))
        self.office_location = Text(self.activate_frame, height=2, width=30, font=content_font)
        self.office_location.insert(0.0, 'C:\\location')
        self.office_location.grid(column=1, row=4, sticky=(W, E))
        ttk.Button(self.activate_frame, text='搜索Office位置', command=lambda: next(self.office_location_iterator)) \
            .grid(column=2, row=4, sticky=(W, E))
        ttk.Button(self.activate_frame, text='激活Office', command=self._activate_office). \
            grid(column=1, columnspan=2, row=5, sticky=(W, E))

        for child in self.activate_frame.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def show_content_frame(self):
        self.content_frame = ttk.Frame(self.root, padding='3 3 10 10')
        self.content_frame.grid(column=0, row=2, sticky=(N, W, E, S))
        # self.content_frame.columnconfigure(0, weight=1)
        # self.content_frame.rowconfigure(0, weight=1)

        scrollbarX = Scrollbar(self.content_frame, orient=HORIZONTAL)
        scrollbarX.pack(side=BOTTOM, fill=X, anchor=N)
        scrollbarY = Scrollbar(self.content_frame, orient=VERTICAL)
        scrollbarY.pack(side=RIGHT, fill=Y, anchor=N)

        self.content_text = Text(self.content_frame, width=60, height=15, bg='#000000', fg='#33FF66', wrap='none',
                                 font=("Times New Roman", 10, "normal"), state='disabled',
                                 xscrollcommand=scrollbarX.set, yscrollcommand=scrollbarY.set)
        scrollbarX.configure(command=self.content_text.xview)
        scrollbarY.configure(command=self.content_text.yview)

        # self.content_text = scrolledtext.ScrolledText(self.content_frame, width=80, height=15, bg='black', fg='green',
        #                                               wrap='none',
        #                                               font=("Times New Roman", 10, "bold"),
        #                                               )
        # self.content_text.insert(INSERT, quote)
        # self.content_text.configure(state='disabled')
        # self.content_text.configure(state='normal')
        self.content_text.pack(expand=YES, fill=BOTH)
        # self.content_text.grid(column=0, row=0, sticky=(N, W, E, S))

    def _insert_content(self, content):
        self.content_text.configure(state='normal')
        self.content_text.insert(END, '*' * 50 + '\n')
        self.content_text.insert(END, content)
        self.content_text.configure(state='disabled')

    def show(self):
        self.root.mainloop()


def main():
    my_frame = My_frame()
    my_frame.show()


if __name__ == '__main__':
    main()
