'''
@author Cheain
@Github https://github.com/Cheain/kms_tool
'''
import base64
import os
import threading
import time
import tkinter as tk
import winreg
from subprocess import getstatusoutput
from tkinter import ttk

import server
from GVLK import GVLKs as g
from kms_icon import icon

label_font = ('Times New Roman', 12, 'normal')
content_font = ('Times New Roman', 14, 'normal')


class KMSTool():
    def __init__(self):
        self._win_vbs = '%windir%\system32\slmgr.vbs'
        self._office_keypaths = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                                 r'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall']

        self.root = tk.Tk()
        self.root.title('KMS激活')
        # self.root.geometry('600x800')
        with open('tmp.ico', 'wb+') as tmp:
            tmp.write(base64.b64decode(icon))
        self.root.iconbitmap('tmp.ico')
        os.remove('tmp.ico')

        self.root.resizable(False, False)

        self._get_gvlks()
        self.show_activate_frame()
        self.show_content_frame()

        self._office_location_iterator = self._search_office_vbs()

    def _get_gvlks(self):
        self.gvlks = {'******  选择Windows版本  ******': None}
        for (windows_editions, gvlk_keys) in g.items():
            self.gvlks['******  {}  ******'.format(windows_editions)] = None
            for (windows_edition, gvlk_key) in gvlk_keys.items():
                self.gvlks[windows_edition] = gvlk_key

    def open_kms_server(self):
        t = threading.Thread(target=server.main, daemon=True)
        t.start()
        time.sleep(1)
        if t.is_alive():
            self.insert_content(text_num=0, tag='green')
        else:
            self.insert_content(text_num=1, tag='red')

    def install_gvlk(self):
        gvlk_key = self.gvlks[self.windows_gvlk_list.get()]
        if gvlk_key is not None:
            self.insert_content('尝试为本机安装 {} 的GVLK密钥 {}'.format(self.windows_gvlk_list.get(), gvlk_key))
            status, output = getstatusoutput('cscript {} /ipk {}'.format(self._win_vbs, gvlk_key))
            if status == 0:
                self.insert_content('\n'.join(output.split('\n')[3:]), 2, 'green')
            else:
                self.insert_content('\n'.join(output.split('\n')[3:]), 3, 'red')
        else:
            self.insert_content('请选择操作系统版本')

    def activate_win(self):
        kms_server = self.kms_server_text.get(0.0, tk.END).strip()

        status, output = getstatusoutput('cscript "{}" /skms {}'.format(self._win_vbs, kms_server))

        self.insert_content('\n'.join(output.split('\n')[3:]), 4, 'green')
        self.insert_content('当前KMS服务器为 {}'.format(kms_server))

        status, output = getstatusoutput('cscript "{}" /ato'.format(self._win_vbs))
        if status == 0:
            self.insert_content('\n'.join(output.split('\n')[3:]), 5, 'green')
        else:
            self.insert_content('\n'.join(output.split('\n')[3:]), 6, 'red')

        status, output = getstatusoutput('cscript "{}" /dlv'.format(self._win_vbs))
        self.insert_content('\n'.join(output.split('\n')[3:]), 7, 'green')

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
                            if software_name in str(display_name).lower() and install_location != '':
                                display_version = str(display_version).split('.')[0]
                                result.append([display_name, install_location, display_version])
                        except Exception as e:
                            self.insert_content(e, 14, 'red')
                except Exception as e:
                    self.insert_content(e, 14, 'red')
                    break
                i += 1
        return result

    def _search_office_vbs(self):
        activated = []
        for keypath in self._office_keypaths:
            for office_info in self._get_install_location('office', keypath):
                display_name, install_location, display_version = office_info[0], office_info[1], office_info[2]
                office_path = os.path.join(install_location, 'Office{}'.format(display_version))
                vbs_path = os.path.join(install_location, 'Office{}'.format(display_version), 'OSPP.VBS')
                if os.path.isfile(vbs_path) and install_location.strip('\\') not in activated:
                    activated.append(install_location.strip('\\'))
                    self.office_location.delete(0.0, tk.END)
                    self.office_location.insert(tk.INSERT, office_path)
                    yield

    def activate_office(self):
        office_path = self.office_location.get(0.0, tk.END).strip()
        kms_server = self.kms_server_text.get(0.0, tk.END).strip()
        vbs_path = os.path.join(office_path, 'OSPP.VBS')
        if os.path.isfile(vbs_path):
            self.insert_content(office_path, 8, 'green')
            status, output = getstatusoutput('cscript "{}" /sethst:{}'.format(vbs_path, kms_server))
            self.insert_content('\n'.join(output.split('\n')[3:]), 9, tag='green')
            self.insert_content('当前KMS服务器为 {}'.format(kms_server))

            status, output = getstatusoutput('cscript "{}" /act'.format(vbs_path))
            if '<Product activation successful>' in output:
                self.insert_content('\n'.join(output.split('\n')[3:]), 10, 'green')
            else:
                self.insert_content('\n'.join(output.split('\n')[3:]), 11, 'red')
            status, output = getstatusoutput('cscript "{}" /dstatus'.format(vbs_path))
            if status == 0:
                self.insert_content('\n'.join(output.split('\n')[3:]), 12, 'green')
        else:
            self.insert_content(text_num=13, tag='red')

    def show_activate_frame(self):
        self.activate_frame = ttk.Frame(self.root, padding='3 3 10 10')
        self.activate_frame.grid(column=0, row=0)

        ttk.Button(self.activate_frame, text='开启本机KMS服务器', command=self.open_kms_server) \
            .grid(column=0, row=0, columnspan=3, sticky=(tk.W, tk.E))

        ttk.Label(self.activate_frame, text='KMS', font=label_font).grid(column=0, row=1, sticky=(tk.E, tk.N))
        self.kms_server_text = tk.Text(self.activate_frame, height=1, width=30, font=content_font)
        self.kms_server_text.insert(0.0, '127.0.0.1')
        self.kms_server_text.grid(column=1, row=1, columnspan=2, sticky=(tk.W, tk.E))

        ttk.Label(self.activate_frame, text='Windows', font=label_font).grid(column=0, row=2, sticky=(tk.E, tk.N))
        self.windows_gvlk_list = ttk.Combobox(self.activate_frame, font=label_font, state='readonly',
                                              values=list(self.gvlks.keys()))
        self.windows_gvlk_list.current(0)
        self.windows_gvlk_list.grid(column=1, row=2, sticky=(tk.W, tk.E))
        ttk.Button(self.activate_frame, text='安装GVLK密钥', command=self.install_gvlk).grid(column=2, row=2)
        ttk.Button(self.activate_frame, text='激活Windows', command=self.activate_win). \
            grid(column=1, columnspan=2, row=3, sticky=(tk.W, tk.E))

        ttk.Label(self.activate_frame, text='Office', font=label_font).grid(column=0, row=4, rowspan=2,
                                                                            sticky=(tk.E, tk.N))
        self.office_location = tk.Text(self.activate_frame, height=2, width=30, font=content_font)
        self.office_location.insert(0.0, 'C:\\location')
        self.office_location.grid(column=1, row=4, sticky=(tk.W, tk.E))
        ttk.Button(self.activate_frame, text='搜索Office位置', command=lambda: next(self._office_location_iterator)) \
            .grid(column=2, row=4, sticky=(tk.W, tk.E))
        ttk.Button(self.activate_frame, text='激活Office', command=self.activate_office). \
            grid(column=1, columnspan=2, row=5, sticky=(tk.W, tk.E))

        for child in self.activate_frame.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def show_content_frame(self):
        self.content_frame = ttk.Frame(self.root, padding='3 3 10 10')
        self.content_frame.grid(column=0, row=2, sticky=(tk.N, tk.W, tk.E, tk.S))
        # self.content_frame.columnconfigure(0, weight=1)
        # self.content_frame.rowconfigure(0, weight=1)

        scrollbarX = tk.Scrollbar(self.content_frame, orient=tk.HORIZONTAL)
        scrollbarX.pack(side=tk.BOTTOM, fill=tk.X, anchor=tk.N)
        scrollbarY = tk.Scrollbar(self.content_frame, orient=tk.VERTICAL)
        scrollbarY.pack(side=tk.RIGHT, fill=tk.Y, anchor=tk.N)

        self.content_text = tk.Text(self.content_frame, width=60, height=15, bg='#242424', fg='#33FF66', wrap='none',
                                    font=('Times New Roman', 10, 'normal'), state='disabled',
                                    xscrollcommand=scrollbarX.set, yscrollcommand=scrollbarY.set)
        scrollbarX.configure(command=self.content_text.xview)
        scrollbarY.configure(command=self.content_text.yview)

        self.content_text.tag_configure('green', foreground='#00FF00')  # success
        self.content_text.tag_configure('red', foreground='#FF0000')  # fail
        self.content_text.tag_configure('yellow', foreground='#FFFF00')  # normal
        self.content_text.tag_configure('blue', foreground='#00FFFF')

        self.content_text.pack(expand=tk.YES, fill=tk.BOTH)

    def insert_content(self, content=None, text_num=None, tag='yellow'):
        output_texts = {
            0: '密钥管理服务成功开启在 127.0.0.1:1688\n',
            1: '密钥管理服务开启失败\n通常每个套接字地址(协议/网络地址/端口)只允许使用一次。\n',
            2: 'Windows的GVLK密钥安装成功\n',
            3: 'Windows的GVLK密钥安装失败\n请选择正确的Windows版本\n',
            4: '设置Windows的KMS服务器成功\n',
            5: '激活Windows成功\n',
            6: '激活Windows失败\n',
            7: '以上为Windows激活的详细信息\n',
            8: 'Office安装位置如上\n',
            9: '设置Office的KMS服务器成功\n',
            10: '激活Office成功\n',
            11: '激活Office失败\n',
            12: '以上为Office激活的详细信息\n',
            13: 'Office安装位置错误，请搜索位置或者手动输入。\n',
            14: '出错啦！！！以上为报错信息。\n',
            20: '▃' * 62 + '\n',
        }
        self.content_text.configure(state='normal')
        self.content_text.insert(tk.END, output_texts[20], 'blue')
        if content is not None:
            self.content_text.insert(tk.END, str(content).strip() + '\n', 'yellow')
        if text_num is not None:
            self.content_text.insert(tk.END, output_texts[text_num], tag)
        self.content_text.configure(state='disabled')
        self.content_text.see(tk.END)

    def show(self):
        self.root.mainloop()


def main():
    my_frame = KMSTool()
    my_frame.show()


if __name__ == '__main__':
    main()
