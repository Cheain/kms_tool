from tkinter import *
from tkinter import ttk
from tkinter import scrolledtext
import threading
import json
import os

import server

label_font = ("Times New Roman", 12, "normal")
content_font = ("Times New Roman", 14, "normal")


class My_frame():
    def __init__(self):
        self.root = Tk()
        self.root.title('KMS激活')
        # self.root.geometry('600x800')
        self.root.resizable(width=False, height=False)

        self._get_gvlks()
        self.show_activate_frame()
        self.show_content_frame()

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
        for (k, v) in self.gvlks.items():
            print(str(k).ljust(80), str(v).ljust(30))

    def _install_gvlk(self, vbs_path, gvlk_key):
        output=os.popen('cscript {} /ipk {}'.format(vbs_path, gvlk_key)).read()
        print(output)

    def show_activate_frame(self):
        self.activate_frame = ttk.Frame(self.root, padding='3 3 10 10')
        self.activate_frame.grid(column=0, row=0)
        # self.switch_frame.columnconfigure(0, weight=1)
        # self.switch_frame.rowconfigure(0, weight=1)

        ttk.Button(self.activate_frame, text='开启KMS服务器',
                   command=lambda: threading.Thread(target=server.main, daemon=True).start()) \
            .grid(column=0, row=0, columnspan=3, sticky=(W, E))

        ttk.Label(self.activate_frame, text='Windows', font=label_font) \
            .grid(column=0, row=1, rowspan=2, sticky=(E, N))
        self.windows_gvlk_list = ttk.Combobox(self.activate_frame, values=list(self.gvlks.keys()), font=label_font)
        self.windows_gvlk_list.current(0)
        self.windows_gvlk_list.grid(column=1, row=1, sticky=(W, E))
        ttk.Button(self.activate_frame, text='安装GVLK密钥', command=self._test()).grid(column=2, row=1)
        ttk.Button(self.activate_frame, text='激活Windows', command=self._test()).grid(column=1, columnspan=2, row=2,
                                                                                     sticky=(W, E))

        ttk.Label(self.activate_frame, text='Office', font=label_font).grid(column=0, row=3, rowspan=2, sticky=(E, N))
        self.office_location = Text(self.activate_frame, height=1, width=30, font=content_font)
        self.office_location.insert(0.0, 'C:\\location')
        self.office_location.grid(column=1, row=3, sticky=(W, E))
        ttk.Button(self.activate_frame, text='搜索Office位置', command=self._search_office_location) \
            .grid(column=2, row=3, sticky=(W, E))
        ttk.Button(self.activate_frame, text='激活Office', command=self._test()).grid(column=1, columnspan=2, row=4,
                                                                                    sticky=(W, E))

        for child in self.activate_frame.winfo_children():
            child.grid_configure(padx=5, pady=5)

    def show_content_frame(self):
        self.content_frame = ttk.Frame(self.root, padding='3 3 10 10')
        self.content_frame.grid(column=0, row=2, sticky=(N, W, E, S))
        # self.content_frame.columnconfigure(0, weight=1)
        # self.content_frame.rowconfigure(0, weight=1)

        quote = """HAMLET: To be, or not to be--that is the question: Whether 'tis nobler in the mind to suffer The slings and arrows of outrageous fortune
        Or to take arms against a sea of troubles
        And by opposing end them. To die, to sleep--
        No more--and by a sleep to say we end
        The heartache, and the thousand natural shocks
        That flesh is heir to. 'Tis a consummation
        Devoutly to be wishedHAMLET: To be, or not to be--that is the question:
        Whether 'tis nobler in the mind to suffer
        The slings and arrows of outrageous fortune
        Or to take arms against a sea of troubles
        And by opposing end them. To die, to sleep--
        No more--and by a sleep to say we end
        The heartache, and the thousand natural shocks
        That flesh is heir to. 'Tis a consummation
        Devoutly to be wishedHAMLET: To be, or not to be--that is the question:
        Whether 'tis nobler in the mind to suffer
        The slings and arrows of outrageous fortune
        Or to take arms against a sea of troubles
        And by opposing end them. To die, to sleep--
        No more--and by a sleep to say we end
        The heartache, and the thousand natural shocks
        That flesh is heir to. 'Tis a consummation
        Devoutly to be wishedHAMLET: To be, or not to be--that is the question:
        Whether 'tis nobler in the mind to suffer
        The slings and arrows of outrageous fortune
        Or to take arms against a sea of troubles
        And by opposing end them. To die, to sleep--
        No more--and by a sleep to say we end
        The heartache, and the thousand natural shocks
        That flesh is heir to. 'Tis a consummation
        Devoutly to be wished."""

        scrollbarX = Scrollbar(self.content_frame, orient=HORIZONTAL)
        scrollbarX.pack(side=BOTTOM, fill=X, anchor=N)
        scrollbarY = Scrollbar(self.content_frame, orient=VERTICAL)
        scrollbarY.pack(side=RIGHT, fill=Y, anchor=N)

        self.content_text = Text(self.content_frame, width=60, height=15, bg='#000000', fg='#33FF66', wrap='none',
                                 font=("Times New Roman", 10, "normal"),
                                 xscrollcommand=scrollbarX.set, yscrollcommand=scrollbarY.set)
        scrollbarX.configure(command=self.content_text.xview)
        scrollbarY.configure(command=self.content_text.yview)

        # self.content_text = scrolledtext.ScrolledText(self.content_frame, width=80, height=15, bg='black', fg='green',
        #                                               wrap='none',
        #                                               font=("Times New Roman", 10, "bold"),
        #                                               )
        self.content_text.insert(INSERT, quote)
        self.content_text.configure(state='disabled')
        self.content_text.pack(expand=YES, fill=BOTH)
        # self.content_text.grid(column=0, row=0, sticky=(N, W, E, S))

    def _search_office_location(self):
        self.office_location.delete(0.0, END)
        self.office_location.insert(INSERT, '123')

    def show(self):
        self.root.mainloop()


def main():
    my_frame = My_frame()
    my_frame.show()


if __name__ == '__main__':
    main()
