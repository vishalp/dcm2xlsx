#!/usr/bin/env python3

import pathlib
import threading
import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import Button
if __name__ == '__main__':
    from dcm2xlsx import convert
else:
    from .dcm2xlsx import convert

class GUI(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack()

        self.active_btn_txt = 'Choose DICOM folder'
        self.inactive_btn_txt = 'Busy'
        self.button = Button(text=self.active_btn_txt, command=self.choose_and_convert)
        self.button.pack(fill='x', expand=False)

        self.log = ScrolledText(state='disabled')
        self.log.pack(fill='both', expand=True)

        self.worker = None
        self.monitor = None

    def choose_and_convert(self):
        dicom_directory = pathlib.Path(askdirectory())
        self.button['state'] = 'disabled'
        self.button['text'] = self.inactive_btn_txt
        self.log['state'] = 'normal'
        self.log.delete('1.0', 'end')
        self.log['state'] = 'disabled'
        self.worker = threading.Thread(target=convert, args=(dicom_directory, self.logger))
        self.monitor = threading.Thread(target=self.reactivate_button)
        self.worker.start()
        self.monitor.start()

    def logger(self, msg):
        self.log['state'] = 'normal'
        self.log.insert('end', f'{msg}\n')
        self.log.see('end')
        self.log['state'] = 'disabled'
    
    def reactivate_button(self):
        self.worker.join()
        self.button['state'] = 'normal'
        self.button['text'] = self.active_btn_txt

def main():
    root = tk.Tk()
    root.title('dcm2xlsx')
    root.geometry('480x240')
    gui = GUI(root)
    gui.mainloop()

if __name__ == '__main__':
    main()
