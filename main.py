import configparser
import os
import tkinter as tk

from tkinter import ttk, filedialog


class MainApp:
    CUR_DIR_PATH = os.path.dirname(os.path.realpath(__file__))
    CONFIG_FILE = os.path.join(CUR_DIR_PATH, "config.ini")

    def __init__(self, app: tk.Tk):
        self.barcode_filepath = ""
        self._read_config()
        self.app = app
        self.app.geometry('680x400')
        self.app.title('Helper')
        self._init_ui()

    def _init_ui(self):
        self.notebook = ttk.Notebook()
        self.notebook.pack(expand=True, fill=tk.BOTH)

        self.frame = ttk.Frame(self.notebook)
        self.notebook.add(self.frame, text="BARCODE")
        self.barcode_btn_open_file = tk.Button(self.frame, width=15, text="Открыть файл",
                                               command=self.open_file_for_barcode)
        self.barcode_btn_open_file.pack(padx=5, pady=10, side=tk.LEFT, anchor=tk.NW)
        self.barcode_lbl_open_file = tk.Label(
            self.frame,
            text=self.barcode_filepath if self.barcode_filepath else "ВЫБЕРИТЕ ФАЙЛ С ДАННЫМИ"
        )
        self.barcode_lbl_open_file.pack(ipady=4, pady=10, side=tk.LEFT, anchor=tk.NW)

    def open_file_for_barcode(self):
        self.barcode_filepath = self._open_file()
        self.barcode_lbl_open_file.config(text=self.barcode_filepath)
        self._write_config()

    def _open_file(self):
        self.barcode_work_dir = self.barcode_work_dir \
            if self.barcode_work_dir else self.CUR_DIR_PATH
        filename = filedialog.askopenfile(initialdir=self.barcode_work_dir).name
        self.barcode_work_dir = os.path.dirname(filename)
        return filename

    def _read_config(self):
        self.config = configparser.ConfigParser()
        self.config.read(self.CONFIG_FILE)
        self.barcode_work_dir = self.config['UI']['work_dir']

    def _write_config(self):
        self.config['UI']['work_dir'] = self.barcode_work_dir
        with open(self.CONFIG_FILE, "w") as configfile:
            self.config.write(configfile)


if __name__ == '__main__':
    root = tk.Tk()
    my_app = MainApp(root)
    root.mainloop()
