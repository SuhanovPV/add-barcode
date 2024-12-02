import os.path
import sys

from PyQt6.QtCore import pyqtSlot
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QTabWidget, QHBoxLayout, QLabel, QWidget, \
    QGridLayout, QLineEdit, QVBoxLayout, QFileDialog, QMessageBox
from utils.config_manager import ConfigManager
from utils import add_barcode, excel_helper


# TODO добавить возможность настройки параметров вставки штрих-кода и надписи в UI
# TODO добавить возможность выбора файлов для обработки файлов на дубликаты и фейковые телефоны
# TODO добавить в конфиг шрифты


class MainWindow(QMainWindow):
    # sender_btn = QObject.sender(self) - получить отправителя сигнала
    # TODO добавить функцию обрезки текста для полей, если не помещается
    # TODO заменять разделители в тексте согласно OS
    # TODO добавить подсказку для полей с полным путем

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config = ConfigManager()
        self.barcode_xls_data_file = ''
        self.barcode_work_dir_path = self.config.get_path_work_dir()
        self.barcode_result_dir = self.config.get_path_result_dir()
        self.barcode_template_file = self.config.get_path_template_file()

        self.setWindowTitle("Helper")
        self.setMinimumSize(550, 350)
        self.statusBar().showMessage("Выберите файл")

        # TODO добавить текст в status bar в зависимости
        # Tab widget
        tab = QTabWidget(self)

        # Barcode page
        barcode_page = self._crate_barcode_page()

        # Work with xlsx
        excel_page = QWidget(self)
        excel_layout = QHBoxLayout()
        excel_page.setLayout(excel_layout)

        select_excel_file = QPushButton("Выбрать файл")
        excel_layout.addWidget(select_excel_file)

        # add pane to tab widget
        tab.addTab(barcode_page, 'Создание листовок')
        tab.addTab(excel_page, 'Обработка excel')

        main_layout = QGridLayout()
        main_layout.addWidget(tab)

        container = QWidget()
        container.setLayout(main_layout)

        self.setCentralWidget(container)
        self.show()

    def closeEvent(self, event):
        self.config.save_config()

    def _crate_barcode_page(self):
        widget = QWidget(self)
        widget_layout = QVBoxLayout()

        data_file_layout = self._add_row_h_layout(
            lbl_text='Файл c данными',
            input_name='data_file_input',
            btn_text='Выбрать',
            btn_func=self.click_barcode_select_file_button
        )
        result_dir_layout = self._add_row_h_layout(
            lbl_text='Папка с результатами',
            input_name='result_dir_input',
            btn_text='Выбрать',
            btn_func=self.click_barcode_select_result_folder,
            input_text=self.config.get_path_result_dir()
        )

        result_template_layout = self._add_row_h_layout(
            lbl_text='Шаблон',
            input_name='template_input',
            btn_text='Выбрать',
            btn_func=self.click_barcode_select_template_file,
            input_text=self.config.get_path_template_file()
        )

        button_layout = QHBoxLayout()
        btn = QPushButton('Создать')
        btn.setObjectName('barcode_btn')
        btn.setDisabled(True)
        # btn.clicked.connect(self.click_barcode_create_leaflets)
        btn.pressed.connect(self.click_barcode_create_leaflets)
        button_layout.addStretch()
        button_layout.addWidget(btn)

        widget_layout.addLayout(data_file_layout)
        widget_layout.addLayout(result_dir_layout)
        widget_layout.addLayout(result_template_layout)
        widget_layout.addStretch()
        widget_layout.addLayout(button_layout)
        widget.setLayout(widget_layout)
        return widget

    @staticmethod
    def _add_row_h_layout(lbl_text, input_name, btn_text, btn_func, input_text="Не выбрано"):
        layout = QHBoxLayout()

        label = QLabel(lbl_text)
        label.setMinimumWidth(120)
        line_input = QLineEdit()
        line_input.setReadOnly(True)
        line_input.setText(input_text)
        line_input.setObjectName(input_name)
        button = QPushButton(btn_text)
        button.clicked.connect(btn_func)

        layout.addWidget(label)
        layout.addWidget(line_input)
        layout.addWidget(button)

        return layout

    @pyqtSlot()
    def click_barcode_select_file_button(self):
        file_name = self.open_file_dialog(
            title='Выберите файл с данными',
            folder=self.config.get_path_work_dir(),
            extension='Excel files (*.xlsx *.xls)'
        )

        if file_name:
            self.barcode_xls_data_file = file_name
            folder, file = os.path.split(self.barcode_xls_data_file)
            line_input = self._get_element_by_name(QLineEdit, 'data_file_input')
            line_input.setText(file)
            self.config.set_path_work_dir(folder)
            btn = self._get_element_by_name(QPushButton, 'barcode_btn')
            btn.setDisabled(False)

    @pyqtSlot()
    def click_barcode_select_result_folder(self):
        folder = self.open_dir_dialog(title="Выберите папку, куда будут сохраняться изображения")
        if folder:
            self.barcode_result_dir = folder
            self.config.set_path_result_dir(folder)
            line_input = self._get_element_by_name(QLineEdit, 'result_dir_input')
            line_input.setText(folder)

    @pyqtSlot()
    def click_barcode_select_template_file(self):
        file = self.open_file_dialog(
            title='Выберите изображение в качестве шаблона',
            folder=self.config.get_path_work_dir(),
            extension='Image jpg (*.jpg *.jpeg)'
        )
        if file:
            self.barcode_template_file = file
            self.config.set_path_template_file(file)

    @pyqtSlot()
    def click_barcode_create_leaflets(self):
        if self.is_files_exist([self.barcode_xls_data_file, self.barcode_template_file]) and \
                self.is_folder_exist_or_create(self.barcode_result_dir):
            for code, price in excel_helper.get_barcode_data_from_xsl(self.barcode_xls_data_file):
                add_barcode.create_leaflets(code, price, self.barcode_result_dir, self.barcode_template_file,
                                            self.config)

    def open_file_dialog(self, title, folder, extension):
        return QFileDialog.getOpenFileName(self, title, folder, f"{extension};; All Files (*)")[0]

    def open_dir_dialog(self, title):
        open_dir = QFileDialog.getExistingDirectory(self, title)
        return open_dir

    def _get_element_by_name(self, widget_type, name):
        return self.findChildren(widget_type, name)[0]

    def is_files_exist(self, files):
        for file in files:
            if not os.path.exists(file):
                self.show_message(f'Файл\n{file}\n не существует.\nПожалуйста, укажите другой файл!')
                return False
        return True

    @staticmethod
    def is_folder_exist_or_create(folder):
        if not os.path.exists(folder):
            os.mkdir(folder)
        return True

    @staticmethod
    def show_message(message):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle('Внимание!')
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    app.exec()
