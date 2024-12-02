from configparser import ConfigParser
import os.path
import utils


class ConfigManager:
    CUR_DIR_PATH = os.path.dirname(os.path.abspath(utils.__file__))
    CONFIG_FILE = os.path.join(CUR_DIR_PATH, "config.ini")

    def __init__(self):
        self.config = ConfigParser()
        self.config.read(self.CONFIG_FILE)

    def get_path_work_dir(self):
        self.config['PATH']['dir_with_file'] = self.config['PATH']['dir_with_file'] or \
                                               os.path.abspath(os.path.join(self.CUR_DIR_PATH, ".."))
        return self.config['PATH']['dir_with_file']

    def set_path_work_dir(self, value):
        self.config['PATH']['dir_with_file'] = value

    def get_path_result_dir(self):
        self.config['PATH']['result_dir'] = self.config['PATH']['result_dir'] or \
                                            os.path.abspath(os.path.join(self.CUR_DIR_PATH, "..", "result"))
        return self.config['PATH']['result_dir']

    def set_path_result_dir(self, value):
        self.config['PATH']['result_dir'] = value

    def get_path_template_file(self):
        self.config['PATH']['template'] = self.config['PATH']['template'] or \
                                          os.path.abspath(os.path.join(self.CUR_DIR_PATH, "..", "template.jpg"))
        return self.config['PATH']['template']

    def set_path_template_file(self, file):
        self.config['PATH']['template'] = file

    @property
    def text_font_size(self):
        return self.config['TEXT']['font_size']

    @property
    def text_font_color(self):
        return self.config['TEXT']['font_color']

    @property
    def text_x(self):
        return self.config['TEXT']['text_x']

    @property
    def text_y(self):
        return self.config['TEXT']['text_y']

    @property
    def barcode_height(self):
        return self.config['BARCODE']['height']

    @property
    def barcode_width(self):
        return self.config['BARCODE']['width']

    @property
    def barcode_color(self):
        return self.config['BARCODE']['color']

    @property
    def barcode_text_color(self):
        return self.config['BARCODE']['text_color']

    @property
    def barcode_border_v(self):
        return self.config['BARCODE']['border_v']

    @property
    def barcode_border_h(self):
        return self.config['BARCODE']['border_h']

    @property
    def barcode_x(self):
        return self.config['BARCODE']['x']

    @property
    def barcode_y(self):
        return self.config['BARCODE']['y']

    def save_config(self):
        with open(self.CONFIG_FILE, "w") as file:
            self.config.write(file)


if __name__ == "__main__":
    c = ConfigManager()
    print(c.get_path_result_dir())
