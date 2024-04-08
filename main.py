import openpyxl
import os
import configparser
from reportlab.graphics import barcode
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.renderPM import drawToFile
from PIL import Image, ImageDraw, ImageFont

config = configparser.ConfigParser()
config.read("config.ini")

BC_WIDTH = int(config["BARCODE"]["width"])
BC_HEIGHT = int(config["BARCODE"]["height"])
BC_x = int(config["BARCODE"]["x"])
BC_y = int(config["BARCODE"]["y"])

TEXT_x = int(config["TEXT"]["text_x"])
TEXT_y = int(config["TEXT"]["text_y"])
TEXT_SIZE = int(config["TEXT"]["font_size"])
TEXT_COLOR = config["TEXT"]["font_color"]

EXCEL_EXT = [".xsl", ".xlsx", ".XSL", ".XLSX"]
PICTURE_EXT = [".jpg", ".jpeg", ".bmp", ".png", ".JPG", ".JPEG", ".BMP", ".PNG"]
FONT_EXT = [".TTF", ".ttf"]

BARCODE_FILE = "_barcode.png"
RESULT_DIR = "results"


def create_barcode(code):
    draw = Drawing(BC_WIDTH, BC_HEIGHT)
    new_barcode = barcode.createBarcodeDrawing('EAN13', value=code, width=BC_WIDTH, height=BC_HEIGHT)
    draw.add(new_barcode)
    drawToFile(draw, BARCODE_FILE)


def put_barcode_to_cert(image):
    bc = Image.open(BARCODE_FILE)
    image.paste(bc, (BC_x, BC_y))
    bc.close()
    return image


def put_text_to_cert(image, text):
    draw = ImageDraw.Draw(image)
    font = get_font()
    text_font = ImageFont.truetype(font=font, size=TEXT_SIZE)
    draw.text((TEXT_x, TEXT_y), text, font=text_font, fill=TEXT_COLOR)
    return image


def insert_data_to_picture(cert_filename, code, price):
    create_barcode(code)
    cert = Image.open(cert_filename).convert('RGB')
    cert = put_barcode_to_cert(cert)
    cert = put_text_to_cert(cert, f"{price} ₽")
    cert.save(f"{RESULT_DIR}/{code}.jpg",
              format="JPEG",
              quality=100,
              icc_profile=cert.info.get('icc_profile',''))
    cert.close()
    os.remove(BARCODE_FILE)


def get_data_from_xsl(file_name):
    workbook = openpyxl.load_workbook(file_name)
    worksheet = workbook.active

    for i in range(0, worksheet.max_row):
        data = []
        for col in worksheet.iter_cols(1, 2):
            if col[i].value is not None:
                data.append(col[i].value)
        if len(data) == 2:
            yield data


def create_dir(name):
    if not os.path.exists(name):
        os.mkdir(name)


def get_filename(extension):
    files_list = [f for f in os.listdir(".") if os.path.isfile(f) and os.path.splitext(f)[1] in extension]
    if files_list:
        return files_list[0]
    print(f"Не найдено файла с расширением {extension}")
    return None


def get_font():
    font_file = get_filename(FONT_EXT)
    if font_file:
        return font_file
    return config["TEXT"]["font"]


def remove_tmp_files():
    tmp_files = [f for f in os.listdir(".") if os.path.isfile(f) and os.path.splitext(f)[1] in PICTURE_EXT]
    for f in tmp_files:
        if f.startswith("_"):
            os.remove(f)


def init():
    create_dir(RESULT_DIR)
    remove_tmp_files()


if __name__ == "__main__":
    init()
    create_dir(RESULT_DIR)
    excel_filename = get_filename(EXCEL_EXT)
    cert_filename = get_filename(PICTURE_EXT)

    for data in get_data_from_xsl(excel_filename):
        insert_data_to_picture(cert_filename, *data)
