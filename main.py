"""
TODO:
Уточнить, какой формат штрих-кода нужен
"""


import openpyxl
import os
from reportlab.graphics import barcode
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.renderPM import drawToFile
from PIL import Image, ImageDraw, ImageFont

BARCODE_FILE = "barcode.png"
CERT_FILE = "certificate.jpg"
RESULT_DIR = "results"

BC_WIDTH = 300
BC_HEIGHT = 100
BC_x = 450
BC_y = 300
TEXT_FONT = "font.ttf"
TEXT_x = 450
TEXT_y = 150
TEXT_SIZE = 100
TEXT_COLOR = "#56100A"
XLS_FILE = "test.xlsx"


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
    font = ImageFont.truetype(font=TEXT_FONT, size=TEXT_SIZE)
    draw.text((TEXT_x, TEXT_y), text, font=font, fill=TEXT_COLOR)
    return image


def insert_data_to_picture(code, price):
    create_barcode(code)
    cert = Image.open(CERT_FILE)
    cert = put_barcode_to_cert(cert)
    cert = put_text_to_cert(cert, f"{price} ₽")
    cert.save(f"{RESULT_DIR}/{code}.jpg")
    cert.close()
    os.remove(BARCODE_FILE)


def get_data_from_xsl(file_name):
    print("GENERATOR")
    workbook = openpyxl.load_workbook(file_name)
    worksheet = workbook.active

    for i in range(0, worksheet.max_row):
        data = []
        for col in worksheet.iter_cols(1, 2):
            if col[i].value is not None:
                data.append(col[i].value)
        if len(data) == 2:
            yield data


def create_dir():
    if not os.path.exists(RESULT_DIR):
        os.mkdir(RESULT_DIR)


if __name__ == "__main__":
    create_dir()
    for data in get_data_from_xsl(XLS_FILE):
        print(data)
        insert_data_to_picture(*data)
