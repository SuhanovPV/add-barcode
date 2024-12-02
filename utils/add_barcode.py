import os
import utils
from utils.config_manager import ConfigManager

from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics import barcode
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.renderPM import drawToFile
from PIL import Image, ImageDraw, ImageFont

BC_RATIO = 1.4
CUR_DIR_PATH = os.path.dirname(os.path.abspath(utils.__file__))
BARCODE_FILE = os.path.abspath(os.path.join(CUR_DIR_PATH, "_tmp.jpg"))

FONT_OTF = os.path.abspath(os.path.join(CUR_DIR_PATH, "..", "fonts", "ALS_Granate_Book_1.1.otf"))
FONT_TTF = os.path.abspath(os.path.join(CUR_DIR_PATH, "..", "fonts", "ALS_Granate_Book_1.1.ttf"))


def create_barcode(code, width, color, text_color):
    """Save tmp barcode file with 'code' data"""
    height = int(width / BC_RATIO)
    draw = Drawing(width, height)
    pdfmetrics.registerFont(TTFont('Granate_Book', FONT_TTF))
    new_barcode = barcode.createBarcodeDrawing('EAN13', value=code, height=height, width=width,
                                               barFillColor=HexColor(color), fontName='Granate_Book',
                                               textColor=HexColor(text_color))
    draw.add(new_barcode)
    drawToFile(draw, BARCODE_FILE)


def is_need_crop(width, height):
    return width / BC_RATIO > height


def put_barcode_to_cert(image, x, y, width, height):
    bc = Image.open(BARCODE_FILE)
    if is_need_crop(width, height):
        image.paste(bc.crop((0, bc.height - height, bc.width, bc.height)), (x, y))
    else:
        image.paste(bc, (x, y))
    bc.close()
    return image


def put_barcode_background(image, x, y, border_h, border_v, width, height):
    draw = ImageDraw.Draw(image)
    draw.rectangle((x - border_h, y, x + width, y + height + border_v), fill="#FFFFFF")
    return image


def put_text_to_cert(image, text, x, y, text_size, text_color):
    draw = ImageDraw.Draw(image)
    font = FONT_OTF
    text_font = ImageFont.truetype(font=font, size=text_size)
    draw.text((x, y), text, font=text_font, fill=text_color)
    return image


def insert_data_to_picture(code: str, price: str, template: str, config: ConfigManager):
    create_barcode(
        code=code,
        width=int(config.barcode_width),
        color=config.barcode_color,
        text_color=config.barcode_text_color
    )

    cert = Image.open(template).convert('RGB')
    put_barcode_background(
        image=cert,
        x=int(config.barcode_x),
        y=int(config.barcode_y),
        border_v=int(config.barcode_border_v),
        border_h=int(config.barcode_border_h),
        width=int(config.barcode_width),
        height=int(config.barcode_height)
    )

    cert = put_barcode_to_cert(
        image=cert,
        x=int(config.barcode_x),
        y=int(config.barcode_y),
        width=int(config.barcode_width),
        height=int(config.barcode_height)
    )
    # TODO добавить вычисление координаты X в зависимости от длинны строки
    cert = put_text_to_cert(
        image=cert,
        text=f"{price} ₽",
        x=int(config.text_x),
        y=int(config.text_y),
        text_size=int(config.text_font_size),
        text_color=config.text_font_color
    )
    return cert


def save_image(image, name, path):
    image.save(f"{path}/{name}.jpg",
               format="JPEG",
               quality=100,
               icc_profile=image.info.get('icc_profile', '')
               )
    image.close()
    os.remove(BARCODE_FILE)


def create_leaflets(code: str, price: str, output_dir: str, template: str, config: ConfigManager):
    image = insert_data_to_picture(code, price, template, config)
    save_image(image, code, output_dir)


if __name__ == "__main__":
    print(BARCODE_FILE)
    print(CUR_DIR_PATH)
