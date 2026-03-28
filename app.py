import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
import os

# =========================
# SETTINGS
# =========================
QR_SIZE = 260
TEXT_COLOR = (255, 102, 0)
FONT_SIZE = 48
LOGO_PATH = "logo.png"

# =========================
# LOAD FONT
# =========================
def load_font():
    try:
        return ImageFont.truetype("arialbd.ttf", FONT_SIZE)
    except:
        return ImageFont.load_default()

# =========================
# GENERATE QR
# =========================
def generate_qr(data):
    qr = qrcode.make(data)
    return qr.resize((QR_SIZE, QR_SIZE))

# =========================
# CREATE STICKER
# =========================
def create_sticker(building_code, national_address, logo):
    font = load_font()

    qr_data = f"{building_code} - {national_address}"
    qr_img = generate_qr(qr_data)

    logo = logo.resize((QR_SIZE, int(logo.height * (QR_SIZE / logo.width))))

    # canvas
    width = QR_SIZE + 40
    height = logo.height + QR_SIZE + 120
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    # positions
    x_center = width // 2

    y_logo = 10
    y_qr = y_logo + logo.height + 20
    y_text = y_qr + QR_SIZE + 30

    # paste
    img.paste(logo, (x_center - logo.width // 2, y_logo))
    img.paste(qr_img, (x_center - QR_SIZE // 2, y_qr))

    # text
    draw.text(
        (x_center, y_text),
        building_code,
        fill=TEXT_COLOR,
        font=font,
        anchor="mm"
    )

    return img

# =========================
# MAIN
# =========================
input_file = "JED-HRR2-SALH-11-MP.xlsx"
output_file = "FINAL_QR_Customers.xlsx"

df = pd.read_excel(input_file, sheet_name="Customers")
logo = Image.open(LOGO_PATH)

wb = Workbook()
ws = wb.active
ws.title = "Customers_QR"

ws.append(["SR", "Building Code", "National Address", "QR"])

for i, row in df.iterrows():
    building = str(row["Building Code"])
    address = str(row["National Address"])

    img = create_sticker(building, address, logo)

    img_path = f"temp_{i}.png"
    img.save(img_path)

    ws.append([row["SR"], building, address])

    xl_img = XLImage(img_path)
    xl_img.width = 180
    xl_img.height = 220

    ws.add_image(xl_img, f"D{i+2}")

wb.save(output_file)
