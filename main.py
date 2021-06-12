import sys
import pyocr
import pyocr.builders
import cv2
import gspread
import os
from PIL import Image
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

load_dotenv(verbose=True)

sys.path.append('/Users/shotaro/projects/umamusume-analyzer')

tools = pyocr.get_available_tools()
if len(tools) == 0:
    print("No OCR tool found")
    sys.exit(1)
tool = tools[0]


def ocr(img):
    img_pil = Image.fromarray(img)
    txt = tool.image_to_string(
        img_pil,
        lang='jpn+eng',
        builder=pyocr.builders.TextBuilder()
    )
    return txt


def analyze_image(path):

    img = cv2.imread(os.path.abspath(path))
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    # character
    character = img[390:440, 450:800]
    character = ocr(character)

    # status
    speed = img[580:650, 90:190]
    stamina = img[580:650, 250:350]
    power = img[580:650, 410:510]
    guts = img[580:650, 570:670]
    wise = img[580:650, 730:830]

    speed = ocr(speed)
    stamina = ocr(stamina)
    power = ocr(power)
    guts = ocr(guts)
    wise = ocr(wise)

    return ['', speed, stamina, power, guts, wise]


def connect_gspread():
    jsonf = os.environ.get('CREDENCIAL_PATH', '')
    key = os.environ.get('SHEET_KEY', '')
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        jsonf, scope)
    gc = gspread.authorize(credentials)
    worksheet = gc.open_by_key(key).worksheet('育成ウマ娘')
    return worksheet


def main():
    if len(sys.argv) <= 1:
        print("path not found")
        sys.exit(1)

    path = sys.argv[1]
    sheet = connect_gspread()

    cells = sheet.get_all_values()
    uma_musume = analyze_image(path)

    if len(cells[-1]) > 0:
        id = int(cells[-1][0])+1 if len(cells) > 2 else 1
        sheet.update_cell(len(cells)+1, 1, id)
        for index, txt in enumerate(uma_musume):
            sheet.update_cell(len(cells)+1, index+2, txt)


if __name__ == "__main__":
    main()
