import json


def load_config():
    with open('config.json', 'r') as file:
        cfg = json.load(file)
    return cfg


config = load_config()

scale = 1.0     # 屏幕缩放比例

img_dir = './img/shuiwu/100/'

excel_dir = './info.xlsx'
