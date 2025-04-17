import json
import sys


def load():
    with open('config.json', 'r') as file:
        cfg = json.load(file)
    return cfg


def update(new_config):
    with open('config.json', 'w') as file:
        json.dump(new_config, file, indent=4)


config = load()

Author = 'BeileLiu'

device = None

driver = None

panel = None

panel_root = None

panel_pop = 1

run = False

abort = False

scale = 1.0

base_dir = './'

img_dir = './img/etax/100/'

excel_path = f"{config['excel_path']}"

result_base_dir = f"{config['result_dir']}"

result_dir = f"{result_base_dir}/"

result_excel_path = None

screenshot_dir = f'{result_dir}截图/'

ptax_app_path = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\自然人电子税务局（扣缴端）\自然人电子税务局（扣缴端）.lnk"

task = 'etax'

task_succeed = 0

task_fail = 0

screen_box = (0, 0, 1919, 1079)

browser_box = (22, 136, 1256, 630)

if __name__ == '__main__':
    print('此文件不可单独运行！！')
    sys.exit()
