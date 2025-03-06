from selenium.webdriver.edge.service import Service
from collections import namedtuple
from selenium import webdriver
from tkinter import messagebox
from logger import logger
import tkinter as tk
import win32print
import win32gui
import win32con
import pyautogui
import config
import random
import ctypes
import time
import sys
import re


position = namedtuple('position',
                      ['left', 'top', 'width', 'height'])

corpInfo = namedtuple('corpInfo',
                      ['serial', 'name', 'shibiehao', 'shoujihao', 'password'])

deviceInfo = namedtuple('DeviceInfo',
                        ['available', 'logic_w', 'logic_h', 'phys_w', 'phys_h', 'scale_factor'])


def get_current_time():
    current_time = time.localtime()
    formatted_time = time.strftime("%H:%M:%S", current_time)
    return formatted_time


def pop_error_window(error_message, title='错误'):
    root = tk.Tk()  # 创建一个根窗口
    root.withdraw()  # 隐藏根窗口
    messagebox.showerror(title, error_message)  # 显示错误消息框
    root.quit()  # 退出主循环


def re_coord(coordinate):
    return int(coordinate * config.scale)


def click(box, x, y):  # 这个函数的坐标系是基于box的，不是整个屏幕的
    click_x = box.left + re_coord(random.randint(-3, 3) + x)
    click_y = box.top + re_coord(random.randint(-3, 3) + y)
    pyautogui.click(x=click_x, y=click_y)


def set_img_dir(step='shuiwu'):
    config.img_dir = f'./img/{step}/{int(config.scale * 100)}/'
    return config.img_dir


def set_scale():
    device = device_check()
    config.scale = device.scale_factor


def start_browser():
    edge_driver_path = 'msedgedriver.exe'  # 设置EdgeDriver路径
    service = Service(edge_driver_path)  # 创建Edge服务
    driver = webdriver.Edge(service=service)  # 创建Edge浏览器实例
    driver.set_window_size(1280, 768)  # 设置浏览器窗口大小
    driver.set_window_position(10, 10)  # 设置浏览器窗口位置
    return driver


def check_img(img_name, box):
    try:
        img = pyautogui.locateOnScreen(f'{config.img_dir}{img_name}.png', region=box, confidence=0.9)
        if img is not None:
            if img_name != 'waiting_code_p1':
                logger.info(f"{img_name} Detected.")
            return True
        else:
            return False
    except pyautogui.ImageNotFoundException:
        return False
    except Exception as e:
        logger.error(f'Unknown Error Occurred While Checking Image {img_name}.\n')
        logger.exception(e)
        raise UnknownError(f"Unknown Error Occurred While Checking Image {img_name}.Info:{e}\n")


def device_check():
    # 设置 DPI 感知模式，不可或缺
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    # 获取物理分辨率
    hDC = win32gui.GetDC(0)
    phys_w = win32print.GetDeviceCaps(hDC, win32con.DESKTOPHORZRES)  # 横向物理分辨率
    phys_h = win32print.GetDeviceCaps(hDC, win32con.DESKTOPVERTRES)  # 纵向物理分辨率
    # 获取系统DPI以及缩放比例
    user32 = ctypes.WinDLL('user32')
    GetDpiForSystem = user32.GetDpiForSystem
    GetDpiForSystem.argtypes = [ctypes.c_uint32]
    GetDpiForSystem.restype = ctypes.c_uint32
    dpi = GetDpiForSystem(0)  # 缩放100%时DPI为96，125%时为120，150%时为144，175%时为168， 200%时为192
    scale_factor = dpi / 96.0  # 将DPI转换为缩放因子
    # 计算逻辑分辨率
    logic_w = int(phys_w / scale_factor)
    logic_h = int(phys_h / scale_factor)
    # 判断是否可以运行
    if logic_w < 1280 or logic_h < 768:
        Available = False
    else:
        Available = True
    return deviceInfo(Available, logic_w, logic_h, phys_w, phys_h, scale_factor)


class StatusDisplay:
    def __init__(self, text, position=(0, 0)):
        self.root = tk.Tk()
        self.root.overrideredirect(True)  # 去掉窗口的标题栏和边框
        self.root.geometry(f"+{position[0]}+{position[1]}")  # 设置窗口位置
        self.root.attributes("-topmost", True)  # 窗口置顶
        self.root.attributes("-alpha", 0.9)  # 设置窗口透明度
        self.root.attributes("-transparentcolor", "green")  # 设置白色为透明色
        self.label = tk.Label(self.root, text=text, bg="white", fg="black", font=("黑体", 13))
        self.label.pack()
        self.root.update()  # 在主线程中运行 tkinter 的主循环

    def update(self, text, color='black'):
        self.label.config(text=text, fg=color)
        self.root.update()

    def move_to(self, position):
        self.root.geometry(f"+{position[0]}+{position[1]}")
        self.root.update()

    def panel_style(self, left, top, width, height):
        self.label.config(bg="white", fg="red", font=("黑体", 13), anchor="w", justify="left")
        self.label.pack(fill="both", expand=True)  # Label 填充整个窗口
        self.root.geometry(f"{width}x{height}+{left}+{top}")  # 设置窗口位置和尺寸

    def close(self):
        self.root.destroy()


class LoginInfoError(Exception):
    pass


class CheckImgError(Exception):
    pass


class WeChatError(Exception):
    pass


class GetCodeError(Exception):
    pass


class NetworkError(Exception):
    pass


class ContactNotFount(Exception):
    pass


class UnknownError(Exception):
    pass
