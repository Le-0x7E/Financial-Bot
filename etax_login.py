# from selenium.webdriver.edge.service import Service
# from pynput.keyboard import Controller
# from selenium import webdriver
# from logger import logger
from common import *
import pygetwindow as gw
import pandas as pd
import pyautogui
import pyperclip
# import random
# import time
import config
import sys


class WeChat:
    def __init__(self, display):
        logger.info('Initializing WeChat Window...')
        display.update(f"正在初始化微信窗口...")
        self.wechat_window = None
        self.init(display)

    def init(self, display):
        flag = False
        while True:
            try:
                pyautogui.hotkey('ctrl', 'alt', 'w')
                time.sleep(2)
                self.wechat_window = gw.getWindowsWithTitle("微信")  # 根据窗口标题匹配
                if self.wechat_window[0]:
                    self.wechat_window[0].moveTo(re_coord(553), re_coord(256))  # 窗口移动
                    time.sleep(0.5)
                    self.wechat_window[0].resizeTo(re_coord(710), re_coord(510))  # 设置窗口大小
                    break  # 如果成功退出，实际上就是已经初始化完成了，也可以用return
            except Exception as e:  # 暂时未能健康实现检测微信没启动，主要是打开关闭后可能检测不到，初次没打开的能检测到，待优化
                if flag:
                    logger.error('Wechat Window Still Not Found After Reattempt, Process Abort.\n')
                    display.update(f"未找到微信窗口！请确保微信已打开并且已登录！", 'red')
                    pop_error_window(f"请确保微信已打开并且已登录！\n请登录微信后点击确定，程序会再次检测！",
                                     title=f"{get_current_time()} 未找到微信窗口！")
                    sys.exit()  # 这里未来可以加上让用户打开微信后再检测的逻辑，或者放到总控制面板步骤（启动这个脚本前检测）
                logger.warning(f"Wechat Window Not Found : {e}, trying again...")
                self.open(display)  # 尝试打开微信窗口（有概率把前面打开的关闭了，待优化）
                flag = True
                time.sleep(2)

    @staticmethod
    def open(display):
        logger.info('Opening WeChat Window...')
        display.update(f"正在打开微信窗口...")
        pyautogui.hotkey('ctrl', 'alt', 'w')

    def close(self, display):
        logger.info('Closing WeChat Window...')
        display.update(f"正在关闭微信窗口...")
        self.wechat_window[0].minimize()  # 窗口最小化

    @staticmethod
    def send(msg, corp, display):
        # 搜索用户
        logger.debug(f'Searching Contact Of {corp.name}...')
        display.update(f"正在搜索用户...")
        pyautogui.moveTo(re_coord(666), re_coord(292), duration=random.uniform(0.2, 0.3))
        time.sleep(random.uniform(0.3, 0.8))
        pyautogui.click()
        time.sleep(random.uniform(0.5, 1.5))
        # pyperclip.copy(corp.name)
        pyperclip.copy('电局验证码')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(random.uniform(0.3, 0.8))

        # 检测是否存在该用户
        logger.debug(f'Checking Whether Contact Of {corp.name} Exist...')
        contact_box = (re_coord(610), re_coord(320), re_coord(52), re_coord(30))
        if not check_img('contact', contact_box):
            logger.error(f'Contact {corp.name} Not Found.')
            for i in range(5):
                display.update(f"未找到{corp.name}，请确认是否添加并备注好！{5 - i} 秒后开始下一公司！", 'red')
                time.sleep(1)
            raise ContactNotFount(f'Contact {corp.name} Not Found.')
            # pop_error_window(f"{corp.name}\n请确认是否添加并备注好该联系人！\n请点击确定，程序会自动终止！",
            #                  title=f"{get_current_time()} 未找到联系人！")
            # sys.exit()  # 未来需要记录报错后进行下一家公司

        # 进入聊天窗口
        logger.info(f'Entering Message Window Of {corp.name}')
        display.update(f"正在进入聊天窗口...")
        pyautogui.press('enter')
        time.sleep(random.uniform(0.5, 1.5))
        pyautogui.moveTo(re_coord(888), re_coord(688), duration=random.uniform(0.2, 0.3))
        time.sleep(random.uniform(0.3, 0.8))
        pyautogui.click()

        # 检查是否进入聊天窗口
        msg_window_box = (re_coord(1180), re_coord(626), re_coord(66), re_coord(34))
        logger.debug(f'Checking Weather Entered Message Window Of {corp.name}')
        if not check_img('msg_window', msg_window_box):
            return False  # 目前，若未进入聊天窗口直接返回，待优化
        time.sleep(random.uniform(0.5, 1.5))

        # 发送消息
        logger.debug(f'Sending Message To {corp.name}')
        display.update(f"正在发送信息...")
        pyperclip.copy(msg)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(random.uniform(0.05, 0.15))
        pyautogui.hotkey('ctrl', 'enter')
        time.sleep(random.uniform(0.05, 0.15))
        pyautogui.press('enter')
        logger.debug(f"Message “{msg}” Sent.")
        display.update(f"信息已发送")
        return True

    def get_code(self, corp, display):
        # 等待用户发送消息
        logger.info(f'Waiting Code From {corp.name}')
        waiting_code_box1 = (re_coord(1166), re_coord(588), re_coord(40), re_coord(36))
        waiting_code_box2 = (re_coord(898), re_coord(600), re_coord(36), re_coord(20))
        timer = 300
        while True:
            display.update(f"正在等待验证码...倒计时：0{int(timer / 60)}:{timer % 60}")
            if not check_img('waiting_code_p1', waiting_code_box1):
                if check_img('waiting_code_p2', waiting_code_box2):
                    break  # 没检测到自己消息右下角，并且检测到对方消息气泡底部时退出，这个逻辑不太好，尤其聊天记录清空后，待优化
            if timer == 0:
                # 获取验证码超时逻辑
                logger.error(f"Waiting Code From {corp.name} Timeout")
                for i in range(5):
                    display.update(f"获取验证码超时！{5 - i} 秒后开始下一公司！", 'red')
                    time.sleep(1)
                self.close(display)
                display.update('正在读取任务列表...')
                raise GetCodeError(f'Waiting Code From {corp.name} Timeout')
            time.sleep(1)
            timer -= 1

        # 用户发送消息，复制消息并提取验证码
        logger.debug(f'Received Message, Copying...')
        display.update(f"用户已发送消息，正在复制...")
        pyautogui.moveTo(re_coord(916), re_coord(590), duration=random.uniform(0.3, 0.6))
        time.sleep(random.uniform(0.3, 0.6))
        pyautogui.rightClick()
        time.sleep(random.uniform(0.3, 0.6))
        pyautogui.press('down')
        time.sleep(random.uniform(0.1, 0.3))
        pyautogui.press('enter')
        clipboard_content = pyperclip.paste()
        logger.debug(f'Copied Message"{clipboard_content}"From {corp.name} ')
        v_code = extract_code(clipboard_content)

        # 后处理
        time.sleep(random.uniform(0.5, 1.5))
        pyautogui.moveTo(re_coord(888), re_coord(688), duration=random.uniform(0.2, 0.3))
        time.sleep(random.uniform(0.3, 0.8))
        pyautogui.click()

        # 验证码有效
        if v_code is not False:
            logger.debug(f'Code {v_code} Got, Sending Confirm Message...')
            pyperclip.copy(f'已收到验证码 {v_code} ，感谢配合！')
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(random.uniform(0.05, 0.15))
            pyautogui.hotkey('ctrl', 'enter')
            time.sleep(random.uniform(0.05, 0.15))
            pyautogui.press('enter')
            display.update(f"验证码 {v_code} 已获取，正在关闭微信...")
            time.sleep(random.uniform(0.3, 0.5))
            self.close(display)

        # 验证码无效
        else:
            logger.debug(f'Can Not Extract Code From Message, Sending Wrong Message...')
            display.update(f"未检测到有效验证码")
            pyperclip.copy(f'未检测到有效验证码（6位数字），请勿发送其它内容，感谢配合！')
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(random.uniform(0.05, 0.15))
            pyautogui.hotkey('ctrl', 'enter')
            time.sleep(random.uniform(0.05, 0.15))
            pyautogui.press('enter')
        return v_code

    def code_verify(self, corp, pos, display):
        # 等待用户消息并获取验证码，等待5分钟，用户发的消息不规范的话最多尝试3次
        code = None
        for i in range(5):
            if i == 3:
                logger.error(f'Could Not Get Code From {corp.name} After 3 Attempts')
                for i in range(5):
                    display.update(f"无法从{corp.name}获取验证码！{5 - i} 秒后开始下一公司！", 'red')
                    time.sleep(1)
                self.close(display)
                display.update('正在读取任务列表...')
                raise GetCodeError(f'Could Not Get Code From {corp.name} After 3 Attempts')
            logger.debug(f'Start Getting Code From {corp.name}, This Is The {i} Time.')
            code = self.get_code(corp, display)
            if code is not False:
                break  # 获取到了验证码则退出循环

        # 已获取验证码并返回到了浏览器验证码界面, 输入验证码进行验证
        logger.info("Returned Browser Window.")
        display.update('已返回浏览器窗口')
        time.sleep(1.5)
        # 检测是否返回（待完成）

        # 点击文本框并输入验证码并点击登录
        logger.info('Inputting Verification Code...')
        display.update(f"正在输入验证码...")
        click(480, 392, pos)
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('del')
        pyautogui.typewrite(code)
        time.sleep(0.3)
        click(618, 486, pos)
        return None


def first_login_op(driver, pos, display):
    # 打开网页
    logger.info('Opening Page...')
    display.update(f"正在打开网页...")
    driver.get('https://etax.shaanxi.chinatax.gov.cn/')
    time.sleep(5)

    # 预处理
    logger.debug('First Login OP Finished, Starting Pre Work...')
    return pre_work(pos, display)


def pre_work(pos, display):
    # 重置浏览器缩放比例
    pyautogui.hotkey('ctrl', '0')
    time.sleep(2)

    # 检测是否进入网站主页：检查是否有登录按钮，若无则检测是否404，若404则尝试刷新
    login_box = (pos.left + re_coord(1148), pos.top + re_coord(10), pos.width + re_coord(75), pos.height + re_coord(36))
    if not check_img('login_button', login_box):
        check_404(pos, "Checking Home Page", display)
        if not check_img('login_button', login_box):
            raise CheckImgError(f'Could not find login button and seems not 404.')

    # 检测是否有网站通知弹窗
    if check_img('close_notification_button', pos):  # 全网页窗口检测
        display.update(f"检测到网站通知，正在关闭...")
        click(628, 518, pos)

    # 点击网站主页右上角的登录
    time.sleep(1)
    logger.debug('Clicking Login Button...')
    display.update(f"正在进入登录界面...")
    click(1185, 30, pos)
    time.sleep(5)


def input_login_info(corp, pos, display):
    # 检测是否有登录信息框，若404则尝试刷新
    logger.info('Checking Login Info Box...')
    login_info_box = (pos.left + re_coord(595), pos.top + re_coord(176),
                      pos.width + re_coord(425), pos.height + re_coord(288))
    if not check_img('login_info_box', login_info_box):
        check_404(pos, "Checking Login Page", display)
        if not check_img('login_info_box', login_info_box):
            raise CheckImgError(f'Could not find login info box and seems not 404.')

    # 点击文本框并输入纳税人识别号
    logger.info('Inputting Tax ID...')
    display.update(f"正在输入纳税人识别号...")
    click(655, 200, pos)
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'space')  # 切换到英文输入法
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('del')
    pyautogui.typewrite(corp.taxid)
    time.sleep(0.3)

    # 点击文本框并输入用户名
    logger.info('Inputting Phone Number...')
    display.update(f"正在输入手机号...")
    click(655, 260, pos)
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('del')
    pyautogui.typewrite(corp.loginid)
    time.sleep(0.3)

    # 点击文本框并输入密码
    logger.info('Inputting Password...')
    display.update(f"正在输入密码...")
    click(655, 320, pos)
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('del')
    pyautogui.typewrite(corp.taxpwd)
    time.sleep(0.3)

    # 完成滑动验证并点击登录
    logger.info('Finishing Slider...')
    display.update(f"正在完成滑动验证...")
    click(620, 380, pos)
    time.sleep(0.3)
    pyautogui.dragTo(pos.left + re_coord(random.randint(-3, 3) + 1050),
                     pos.top + re_coord(random.randint(-3, 3) + 388), duration=0.5)
    time.sleep(0.3)
    logger.info('Clicking Login Button...')
    click(805, 440, pos)  # 登录按钮

    # 检测是否存在登录信息错误
    time.sleep(2)
    logger.debug('Checking Whether Wrong Login Info...')
    wrong_login_info_box = (pos.left + re_coord(365), pos.top + re_coord(95),
                            pos.width + re_coord(510), pos.height + re_coord(100))
    if check_img('wrong_login_info', wrong_login_info_box):
        logger.error(f'Wrong Login Info of {corp.name}.')
        return False  # 当前为了演示，忽略登陆信息错误，实际上这里要return False
    return True


def finish_code_verify(corp, pos, display):
    for i in range(5):
        display.update(f"已进入验证码环节，准备发送验证码并打开微信！倒计时：{5 - i}")
        time.sleep(1)
    code_send_button_box = (pos.left + re_coord(710), pos.top + re_coord(366),
                            pos.width + re_coord(132), pos.height + re_coord(60))
    if not check_img('code_send_button', code_send_button_box):
        raise CheckImgError(f'Could not find code_send_button.')
    logger.debug('Clicking Verification Code Sending Button...')
    display.update('正在发送验证码...')
    click(772, 394, pos)
    display.update(f"已发送验证码，正在打开微信...")
    time.sleep(0.3)

    # 初始化微信窗口，设置微信窗口位置和大小，并检测是否成功
    wechat_box = (re_coord(1126), re_coord(246), re_coord(146), re_coord(40))
    wechat = WeChat(display)  # 初始化微信窗口
    time.sleep(1)
    logger.debug('Checking Whether Wechat Window Opened...')
    if not check_img('wechat_window', wechat_box):
        logger.debug('Wechat Window Not Found, Trying Init Again...')
        time.sleep(1)
        wechat.init(display)
        if not check_img('wechat_window', wechat_box):
            logger.error("Still could not find wechat window after trying again.")
            display.update(f"找不到微信窗口！", 'red')
            raise WeChatError(f'Still could not find wechat window after trying again.')

    # 给用户发提醒消息
    logger.debug(f'Preparing Sending Wechat Message To {corp.name}')
    display.update('准备发送信息')
    time.sleep(0.5)
    if not wechat.send(
            '您好，正在为您进行税务申报，登录需要验证码，请收到验证码后将验证码(6位数字)发送给我，请勿发送其它内容，感谢配合！',
            corp, display):
        wechat.init(display)
        if not wechat.send(
                '您好，正在为您进行税务申报，登录需要验证码，请收到验证码后将验证码(6位数字)发送给我，请勿发送其它内容，感谢配合！',
                corp, display):
            logger.error(f'Still could not send wechat message to {corp.name}')
            display.update(f"无法完成信息发送！", 'red')
            raise WeChatError(f'Still could not send wechat message to {corp.name}')

    # 进行验证码验证
    wechat.code_verify(corp, pos, display)
    # 判断验证码是否正确
    time.sleep(1.5)
    wrong_code_box = (pos.left + re_coord(406), pos.top + re_coord(102),
                      pos.width + re_coord(436), pos.height + re_coord(60))

    # 验证码错误处理逻辑，再尝试获取一次，若还错则抛异常
    if check_img('wrong_code', wrong_code_box):
        logger.warning('Wrong Code, Trying Again.')
        display.update('验证码错误！尝试再次获取！', 'red')
        time.sleep(0.3)
        # 打开微信窗口
        wechat.open(display)
        time.sleep(0.8)

        # 发送错误信息
        pyautogui.moveTo(re_coord(888), re_coord(688), duration=random.uniform(0.2, 0.3))
        time.sleep(random.uniform(0.3, 0.8))
        pyautogui.click()
        logger.debug('Wrong Code, Sending Wrong Message...')
        display.update("正在发送错误信息")
        pyperclip.copy('验证码错误！请仔细检查后重新发送，请勿发送其它内容，感谢配合！')
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(random.uniform(0.05, 0.15))
        pyautogui.hotkey('ctrl', 'enter')
        time.sleep(random.uniform(0.05, 0.15))
        pyautogui.press('enter')
        wechat.code_verify(corp, pos, display)
        time.sleep(1.5)
        if check_img('wrong_code', wrong_code_box):
            logger.error('Wrong Code Again')
            for i in range(5):
                display.update(f"验证码仍然错误！{5 - i} 秒后开始下一公司！", 'red')
                time.sleep(1)
            wechat.close(display)
            display.update('正在读取任务列表...')
            raise GetCodeError(f'Wrong Code From {corp.name} Again.')

    # 验证码正确，判断有无用户协议弹窗
    user_agreement_box = (pos.left + re_coord(380), pos.top + re_coord(96),
                          pos.width + re_coord(108), pos.height + re_coord(54))
    if check_img('user_agreement', user_agreement_box):
        time.sleep(0.3)
        logger.info('User Agreement Window Found, Closing.')
        display.update('检测到用户协议弹窗，正在确认...')
        click(818, 310, pos)

    # 检测有无身份类型选择弹窗
    identity_type_box = (re_coord(408), re_coord(322), re_coord(470), re_coord(260))
    if check_img('identity_type', identity_type_box):
        logger.info('Identity Type Selection Window Found, Selecting...')
        display.update('检测到身份类型选择弹窗，正在选择...')
        time.sleep(0.3)
        # 优先选办税员
        tax_collector = check_img('tax_collector', identity_type_box, pos=True)
        time.sleep(0.3)
        if tax_collector:
            logger.info('Tax Collector Button Detected, Clicking ...')
            display.update("选择办税员")
            x, y = box_center(tax_collector)
            pyautogui.click(x, y)
        else:
            # 没办税员则选财务负责人
            finance_chief = check_img('finance_chief', identity_type_box, pos=True)
            time.sleep(0.3)
            if finance_chief:
                logger.info('No Tax Collector Button But Finance Chief Button, Clicking...')
                display.update("选择财务负责人")
                x, y = box_center(finance_chief)
                pyautogui.click(x, y)
        # 点击确定
        time.sleep(0.3)
        click(818, 546)
        time.sleep(1)
        click(818, 490)

    # 检测是否成功进入登录页
    logger.info('Logging In...')
    display.update('正在登录中...')
    time.sleep(8)
    my_to_do_box = (pos.left + re_coord(508), pos.top + re_coord(76),
                    pos.width + re_coord(116), pos.height + re_coord(56))
    if not check_img('my_to_do', my_to_do_box):
        time.sleep(10)
        if not check_img('my_to_do', my_to_do_box):
            logger.info('Logging In Timeout(15s).')
            display.update('登录超时，正在尝试刷新...')
            pyautogui.press('f5')  # 有可能一刷新又回到登录信息页了，后续有时间可以优化
            time.sleep(5)
            check_404(pos, 'Checking Weather Login Success (Check My To Go Box).', display)
            display.update('无法检测到进入主页！', 'red')
            raise UnknownError('Could Not Find My To Do And Seems Not 404.')
    return None


def login(corp, pos, display):
    # 准备开始登录
    logger.info(f'Prepare Logging in {corp.name}')
    for i in range(5):
        display.update(f"准备登录 {corp.name} ！倒计时：{5 - i}")
        time.sleep(1)

    # 开始登录，输入登录信息，若登录信息错误则返回错误
    logger.info(f'Start Logging in Operation Of {corp.name}')
    if not input_login_info(corp, pos, display):
        for i in range(5):
            display.update(f"{corp.name} 的登录信息错误！{5 - i} 秒后开始下一公司！", 'red')
            time.sleep(1)
        raise LoginInfoError(f'Wrong Login Info of {corp.name}.')

    # 登录信息无误，进入验证码环节，准备点击发送验证码
    logger.info('Login Info Correct, Entering Verification Code Page...')
    finish_code_verify(corp, pos, display)
    logger.info('Login Success.')
    for i in range(5):
        display.update(f'已成功登录！准备开始税务申报！倒计时：{5 - i}')
    return None


def main(driver, corp, display, first=False, skip=False):
    # 设置网页窗口
    logger.info('Login Main Starting...')
    pos = position(left=re_coord(22), top=re_coord(136), width=re_coord(1256), height=re_coord(630))

    # 判断附加操作
    if first:
        first_login_op(driver, pos, display)
    elif skip:
        pass
    else:
        pre_work(pos, display)

    # 登录操作
    login(corp, pos, display)
    logger.debug('login.main Completed, Returning...')
    return None


if __name__ == '__main__':
    pop_error_window('此文件不可单独运行！！')
