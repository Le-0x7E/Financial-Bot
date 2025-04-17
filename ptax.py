from common import *
import config
import subprocess
import pyperclip


# ✅已完工，已测试
# 打开个税客户端，并根据打开结果返回对应状态码
def open_ptax_app(display):
    # 启动个税客户端
    try:
        # subprocess.run('./start_ptax.exe', shell=True)
        subprocess.Popen(config.ptax_app_path, shell=True)
        logger.info("open_ptax_app: Starting APP...")
        display.update('个税客户端正在启动，加载中...')
    except Exception as e:
        # 出现错误可以尝试再试一次
        logger.error(f"Error Occurred While Starting Ptax APP: {e}")
        return -1

    # 检查客户端启动状态
    logger.info("Checking Whether APP Started...")
    time.sleep(5)
    timer = 30
    for _ in range(timer):
        # 如果检查到初次启动界面则返回1
        first_setting_window = check_img('first_setting_window', config.screen_box, pos=True)
        if first_setting_window:
            logger.info("open_ptax_app: APP Not First Set, Returning 1...")
            return 1, first_setting_window

        # 如果检测到登录界面则返回2
        login_window = check_img('login_window', config.screen_box, pos=True)
        if login_window:
            logger.info("open_ptax_app: Login Window Detected, Returning 2...")
            return 2, login_window

        # 检测到安全弹窗则关闭
        secure_pop = check_img('secure_pop', config.screen_box)
        if secure_pop:
            logger.info(f"open_ptax_app: Secure Pop Detected, Closing...")
            display.update('检测到安全弹窗，正在关闭...')
            time.sleep(0.5)
            click(432, 196, box=secure_pop)
            time.sleep(1)
            click(552, 196, box=secure_pop)

        # 超时返回-1
        if timer == 0:
            logger.info("open_ptax_app:Timeout While Checking APP Started, Returning -1 ...")
            return -1, None
        timer -= 1
        time.sleep(1)


# 未完工，未测试，可能会弃用
# 初次打开客户端的设置，暂不继续开发，如果没有pre setting就提示用户自己pre setting
def finish_pre_setting(corp, pos, display):  # 未完工，可能会弃用
    logger.debug(f"finish_pre_setting: Function Entered, Corp: {corp.name}")
    # 选择省份
    logger.info('Selecting Province...')
    display.update('正在选择省份...')
    click(560, 262, box=pos)
    time.sleep(0.5)
    province_box = (
        pos.left + re_coord(410), pos.top + re_coord(244), pos.width + re_coord(202), pos.height + re_coord(34))
    for i in range(35):
        time.sleep(0.3)
        if check_img('shaanxi', province_box):
            pyautogui.press('enter')
            time.sleep(0.3)
            break
        pyautogui.press('down')

    # 填写纳税人识别号（表格中第一项）
    click(550, 302, box=pos)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'space')
    time.sleep(0.3)
    pyautogui.typewrite(corp.taxid)
    time.sleep(0.3)
    click(550, 346, box=pos)
    time.sleep(0.3)
    pyautogui.typewrite(corp.taxid)
    time.sleep(0.3)
    click(760, 500, box=pos)

    # 检查纳税人识别号是否正确（有时间再完善，可以跳过开始下一家）
    time.sleep(3)
    if check_img('corp_no_exist', pos):
        logger.error(f"Corp Info No Exist.")
        display.update('纳税人识别号错误！请在Excel中更正。')
    # 建议取消这个函数，如果没有pre setting就提示用户自己pre setting


# 未完工，未测试，可能会弃用
# 登录界面使用实名信息登录，暂不继续开发，暂时统一使用登录密码登录
def real_name_login(corp, pos, display):
    # 点击实名登录
    logger.info('Real Name Logging In...')
    display.update('正在进行实名登录...')
    click(106, 106, box=pos)

    # 输入登录信息
    click(202, 208, box=pos)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'space')
    time.sleep(0.3)
    pyautogui.typewrite(corp.loginid)
    time.sleep(0.3)
    click(202, 258, box=pos)
    time.sleep(0.3)
    pyautogui.typewrite(corp.taxpwd)
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.5)
    click(220, 380, box=pos)
    return None


# ✅已完工，已测试
# 登录界面，使用申报密码登录，登录Excel中第一个纳税人，进去后再决定是否需要切换
def declare_pwd_login(corp, pos, display):
    # 点击申报密码登录
    logger.info('Logging In The First Corp With Declare PWD...')
    display.update('正在进行申报密码登录...')
    click(326, 108, box=pos)
    time.sleep(0.5)

    # 选择公司
    logger.info('Selecting Corp...')
    display.update('正在选择纳税人...')
    click(200, 168, box=pos)
    time.sleep(0.5)
    pyautogui.press('del')
    time.sleep(0.3)
    pyperclip.copy(str(pd.read_excel(config.excel_path).iloc[2, 1]))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.3)
    pyautogui.press('space')
    time.sleep(0.5)
    click(180, 198, box=pos)
    time.sleep(1)

    # 输入登录密码
    logger.info('Inputting pwd...')
    display.update('正在输入密码...')
    click(202, 260, box=pos)
    time.sleep(0.3)
    pyautogui.typewrite(str(pd.read_excel(config.excel_path).iloc[2, 6]))
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.5)
    logger.info('Clicking Login Button...')
    click(220, 380, box=pos)
    return None


# ✅已完工，已测试，需要补充更多情况下的弹窗
# 检测弹窗是否有，若有则关闭
def close_pop_window(box, display):
    logger.info('Entered close_pop_window...')
    # 如果检测到加载中则等待
    while check_img("loading_pop", box):
        time.sleep(0.5)

    # 检测温馨提示
    friendly_reminder = check_img('friendly_reminder', box, pos=True)
    if friendly_reminder:
        logger.info('Friendly Reminder Found, Closing...')
        display.update('检测到温馨提示，正在关闭...')
        click(78, 426, box=friendly_reminder)
        time.sleep(0.3)
        click(342, 472, box=friendly_reminder)
        time.sleep(0.3)

    # 检测含确定按钮的弹窗(退付手续费核对提示、提示信息等)
    confirm_button = check_img('confirm_button', box, pos=True)
    if confirm_button:
        logger.info('Reminder Which Contains Confirm Button Found, Closing...')
        display.update('检测到弹窗，正在关闭...')
        x, y = box_center(confirm_button)
        pyautogui.click(x, y)

    logger.info('close_pop_window Returning...')
    return None


# ✅已完工，已测试，可优化
# 个税预处理工作，包括打开程序、使用申报密码登录、设置窗口位置、刷新左侧菜单栏
def pre_work(corp, display):
    # 打开个税客户端
    logger.info('Entered pre_work, Preparing To Open Ptax App...')
    open_result, pos = open_ptax_app(display)

    # 客户端启动成功并且进入了初次设置界面则弹窗提示后终止进程
    if open_result == 1:
        logger.info('Ptax APP No Pre Set, Programme Aborting...')
        display.update('检测到客户端未设置，请设置后再运行！', color='red')
        pop_error_window('\n检测到客户端未设置，请设置后再运行！\n点击确定将终止任务。', '客户端未设置')
        config.abort = True
        time.sleep(0.1)

    # 客户端启动出错或超时
    elif open_result == -1:
        logger.info('Ptax APP Starting Failed, Programme Aborting...')
        display.update('个税客户端启动失败！请检查个税客户端状态！程序即将退出。', color='red')
        pop_error_window('\n个税客户端启动失败！请检查个税客户端状态！\n点击确定将终止任务。', '客户端启动失败')
        config.abort = True
        time.sleep(0.1)

    # 客户端启动成功，进入登录界面，开始登录，登录后设置窗口位置，随后刷新左侧菜单栏
    elif open_result == 2:
        logger.info('Login Window Detected, Logging in...')
        display.update(f'已进入登录窗口，准备登录')
        time.sleep(1)

        # 使用申报密码登录Excel中第一个纳税人
        declare_pwd_login(corp, pos, display)
        logger.info('Logging In...')
        display.update('正在登录中...')
        time.sleep(2)

        # 检查登录信息是否错误？是否需要输入验证码？是否登录成功？(检测到欢迎页则视为登录成功)
        timer = 60
        for _ in range(timer):
            # 检测到登陆信息错误
            if check_img('wrong_login_info', config.screen_box):
                # 错误后会出一个图片验证码，影响下一次登录，非常麻烦
                logger.info(f'Wrong Login Info Of {corp.name}, task={config.task}')
                display.update('登录信息错误！！将跳过该公司！！下次登录时需手动输入验证码！！', color='red')
                time.sleep(5)
                logger.info("Raising SkipError, Reason: 该纳税人的个税登录密码错误。")
                raise SkipError("该纳税人的个税登录密码错误。")

            # 检测到需要验证码
            if check_img('v_code_needed', pos):
                # 提示用户手动输入验证码后点击弹窗，完成后继续检测是否进入主页
                logger.info(f'Needed To Finish Picture Verification Code, Popping Notice Window...')
                display.update('需要输入验证码，请手动输入后点击登录，完成后再点击弹窗的确定按钮！', color='red')
                pop_error_window('需要输入验证码！\n请手动输入验证码并点击登录按钮，\n完成后再点击本弹窗的确定按钮！', title='需要输入验证码')
                for i in range(10):
                    display.update(f'程序即将继续运行，倒计时: {10 - i}', color='red')
                    time.sleep(1)
                logger.info("User Clicked Confirm Button In The Pop Window, Programme Continue.")
                display.update('正在登录中...')

            # 检测到成功进入主页
            home_page_flag = check_img('home_page_flag', config.screen_box, pos=True)
            if home_page_flag:
                logger.info(f'Login Success, Entered Home Page, task={config.task}')
                display.update('登录成功，已进入主界面...')
                time.sleep(5)
                break
            timer -= 1
            time.sleep(1)

        # 进入主页后先检查并处理弹窗
        logger.info('Checking Pop Window...')
        close_pop_window(config.screen_box, display)
        time.sleep(0.5)

        # 处理弹窗后，若窗口最大化则取消最大化
        logger.info('Setting Ptax APP Window Position...')
        display.update('正在设置客户端窗口位置...')
        window_restore_button_box = (config.screen_box[2] - 151, 1, 150, 50)
        window_restore_button = check_img('window_restore_button', window_restore_button_box, pos=True)
        if window_restore_button:
            # 检测到窗口最大化
            logger.info(f'Window Seems Maximized, Restoring...')
            display.update('检测到窗口最大化，正在取消...')
            time.sleep(0.5)
            x, y = box_center(window_restore_button)
            pyautogui.click(x, y)
            time.sleep(0.5)
            logger.info('Setting Position...')
            set_window_position_size("自然人电子税务局（扣缴端）", 20, 20, 1280, 800)
            time.sleep(1)

            # 检查是否设置成功
            if check_img('window_restore_button', window_restore_button_box):
                # 如果没设置成功，检查下有没有弹窗，再试一次
                logger.info(f'Still Found window_restore_button, Seems Failed To Set Position.')
                time.sleep(1)
                close_pop_window(config.screen_box, display)
                time.sleep(0.5)
                # 未完成，先不写这个了
        else:
            # 若窗口非最大化，则直接设置窗口
            logger.info(f'Setting Window Position...')
            display.update('正在设置客户端窗口位置...')
            time.sleep(0.5)
            set_window_position_size("自然人电子税务局（扣缴端）", 20, 20, 1280, 800)
            time.sleep(1)

        # 设置成功后刷新左侧菜单栏
        logger.info('Position Set, Refreshing Left Menu...')
        display.update('正在刷新左侧菜单栏...')
        time.sleep(1)
        left_menu_box = (re_coord(22), re_coord(142), re_coord(190), re_coord(666))
        menu_left = check_img('menu_left', left_menu_box, pos=True)
        if menu_left:
            # 收起菜单
            logger.info('Clicking Menu Left Button...')
            time.sleep(0.5)
            x, y = box_center(menu_left)
            pyautogui.click(x, y)
            time.sleep(0.5)
            # 展开菜单
            menu_right = check_img('menu_right', left_menu_box, pos=True)
            logger.info('Clicking Menu Right Button...')
            time.sleep(0.5)
            x, y = box_center(menu_right)
            pyautogui.click(x, y)
            time.sleep(1)
            logger.info('pre_work Completed, Returning To main...')
            return None
        else:
            # 没找到收起按钮咋办？
            logger.error("Could Not Find Menu Left Button...")
            display.update('遇到未知异常，程序即将终止。', color='red')
            time.sleep(1)
            pop_error_window("遇到未知错误！程序即将终止！")
            raise UnknownError('Could Not Find Menu Left Button While Refreshing Left Menu...')


# ✅已完工，已测试，可优化
# 添加人员信息，目前只做了添加，后续可以在添加后执行：返回首页->所属税期改为1月->综合所得申报->正常工资薪金所得->更多操作->减除费用扣除确认
def add_person(corp, display):
    # 开始添加人员信息，进入人员信息采集
    logger.info("Adding Person, 正在进入人员信息采集...")
    display.update('正在进入人员信息采集...')
    window_box = (re_coord(20), re_coord(20), re_coord(1280), re_coord(800))
    left_menu_box = (re_coord(22), re_coord(142), re_coord(190), re_coord(666))
    RYXXCJ = check_img('RYXXCJ', left_menu_box, pos=True)
    logger.info('Clicking RYXXCJ Button...')
    x, y = box_center(RYXXCJ)
    time.sleep(0.5)
    pyautogui.click(x, y)   # 点击人员信息采集

    # 检测是否进入人员信息采集(检测添加按钮)
    logger.info('Checking Whether Entered RYXXCJ...')
    timer = 10
    person_add_button = check_img('person_add_button', window_box, pos=True)
    for _ in range(timer):
        if person_add_button:
            logger.info('Entered 人员信息采集.')
            display.update('已进入人员信息采集')
            time.sleep(2)
            break
        if timer == 0:
            logger.error('Timeout While Checking Whether Entered 人员信息采集, Aborting...')
            display.update('发生未知错误，程序即将终止！', color='red')
            time.sleep(1)
            pop_error_window("进入人员信息采集超时！\n可能出现未知错误！程序即将终止！")
            raise UnknownError('检测是否进入人员信息采集时超时，遇到未知异常。')
        timer -= 1
        time.sleep(1)
        person_add_button = check_img('person_add_button', window_box, pos=True)

    # 已进入人员信息采集，先检测并处理弹窗
    logger.info('Starting Checking Pop...')
    RYXXXZTS = check_img('pop_RYXXXZTS', window_box, pos=True)
    if RYXXXZTS:
        # 检测到人员信息下载提示弹窗，关闭
        logger.info('Found RYXXXZTS Pop, Closing...')
        display.update('检测到人员信息下载提示弹窗，正在关闭...')
        time.sleep(0.5)
        click(300, 166, box=RYXXXZTS)
        time.sleep(1)

    # 已关闭弹窗，开始添加人员信息
    logger.info('Starting Add Person, Clicking Add Button...')
    display.update('开始添加人员信息')
    x, y = box_center(person_add_button)
    time.sleep(0.5)
    pyautogui.click(x, y)   # 点击添加人员按钮
    time.sleep(1)
    pyautogui.click(re_coord(280), re_coord(334))   # 点击境内人员
    time.sleep(1)
    add_person_window = check_img('add_person_window', config.screen_box, pos=True)
    if not add_person_window:
        # 找不到添加人员的弹窗
        logger.error('Could Not Find add_person_window.')
        display.update('发生未知错误，程序即将终止！', color='red')
        time.sleep(1)
        pop_error_window("检测不到添加人员界面！\n可能出现未知错误！程序即将终止！")
        raise UnknownError('点击添加人员按钮后检测不到添加界面')

    # 添加人员的对话窗已打开，开始填写
    logger.info('add_person_window Detected, Filling Entries Of Person Info...')
    display.update('已进入添加人员信息填写界面，正在填写...')
    df = pd.read_excel(config.excel_path)   # 读取 Excel
    click(780, 178, box=add_person_window)   # 点击身份证
    time.sleep(0.5)
    pyautogui.typewrite(str(df.iloc[corp.serial + 1, 10]))
    time.sleep(0.5)
    click(310, 222, box=add_person_window)   # 点击姓名
    time.sleep(0.5)
    pyperclip.copy(str(df.iloc[corp.serial + 1, 9]))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    click(300, 624, box=add_person_window)   # 点击日期文本框
    time.sleep(0.5)
    pyautogui.typewrite('2025.1.1')   # 输入日期
    time.sleep(0.5)
    click(1000, 200, box=add_person_window)   # 点击拖动条
    time.sleep(0.2)
    current_x, current_y = pyautogui.position()
    pyautogui.dragTo(current_x, current_y + 400, duration=0.2)  # 下拖
    time.sleep(0.5)
    click(300, 276, box=add_person_window)   # 点击电话号码
    pyautogui.typewrite(str(df.iloc[corp.serial + 1, 11]))
    time.sleep(0.5)
    click(500, 660, box=add_person_window)   # 点击保存
    logger.info(' Entries Of Person Info Filling Completed.')
    time.sleep(2)

    # 检测是否有信息错误弹窗
    if check_img('add_person_wrong_info', box=add_person_window):
        logger.info('add_person_wrong_info Detected, Seems Person Info Is Not Correct, Processing...')
        display.update('人员信息异常，无法添加，将跳过该纳税人，正在关闭界面...', color='red')
        time.sleep(1)
        click(500, 410, box=add_person_window)   # 点击确定
        time.sleep(0.5)
        click(986, 22, box=add_person_window)   # 点击关闭按钮
        time.sleep(1)
        click(500, 410, box=add_person_window)   # 点击不保存
        time.sleep(3)
        logger.info('Raising SkipError To Skip This Corp...')
        raise SkipError('无可申报人员，且人员信息异常，无法添加。')

    # 若无信息错误弹窗，检测确定按钮并点击（报送成功弹窗）
    confirm_button = check_img('confirm_button', window_box, pos=True)
    if not confirm_button:
        # 没检测到则可能有问题
        logger.error('Could Not Find Person Added Success Pop, Seems UnknownError Occurred.')
        display.update('发生未知错误，程序即将终止！', color='red')
        time.sleep(1)
        pop_error_window("添加人员后未检测到成功或错误弹窗！\n可能出现未知错误！程序即将终止！")
        raise UnknownError('添加人员后检测不到成功或错误弹窗')
    logger.info('Confirm Button Found, Seems Added Success, Clicking...')
    display.update('正在关闭弹窗...')
    time.sleep(0.5)
    x, y = box_center(confirm_button)
    pyautogui.click(x, y)
    time.sleep(1)

    # 检测是否还是0条，若非0条则添加成功，否则抛异常
    zero_record_box = (re_coord(236), re_coord(740), re_coord(126), re_coord(50))
    if check_img('zero_record', zero_record_box):
        logger.error('Person Amount is Still Zero, Seems Add Person Failed.')
        display.update('发生未知错误，程序即将终止！', color='red')
        time.sleep(1)
        pop_error_window("添加人员后数量仍显示为 0 ！\n可能出现未知错误！程序即将终止！")
        raise UnknownError('添加人员后数量仍显示为 0')

    # 添加成功则点击报送，随后处理确认报送弹窗
    logger.info('Seems Person Added Success, Clicking Submit Button...')
    display.update('正在报送...')
    pyautogui.click(re_coord(436), re_coord(296))   # 点击报送按钮
    time.sleep(1)
    add_person_confirm = check_img('add_person_confirm', window_box, pos=True)
    if not add_person_confirm:
        # 若没检测到确认报送的弹窗，则可能有异常
        logger.error('Could Not Find add_person_confirm After Submit Button Clicked.')
        display.update('发生未知错误，程序即将终止！', color='red')
        time.sleep(1)
        pop_error_window("点击报送按钮后无确认报送弹窗\n可能出现未知错误！程序即将终止！")
        raise UnknownError('添加人员并点击报送按钮后无确认报送的弹窗')
    time.sleep(0.5)
    click(210, 150, add_person_confirm)   # 点击确认报送弹窗的确定

    # 检测并处理报送结果的弹窗
    timer = 10
    for i in range(timer):
        add_person_result = check_img('add_person_result', window_box, pos=True)
        if add_person_result:
            logger.info('add_person_result Detected, Seems Submit Success.')
            display.update('报送成功，准备返回并继续申报...')
            time.sleep(2)
            click(460, 330, add_person_result)   # 点击报送结果的关闭按钮
            time.sleep(1)
        if time == 0:
            logger.error('Could Not Find add_person_result After Submit.')
            display.update('发生未知错误，程序即将终止！', color='red')
            time.sleep(1)
            pop_error_window("报送后无报送结果弹窗\n可能出现未知错误！程序即将终止！")
            raise UnknownError('添加人员并点击报送按钮并确认报送后无报送结果的弹窗')
        timer -= 1
        time.sleep(1)

    # 添加完成，返回
    logger.info('Add Person Confirmed, Returning main.')
    return None


# ✅已完工，已测试，可优化
# 个税申报的逻辑，从主页进入综合所得申报，若无人员则返回-1，若无法申报则返回-2
def ptax_declare(corp, display):
    # 已进入主页，检测综合所得申报
    logger.info("Entered ptax_declare, Entering 综合所得申报...")
    display.update('正在进入综合所得申报...')
    window_box = (re_coord(20), re_coord(20), re_coord(1280), re_coord(800))
    left_menu_box = (re_coord(22), re_coord(142), re_coord(190), re_coord(666))
    top_item_box = (re_coord(220), re_coord(140), re_coord(1080), re_coord(70))
    ZHSDSB = check_img('ZHSDSB', left_menu_box, pos=True)
    if not ZHSDSB:
        # 若未检测到综合所得申报按钮，则抛异常
        logger.error('Could Not Find 综合所得申报 Entrance In Home Page, Seems UnknownError Occurred.')
        display.update('发生未知错误，程序即将终止！', color='red')
        time.sleep(1)
        pop_error_window("找不到综合所得申报入口！\n可能出现未知错误！程序即将终止！")
        raise UnknownError('Could Not Find 综合所得申报 Entrance In Home Page.')

    # 检测到综合所得申报按钮则点击，随后检测一次弹窗
    x, y = box_center(ZHSDSB)
    pyautogui.click(x, y)   # 点击综合所得申报按钮则
    time.sleep(2)
    logger.info("Checking Pop...")
    close_pop_window(window_box, display)

    # 检测是否进入综合所得申报（检测收入及减除填写）
    logger.info("Checking Whether Entered 综合所得申报...")
    timer = 10
    for i in range(timer):
        if check_img('income_and_deduction', top_item_box):
            logger.info('Entered 综合所得申报')
            display.update('已进入综合所得申报，准备开始申报')
            time.sleep(1)
            break
        # 超时则尝试关弹窗再试后报错
        if timer == 0:
            logger.info('Timeout While Entering 综合所得申报, Trying To Close Pop...')
            close_pop_window(window_box, display)
            time.sleep(1)
            if check_img('income_and_deduction', top_item_box):
                logger.info('Entered 综合所得申报')
                display.update('已进入综合所得申报，准备开始申报')
                time.sleep(1)
                break
            else:
                # 若仍未检测到则抛异常
                logger.error('Timeout While Confirming Whether Entered 综合所得申报, Seems UnknownError Occurred.')
                display.update('发生未知错误，程序即将终止！', color='red')
                time.sleep(1)
                pop_error_window("进入综合所得申报超时！\n可能出现未知错误！程序即将终止！")
                raise UnknownError('Timeout While Entering 综合所得申报')
        timer -= 1
        time.sleep(1)

    # 已进入综合所得申报，寻找并进入正常工资薪金所得，若检测到清除数据则说明存在填写的数据，则跳过该纳税人
    ZCGZXJSD = check_img('ZCGZXJSD', window_box, pos=True)
    x, y = box_center(ZCGZXJSD)
    clear_data_box = (re_coord(1150), ZCGZXJSD.top - 10, re_coord(130), ZCGZXJSD.height + 20)
    if check_img('clear_data', clear_data_box):
        # 检测到清除数据按钮则说明填写过数据
        logger.info('clear_data Detected, There is existing filled-in data, Skipping This Corp...')
        display.update('检测到存在已填写的数据，需要人工进行申报！即将跳过该纳税人！', color='red')
        time.sleep(3)
        raise SkipError("正常工资薪金所得存在已填写的数据，请人工判断并申报。")
    logger.info('Entering 正常工资薪金所得')
    display.update('正在进入正常工资薪金所得')
    time.sleep(0.5)
    pyautogui.click(re_coord(1168), y)   # 点击正常工资薪金所得的填写按钮

    # 检测是否进入自动导入向导的弹窗，若有则点击取消，不一定有
    timer = 10
    for i in range(timer):
        yes_no_pop = check_img('yes_no_pop', window_box, pos=True)
        if yes_no_pop:
            logger.info('Found 是否进入自动导入向导弹窗, Closing...')
            display.update('正在关闭弹窗')
            time.sleep(0.5)
            click(106, 20, box=yes_no_pop)   # 点击取消
            break
        timer -= 1
        time.sleep(1)

    # 开始导入申报数据，检测导入按钮
    logger.info('Starting Import Declare Data.')
    display.update('已进入正常工资薪金所得，开始导入申报数据')
    time.sleep(0.5)
    import_button = check_img('import_button', window_box, pos=True)
    if not import_button:
        # 找不到导入按钮，抛异常
        logger.error('Could Not Find Import Button In Declare Page.')
        display.update('发生未知错误，程序即将终止！', color='red')
        time.sleep(1)
        pop_error_window("找不到导入按钮！\n可能出现未知错误！程序即将终止！")
        raise UnknownError('Could Not Find Import Button In Declare Page.')

    # 检测到导入按钮则点击
    x, y = box_center(import_button)
    time.sleep(0.5)
    logger.info('Clicking Import Button...')
    pyautogui.click(x, y)  # 点击导入按钮
    time.sleep(1)
    pyautogui.click(re_coord(422), re_coord(256))  # 点击导入数据
    time.sleep(2)

    # 选择导入方式
    ZCGZXJSD_pop = check_img('ZCGZXJSD_pop', window_box, pos=True)
    ZCGZXJSD_pop_box = position(ZCGZXJSD_pop.left, ZCGZXJSD_pop.top, ZCGZXJSD_pop.width, re_coord(256))
    logger.info('Selecting Import Method...')
    if check_img('no_archives', ZCGZXJSD_pop_box):
        # 检测到无往期数据，生成零工资记录
        logger.info("No Archives Detected, Clicking 生成零工资记录...")
        display.update("检测到无往期申报数据，选择生成零工资记录...")
        time.sleep(0.5)
        click(160, 172, box=ZCGZXJSD_pop_box)   # 点击生成零工资记录
    else:
        # 若存在往期数据，则点击复制往期数据
        logger.info("Archives Detected, Clicking 复制往期数据...")
        display.update("存在往期申报数据，选择复制往期数据...")
        time.sleep(0.5)
        click(160, 74, box=ZCGZXJSD_pop_box)  # 点击复制往期数据

    # 点击导入方式弹窗中的导入按钮
    time.sleep(0.5)
    logger.info("Clicking Import Button...")
    click(272, 228, box=ZCGZXJSD_pop_box)   # 点击导入按钮
    time.sleep(2)

    # 检测是否存在无人员弹窗
    no_person_pop = check_img('no_person_pop', window_box, pos=True)
    if no_person_pop:
        # 检测到无人员弹窗，关闭后返回 -1
        logger.info("No Person Detected, Returning -1 To mian...")
        display.update("检测到无可申报人员，准备添加。")
        time.sleep(0.5)
        click(182, 150, box=no_person_pop)   # 点击无人员弹窗的确定按钮
        return -1

    # 检测导入零数据成功弹窗（确定按钮）
    timer = 30
    for i in range(timer):
        confirm_button = check_img('confirm_button', window_box, pos=True)
        if confirm_button:
            logger.info('Confirm Button Found, Seems Import Success, Clicking...')
            display.update('正在关闭确认弹窗...')
            time.sleep(0.5)
            x, y = box_center(confirm_button)
            pyautogui.click(x, y)
            time.sleep(1)
            break
        if timer == 0:
            # 检测导入零数据成功弹窗超时
            logger.error('Timeout While Checking Import Success Pop.')
            display.update('发生未知错误，程序即将终止！', color='red')
            time.sleep(1)
            pop_error_window("检测不到导入成功弹窗！\n可能出现未知错误！程序即将终止！")
            raise UnknownError('Timeout While Checking Import Success Pop.')
        timer -= 1
        time.sleep(1)

    # 导入零数据成功后返回综合所得申报，并依次点击税款计算、附表填写、申报表报送
    logger.info('Returning 综合所得申报, Clicking ZHSDSB...')
    display.update('成功导入数据，正在返回综合所得申报...')
    ZHSDSB = check_img('ZHSDSB', left_menu_box, pos=True)
    x, y = box_center(ZHSDSB)
    time.sleep(0.5)
    pyautogui.click(x, y)
    time.sleep(2)
    logger.info('Checking Pop...')
    close_pop_window(window_box, display)

    # 点击税款计算并检测是否成功计算
    logger.info('Entered 综合所得申报, Clicking 税款计算...')
    display.update('已进入综合所得申报， 正在进行税款计算...')
    time.sleep(1)
    pyautogui.click(re_coord(620), re_coord(178))   # 点击税款计算
    timer = 60
    re_calculate_box = (re_coord(180), re_coord(200), re_coord(230), re_coord(160))
    for i in range(timer):
        # 检测是否计算成功（检测重新计算按钮）
        if check_img('re_calculate', re_calculate_box):
            logger.info('税款计算成功')
            display.update('税款计算成功')
            time.sleep(0.5)
            break

        # 检测是否有税款计算弹窗，若有则跳过该纳税人
        tax_cal_pop = check_img('tax_calculation_pop', window_box, pos=True)
        if tax_cal_pop:
            logger.info('税款计算弹窗 ')
            logger.info('税款计算弹窗 Detected, Seems tax calculated Failed, Returning -2 To main...')
            display.update('检测到税款计算失败！需要人工申报！即将返回并开始下一任务！', color='red')
            time.sleep(3)
            tax_calculation_pop_box = position(tax_cal_pop.left, tax_cal_pop.top, tax_cal_pop.width, re_coord(160))
            click(460, 330, box=tax_calculation_pop_box)
            time.sleep(1)
            return -2   # 无法税款计算，需要人工申报，跳过

        # 检测税款计算结果超时
        if timer == 0:
            logger.error('Timeout While Checking tax_calculation Result.')
            display.update('发生未知错误，程序即将终止！', color='red')
            time.sleep(1)
            pop_error_window("等待税款计算结果超时！\n可能出现未知错误！程序即将终止！")
            raise UnknownError('Timeout While Checking tax_calculation Result.')
        timer -= 1
        time.sleep(1)

    # 税款计算完毕且无异常，点击附表填写
    logger.info('tax_calculation Completed, Clicking 附表填写...')
    display.update('税款计算完毕，正在进入附表填写...')
    time.sleep(1)
    pyautogui.click(re_coord(888), re_coord(178))   # 点击附表填写
    time.sleep(5)

    # 点击申报表报送
    logger.info('Entered 附表填写, Clicking 申报表报送...')
    display.update('已进入附表填写，正在进入申报表报送...')
    time.sleep(1)
    pyautogui.click(re_coord(1160), re_coord(178))  # 点击申报表报送
    timer = 36
    for i in range(timer):
        # 检测到导出申报表按钮则视为已进入申报表报送
        if check_img('export_declare_table', window_box):
            logger.info('Entered 申报表报送.')
            display.update('已进入申报表报送')
            time.sleep(1)
            break

        # 检测税款计算结果超时
        if timer == 0:
            logger.error('Timeout While Checking Whether Entered 申报表报送.')
            display.update('发生未知错误，程序即将终止！', color='red')
            time.sleep(1)
            pop_error_window("进入申报表报送超时！\n可能出现未知错误！程序即将终止！")
            raise UnknownError('Timeout While Checking Whether Entered 申报表报送.')
        timer -= 1
        time.sleep(1)

    # 已进入申报表报送，检测是否可申报，若不可申报则跳过，若可申报则发送申报
    logger.info('Entered 申报表报送, Checking Whether Declare Is Available')
    display.update('已进入申报表报送，正在检查是否可报送...')
    time.sleep(0.5)
    declare_yn_box = (re_coord(1190), re_coord(336), re_coord(90), re_coord(80))
    if check_img('could_not_declare', declare_yn_box):
        # 不可申报
        logger.info('否 Detected, Unable To Declare, Returning -2 To mian...')
        display.update('检测到不可申报！需要人工申报！即将开始下一任务！', color='red')
        time.sleep(1)
        return -2
    if check_img('could_declare', declare_yn_box):
        # 可以申报
        logger.info('是 Detected, Able To Declare, Clicking 发送申报...')
        display.update('检测到可申报，正在发送申报...')
        time.sleep(1)
        pyautogui.click(re_coord(1220), re_coord(274))   # 点击发送申报
        logger.info('Submitted, Waiting Result...')
        timer = 60
        for i in range(timer):
            # 检测到立即获取按钮则5秒后点击
            get_it_now_button = check_img('get_it_now_button', window_box, pos=True)
            if get_it_now_button:
                logger.info('立即获取按钮 Detected, Click After 5 seconds.')
                display.update('成功发送申报，5秒后获取申报结果')
                time.sleep(5)
                x, y = box_center(get_it_now_button)
                logger.info('Clinking 立即获取按钮...')
                display.update('立即获取申报结果')
                pyautogui.click(x, y)
                time.sleep(1)
                continue

            # 点击申报成功弹窗则点击确定
            declare_success_pop = check_img('declare_success_pop', window_box, pos=True)
            if declare_success_pop:
                logger.info('申报成功弹窗 Detected, Declare Success, Clicking Confirm Button...')
                display.update('申报成功，正在关闭弹窗...')
                time.sleep(1)
                click(150, 150, box=declare_success_pop)
                time.sleep(1)
                break

            # 检测申报结果超时
            if timer == 0:
                logger.error('Timeout While Checking Declare Result.')
                display.update('发生未知错误，程序即将终止！', color='red')
                time.sleep(1)
                pop_error_window("等待申报结果超时！\n可能出现未知错误！程序即将终止！")
                raise UnknownError('Timeout While Checking Declare Result.')
            timer -= 1
            time.sleep(1)

        # 截图申报成功界面
        logger.info('Declare Success, Creating Screenshot...')
        display.update('申报成功，正在截图...')
        time.sleep(0.5)
        display.update(f'个税申报成功截图', color='red')
        make_screenshot(window_box, "个税申报成功截图", config.screenshot_dir)
        display.update('截图完毕，准备查询申报记录...')

        # 申报记录查询并截图
        time.sleep(1)
        ps_declare_record(display)
        time.sleep(0.5)
        logger.info('ptax_declare: Returning...')
        return True


# ✅已完工，已测试，可优化
# 申报记录查询并截图
def ps_declare_record(display):
    logger.info("Entered ps_declare_record, Starting Query Declaration Records...")
    display.update('开始查询申报记录，正在进入单位申报记录查询...')
    window_box = (re_coord(20), re_coord(20), re_coord(1280), re_coord(800))
    left_menu_box = (re_coord(22), re_coord(142), re_coord(190), re_coord(666))

    # 处理单位申报记录查询按钮
    DWSBJLCX = check_img('DWSBJLCX', left_menu_box, pos=True)
    if not DWSBJLCX:
        logger.info('Could Not Find 单位申报记录查询, Trying To Click 查询统计...')
        query_statistics = check_img('query_statistics', left_menu_box, pos=True)
        if not query_statistics:
            logger.error('Could Not Find query_statistics In left_menu_box.')
            display.update('出现未知错误，程序即将终止', color='red')
            time.sleep(1)
            pop_error_window("找不到查询统计入口！\n可能出现未知错误！程序即将终止！")
            raise UnknownError("Could Not Find query_statistics In left_menu_box.")
        logger.info('Found 查询统计, Clicking...')
        x, y = box_center(query_statistics)
        time.sleep(0.5)
        pyautogui.click(x, y)  # 点击查询统计
        time.sleep(2)
        DWSBJLCX = check_img('DWSBJLCX', left_menu_box, pos=True)
        if not DWSBJLCX:
            logger.error('Still Could Not Find 单位申报记录查询 After Clicking 查询统计.')
            display.update('出现未知错误，程序即将终止', color='red')
            time.sleep(1)
            pop_error_window("找不到查询统计入口！\n可能出现未知错误！程序即将终止！")
            raise UnknownError("Still Could Not Find 单位申报记录查询 After Clicking 查询统计.")

    # 进入单位申报记录查询
    logger.info('Found 单位申报记录查询, Clicking...')
    x, y = box_center(DWSBJLCX)
    time.sleep(0.5)
    pyautogui.click(x, y)   # 点击单位申报记录查询
    time.sleep(0.5)

    # 检测是否进入单位申报记录查询(检测查询按钮)
    timer = 20
    query_button_box = (re_coord(888), re_coord(130), re_coord(130), re_coord(100))
    for i in range(timer):
        query_button = check_img('query_button', query_button_box, pos=True)
        if query_button:
            logger.info('query_button Detected, Entered 单位申报记录查询.')
            display.update('已进入单位申报记录查询，准备开始查询')
            time.sleep(2)
            break
        if timer == 0:
            logger.error('Timeout While Entering 单位申报记录查询')
            display.update('出现未知错误，程序即将终止', color='red')
            time.sleep(1)
            pop_error_window("进入单位申报记录查询超时！\n可能出现未知错误！程序即将终止！")
            raise UnknownError('Timeout While Entering 单位申报记录查询')
        timer -= 1
        time.sleep(1)

    # 已进入单位申报记录查询，开始查询（暂时不管月份了）
    logger.info('Starting Query Declaration Records...')
    display.update('正在查询...')
    pyautogui.click(re_coord(390), re_coord(180))   # 点击起始日期文本框
    time.sleep(0.5)
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.press('del')
    # time.sleep(0.5)
    # pyautogui.typewrite(f"{get_current_date(y=True)}.{int(get_current_date(m=True) )- 1}")   # 填写起始月份，暂时不管月份了
    time.sleep(0.5)
    pyautogui.click(re_coord(960), re_coord(176))   # 点击查询按钮
    time.sleep(5)

    # 等出来结果后截图
    logger.info('Starting Creating Screenshot...')
    display.update('正在截图...')
    time.sleep(0.5)
    display.update(f'个税申报记录查询截图', color='red')
    make_screenshot(box=window_box, name='个税申报记录查询截图', path=config.screenshot_dir)
    display.update('截图完毕')
    time.sleep(1)

    # 检测若有无结果弹窗则点击确定
    no_record_pop = check_img('no_record_pop', window_box)
    if no_record_pop:
        logger.info('no_record_pop Detected, Seems No Declare Record Exist, Clicking Confirm...')
        display.update('关闭弹窗')
        time.sleep(0.5)
        click(156, 150, box=no_record_pop)
        time.sleep(0.5)
    logger.info('ps_declare_record: Function Completed, Returning...')
    return True


# ✅已完工，已测试，可优化
# 切换纳税人，若检测到已申报则查询申报记录并截图
def switch_corp(corp, display):
    logger.info('Switching Corp...')
    display.update('正在切换企业...')
    window_box = (re_coord(20), re_coord(20), re_coord(1280), re_coord(800))
    corp_management = check_img('corp_management', window_box, pos=True)
    x, y = box_center(corp_management)

    # 点击单位管理,打开单位管理界面
    time.sleep(1)
    pyautogui.click(x, y)
    time.sleep(2)
    corpmgr_pop = check_img('corpmgr_pop', config.screen_box, pos=True)
    if not corpmgr_pop:
        # 没有则报错
        logger.error('Could Not Find corpmgr_pop After Clicking corp_management.')
        display.update('出现未知错误，程序即将终止', color='red')
        time.sleep(1)
        pop_error_window("无法打开企业管理！\n可能出现未知错误！程序即将终止！")
        raise UnknownError('Could Not Find corpmgr_pop After Clicking corp_management.')
    corpmgr_pop_box = position(corpmgr_pop.left, corpmgr_pop.top, corpmgr_pop.width, re_coord(442))

    # 搜索纳税人
    click(160, 70, corpmgr_pop_box)   # 点击taxid文本框
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.3)
    pyautogui.press('del')
    time.sleep(0.3)
    pyautogui.typewrite(corp.taxid)   # 输入taxid
    time.sleep(0.3)
    click(490, 70, corpmgr_pop_box)   # 点击查询按钮
    time.sleep(1)
    corp_enter = check_img('corp_enter', corpmgr_pop_box, pos=True)

    # 若搜索不到纳税人则添加
    if not corp_enter:
        # 搜索无结果则添加
        click(50, 126, corpmgr_pop_box)   # 点击添加按钮
        time.sleep(1)
        add_corp = check_img('add_corp', config.screen_box, pos=True)
        click(280, 120, add_corp)   # 点击纳税人识别号文本框
        time.sleep(0.5)
        pyautogui.typewrite(corp.taxid)   # 输入taxid
        time.sleep(0.5)
        click(280, 166, add_corp)   # 点击确认纳税人识别号文本框
        time.sleep(0.5)
        pyautogui.typewrite(corp.taxid)   # 输入taxid
        time.sleep(0.5)
        click(210, 216, add_corp)   # 点击保存按钮
        time.sleep(3)
        save_success_pop = check_img('save_success_pop', config.screen_box, pos=True)
        click(150, 150, save_success_pop)   # 点击确定
        time.sleep(1)

    # 搜索有结果后则判断是否已申报，若已申报则设置标志，登录后去申报记录查询截图
    declared_flag = False
    if check_img('declared', corpmgr_pop_box):
        logger.info(f"{corp.name} Declared, Skipping...")
        display.update(f"检测到{corp.name}已申报，将登录后查询申报记录!", color='red')
        time.sleep(1)
        declared_flag = True

    # 点击进入，打开登录界面
    corp_enter = check_img('corp_enter', corpmgr_pop_box, pos=True)
    x, y = box_center(corp_enter)
    time.sleep(0.5)
    pyautogui.click(x, y)  # 点击进入
    time.sleep(2)
    switch_corp_login = check_img('switch_corp_login', config.screen_box, pos=True)
    if not switch_corp_login:
        # 没有则说明可能已经在该企业了，返回
        logger.info("Could Not Find switch_corp_login, Seems Already Logged In.")
        return None

    # 输入密码并登录
    click(230, 32, switch_corp_login)   # 点击输入密码文本框
    time.sleep(0.5)
    pyautogui.typewrite(corp.ptaxpwd)  # 输入密码
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    click(282, 200, switch_corp_login)
    time.sleep(5)

    # 检测到需要验证码
    if check_img('v_code_needed', window_box):
        # 提示用户手动输入验证码后点击弹窗，随后检测用户是否点了登录，如果没点则点登录，点了则继续检测是否进入主页
        display.update('需要输入验证码，请手动输入后再点击弹窗的确定按钮！', color='red')
        pop_error_window('需要输入验证码\n请手动输入验证码\n随后再点击本弹窗的确定按钮', title='需要输入验证码')
        for i in range(5):
            display.update(f'程序即将继续运行，倒计时: {5 - i}', color='red')
            time.sleep(1)

    # 检测到登录信息错误
    if check_img('wrong_login_info', config.screen_box):
        logger.info(f'Wrong Login Info Of {corp.name}, task={config.task}')
        display.update('登录信息错误！！', color='red')
        # 处理逻辑待完善，错误后会出一个图片验证码，影响下一次登录，麻烦

    # 检测到成功进入主页则检查弹窗后返回
    home_page_flag = check_img('home_page_flag', config.screen_box, pos=True)
    if home_page_flag:
        logger.info(f'Login Success Of {corp.name}, Entered Home Page, task={config.task}')
        display.update('登录成功，已进入主界面，等待加载中...')
        time.sleep(5)
        close_pop_window(window_box, display)
        time.sleep(0.5)

    # 处理已申报的企业（查询记录并截图）
    if declared_flag:
        logger.info(f'Processing declared corp...')
        time.sleep(0.5)
        ps_declare_record(display)
        logger.info(f'Declared Processing Completed, Raising SkipError...')
        time.sleep(0.5)
        raise SkipError('检测到该纳税人状态为已申报，截图申报记录后跳过。')

    logger.info("switch_corp Returning...")
    return None


# ✅已完工，已测试
# 主函数
def main(corp, display, first=False):
    logger.info('Entered ptax.main')
    # 若初次执行，则完成预先工作
    if first:
        logger.info('First Login, Starting Pre-Work...')
        time.sleep(0.5)
        pre_work(corp, display)
        logger.info('Pre-Work Complete.')

        # 若当前任务的纳税人名称非Excel中第一个，则切换为当前任务的账号。
        if corp.name != str(pd.read_excel(config.excel_path).iloc[2, 1]):
            logger.info('corp.name Of Current Task != The First One In Excel, Switching To Current Corp...')
            display.update('正在切换为当前任务账号...')
            time.sleep(0.5)
            switch_corp(corp, display)
            logger.info('Switching Corp Completed...')

    # 若非初次执行，则切换至对应企业
    else:
        logger.info('Not First Login, Switching Corp...')
        time.sleep(0.5)
        switch_corp(corp, display)
        logger.info('Switching Corp Completed...')

    # 登录完成，开始进行个税申报
    logger.info('Entered Home Page, Starting p-tax declare...')
    display.update('开始进行个税申报')
    time.sleep(1)
    declare_result = ptax_declare(corp, display)

    # 申报返回 -1 则为无人员，添加后重试
    if declare_result == -1:
        logger.info('No Person, Need To Add Person, Starting Adding...')
        display.update('无可申报人员，正在添加...')
        time.sleep(0.5)
        add_person(corp, display)
        logger.info('Person Added, Restarting ptax_declare...')
        display.update('添加人员完毕，重新开始申报')
        time.sleep(0.5)
        declare_result = ptax_declare(corp, display)
        if declare_result == -1:
            # 若还有此问题则跳过
            logger.info('Still No Person After Added, Skipping This Corp...')
            display.update('仍然无可申报人员，即将跳过该纳税人...')
            time.sleep(1)
            raise SkipError('无可申报人员，尝试添加后仍然无可申报人员，请人工检查。')

    # 返回 -2 为无法申报，需要人工申报，跳过
    if declare_result == -2:
        logger.info('Could Not Declare, Skipping This Corp...')
        display.update('该纳税人无法申报，需要人工申报，即将跳过。', color='red')
        time.sleep(1)
        raise SkipError('该纳税人无法申报，请人工检查并申报。')

    # 未返回错误码则视为申报成功，正常返回
    logger.info('Declare Completed, Returning taskmgr...')
    display.update('该纳税人申报成功！即将开始后续任务！')
    time.sleep(5)
    logger.info('ptax.main Returning...')
    return None


if __name__ == '__main__':
    pop_error_window('此文件不可单独运行！！')
