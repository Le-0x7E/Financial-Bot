from common import *
# import datetime
import config


# ✅已完工，已测试，可优化
# 从主页进入全量发票查询，如果可以直接通过URL进的话可能就不需要这个流程了，也可以做为可选设置
def enter_full_invoice_inquiry(driver, display, pos):
    # 进入税务数字账户
    logger.info('Entering Tax Digital Account...')
    display.update('正在进入税务数字账户...')
    hot_service_box = (re_coord(96), re_coord(669), re_coord(1090), re_coord(86))  # 热门服务box
    tax_digit = check_img('tax_digit', hot_service_box, pos=True)  # 按钮位置不固定，识图点击
    time.sleep(0.5)

    # 如果第一页找到了则点击
    if tax_digit:
        time.sleep(0.5)
        x, y = box_center(tax_digit)
        logger.info('Found Tax Digital Account Button, Clicking...')
        pyautogui.click(x, y)
        time.sleep(2)

    # 如果第一页找不到则翻页再找
    else:
        logger.info('Could Not Find Tax Digital Account Button On Current Page, Clicking Next Page...')
        click(1202, 736)
        time.sleep(2)
        tax_digit = check_img('tax_digit', hot_service_box, pos=True)
        time.sleep(0.5)
        if tax_digit:
            # 翻页后找到则点击
            time.sleep(0.5)
            x, y = box_center(tax_digit)
            logger.info('Found Tax Digital Account Button, Clicking...')
            pyautogui.click(x, y)
            time.sleep(2)
        else:
            # 若翻页后还没有(几乎不可能)的处理逻辑待完善
            logger.info('Still Could Not Find Tax Digital Account Button.')
            display.update('找不到税务数字账户入口，出现未知异常，程序将会终止')
            raise UnknownError('Still Could Not Find Tax Digital Account Button.')

    # 判断是否进入税务数字账户（检测发票业务按钮）,若进入则点击发票业务
    timer = 60
    invoice_business_button_box = (re_coord(118), re_coord(416), re_coord(1062), re_coord(266))
    for i in range(timer):
        # 若检测到发票业务则点击后跳出循环
        if check_img('invoice_business_button', invoice_business_button_box):
            logger.info('Found Invoice Business Button, Clicking...')
            display.update('正在进入发票业务...')
            time.sleep(0.5)
            invoice_business_button = check_img('invoice_business_button', invoice_business_button_box, pos=True)
            x, y = box_center(invoice_business_button)
            pyautogui.click(x, y)
            time.sleep(1)
            break

        # 检查404
        check_404(pos, 'Checking Whether Entered Tax Digital Account.', display)
        time.sleep(1)
        timer -= 1

        # 30秒后刷新一次
        if timer == 30:
            logger.info('Waited 30s, Refreshing...')
            pyautogui.press('f5')
            time.sleep(1)

        # 超时处理逻辑
        if timer == 0:
            logger.error('Could Not Confirm Whether Entered Tax Digital Account And Seems Not 404.')
            raise TimeOutError('Could Not Confirm Whether Entered Tax Digital Account And Seems Not 404.')

    # 判断是否进入发票业务（最多等60秒），进入后若无发票查询统计入口则向下翻页一次
    timer = 60
    invoice_business_box = (re_coord(26), re_coord(202), re_coord(126), re_coord(60))
    unable_access_td_box = (re_coord(442), re_coord(142), re_coord(422), re_coord(120))
    for i in range(timer):
        # 检测到发票业务按钮
        if check_img('invoice_business', invoice_business_box):
            logger.info('Entered Invoice Business,.')
            display.update('已进入发票业务')
            time.sleep(1)
            click(1060, 310)
            time.sleep(1)
            if not check_img('invoice_inquiry', pos):
                logger.info('Entered Invoice Business, Pressing PageDown...')
                display.update('已进入发票业务，正在向下翻页...')
                pyautogui.press('pagedown')  # 向下翻页
            time.sleep(1)
            break

        # 检查404
        check_404(pos, 'Checking Whether Entered Invoice Business.', display)
        timer -= 1
        time.sleep(1)

        # 检查有无"不是发票试点纳税人"的弹窗，待完善
        if check_img('unable_access_td', unable_access_td_box):
            # 如果检查到了的处理逻辑，统一taskmgr写表格，这里需要关标签页并退出
            logger.error("'无权操作税务数字账户' Detected, Processing...")
            display.update('检测到无权操作税务数字账户提示！请确保IP及账号无异常！即将退出登录！', color='red')
            time.sleep(3)
            close_tab_return_home(driver, display)
            time.sleep(1)
            logout()
            raise SkipError('检测到无权操作税务数字账户。')

        # 30秒后刷新一次
        if timer == 30:
            logger.info('Waited 30s, Refreshing...')
            pyautogui.press('f5')
            time.sleep(1)

        # 超时处理逻辑
        if timer == 0:
            logger.error('Could Not Confirm Whether Entered Tax Digital Account And Seems Not 404.')
            raise TimeOutError('Could Not Confirm Whether Entered Tax Digital Account And Seems Not 404.')

    # 已进入发票业务，点击发票查询统计
    time.sleep(0.5)
    invoice_inquiry = check_img('invoice_inquiry', pos, pos=True)  # 全网页找按钮
    # 找到按钮则点击
    if invoice_inquiry:
        logger.info('Clicking Invoice Inquiry Statistics Button...')
        display.update('正在进入发票查询统计...')
        time.sleep(1)
        x, y = box_center(invoice_inquiry)
        pyautogui.click(x, y)
        time.sleep(1)
    # 未找到则抛异常（后续可以刷新再试一次或者回主页再试一次）
    else:
        logger.error("Could Not Find Invoice Inquiry Statistics Button After Page Down.")
        raise RetryError("Could Not Find Invoice Inquiry Statistics Button After Page Down.")  # 错误处理逻辑待完善

    # 已进入发票查询统计，点击进入全量发票查询
    timer = 60
    for i in range(timer):
        full_invoice_inquiry_button = check_img('full_invoice_inquiry_button', pos, pos=True)  # 全网页找按钮
        if full_invoice_inquiry_button:
            logger.info('Clicking Full Invoice Inquiry Button...')
            display.update('正在进入全量发票查询...')
            time.sleep(1)
            x, y = box_center(full_invoice_inquiry_button)
            pyautogui.click(x, y)
            time.sleep(1)
            break

        # 检查404
        check_404(pos, 'Checking Whether Entered Invoice Business.', display)
        timer -= 1
        time.sleep(1)

        # 30秒后刷新一次
        if timer == 30:
            logger.info('Waited 30s, Refreshing...')
            pyautogui.press('f5')
            time.sleep(1)

        # 超时处理逻辑
        if timer == 0:
            logger.error("Could Not Find Invoice Inquiry Statistics Button After Page Down.")
            raise TimeOutError("Could Not Find Invoice Inquiry Statistics Button After Page Down.")
    return None


# ✅已完工，已测试，可优化
# 判断当前纳税人是否零申报，通过全量发票查询，一般纳税人查月度的开具和取得，小规模查季度的开具
def judge_zero(driver, corp, display, pos):
    # 进入发票全量查询
    logger.info('Starting Judge Zero, Entering Full Invoice Inquiry Page...')
    display.update('开始判断是否零申报，正在进入发票全量查询...')

    # 直接URL打开可能会不可行？会不会能打开页面但是查不了？并且加载很慢，如果不行就用上面的函数进入
    # driver.execute_script(
    #     "window.open('https://dppt.shaanxi.chinatax.gov.cn:8443/invoice-query/invoice-query', '_blank');")
    enter_full_invoice_inquiry(driver, display, pos)
    time.sleep(1)

    # 判断是否进入发票全量查询
    full_invoice_inquiry_box = (re_coord(66), re_coord(200), re_coord(160), re_coord(60))
    timer = 60
    for i in range(timer):
        # 检查发票全量查询图标
        if check_img('full_invoice_inquiry', full_invoice_inquiry_box):
            logger.info('Entered Full Invoice Inquiry...')
            display.update('已进入发票全量查询')
            time.sleep(1)
            break
        # 检查是否404
        check_404(pos, 'Checking Whether Entered Full Invoice Inquiry', display)
        timer -= 1
        time.sleep(1)
        # 30秒后刷新一次
        if timer == 30:
            logger.info('Waited 30s, Refreshing...')
            pyautogui.press('f5')
            time.sleep(1)
        # 超时处理逻辑，抛异常给Panel，重试一次
        if timer == 0:
            logger.error('Waited 40s, Could Not Confirm Whether Entered Full Invoice Inquiry And Seems Not 404.')
            raise TimeOutError('Could Not Confirm Whether Entered Full Invoice Inquiry And Seems Not 404.')

    # 进入发票全量查询页面后，选择查询日期，若失败则等5秒后再试一次
    if not select_date(corp, display):
        # 找不到当前月和找不到月份减时才会返回False，等5秒后再试一次，还不行就报错
        logger.warning("select_date Error, Could Not Find Month Button, Retry After 5s")
        time.sleep(5)
        if not select_date(corp, display):
            # 重试还不行，关标签页回主页，抛异常返回Panel，重新开始一次declare.main
            logger.error("select_date Error, Still Could Not Find Month Button During Retry")
            display.update('无法选择日期，正在返回主页，随后将重试一次', color='red')
            time.sleep(1)
            close_tab_return_home(driver, display)
            time.sleep(2)
            raise RetryError('select_date: Error, Still Could Not Find Month Button During Retry.')

    # 选择完日期，查询并检查开具发票
    time.sleep(0.5)
    logger.info('Date Selected, Querying And Checking Issued Invoices...')
    if inquiry_and_judge(corp, display):
        # 存在开具发票记录则非零申报
        logger.info(f'Issued Invoices Checked, {corp.name} Is Not Zero.')
        display.update("检测到存在开具发票记录，非零申报！", color='red')
        # 若设置要截图则截图
        if config.config["invoice_inquiry_ps"]:
            logger.info(f'Creating Screenshot...')
            display.update("检测到存在开具发票记录，非零申报，正在截图！", color='red')
            zoom_and_ps(driver, '电局开具发票记录')
        for i in range(5):
            display.update(f"检测到存在发票记录，非零申报，请手动做账并申报！{5 - i}秒后返回", color='red')
        logger.info(f"judge_zero: {corp.name}Is Not Zero Due To Having Issued Invoices, Returning False...")
        return False  # 非零申报返回False
    else:
        # 无开具发票记录
        time.sleep(0.5)
        logger.info(f'{corp.name} Has No Issued Invoices.')
        display.update("无开具发票记录！")
        # 若设置要截图则截图
        if config.config["invoice_inquiry_ps"]:
            logger.info(f'Creating Screenshot...')
            display.update("无开具发票记录，正在截图！")
            zoom_and_ps(driver, '电局开具发票记录')
        # 无开具发票记录，一般纳税人还需要检查取得发票记录（Word里说小规模不用）
        if corp.type == '一般纳税人':
            time.sleep(0.5)
            logger.info(f'{corp.name} Is {corp.type}, Checking Obtained Invoices...')
            display.update(f"{corp.name}是{corp.type}，正在查询取得发票记录")
            # 选择取得发票
            logger.info(f'Selecting Obtained Invoices...')
            click(280, 422)
            time.sleep(0.5)
            click(280, 500)
            time.sleep(0.5)
            # 查询并判断取得发票
            logger.info('Selected, Querying And Checking Obtained Invoices...')
            if inquiry_and_judge(corp, display):
                # 存在取得发票记录处理逻辑
                logger.info(f'Obtained Invoices Checked, {corp.name} Is Not Zero, Creating Screenshot...')
                display.update("检测到存在取得发票记录，非零申报！", color='red')
                # 若设置要截图则截图
                if config.config["invoice_inquiry_ps"]:
                    logger.info(f'Creating Screenshot...')
                    display.update("检测到存在取得发票记录，非零申报，正在截图！", color='red')
                    zoom_and_ps(driver, '电局取得发票记录')
                for i in range(5):
                    display.update(f"检测到存在取得发票记录，非零申报，请手动做账并申报！{5 - i}秒后返回", color='red')
                logger.info(f"judge_zero: {corp.name}Is Not Zero Due To Having Obtained Invoices, Returning False...")
                return False  # 非零申报返回False
            else:
                # 无取得发票记录
                time.sleep(0.5)
                logger.info(f'{corp.name} Has No Obtained Invoices.')
                display.update("无取得发票记录！")
                # 若设置要截图则截图
                if config.config["invoice_inquiry_ps"]:
                    logger.info(f'Creating Screenshot...')
                    display.update("无取得发票记录，正在截图！")
                    zoom_and_ps(driver, '电局取得发票记录')
    # 是零申报则返回True
    logger.info(f"judge_zero: {corp.name} Is Zero, Returning True...")
    display.update(f"{corp.name}是零申报")
    return True


# ✅已完工，已测试，可优化
# 非零申报的处理，关闭标签页，退出登录
def not_zero_process(driver, corp, display):
    # 关闭多的标签页
    logger.info(f"{corp.name} Is Not 'Zero', Closing Tabs And Returning Home Page...")
    display.update(f"{corp.name}非“零申报”，请手动做账并申报，正在关闭标签页...", color='red')
    time.sleep(0.5)

    # 关闭其它标签页并返回主页
    logger.info('Closing Tabs...')
    close_tab_return_home(driver, display)

    # 在主页退出登录
    logger.info('Logging out...')
    display.update(f"{corp.name}非“零申报”，请手动做账并申报，正在退出登录...", color='red')
    time.sleep(0.5)
    if not logout():
        # 登出后检测登录按钮超时（小概率），抛异常，后续有时间可以开发处理逻辑，比如刷新或重试什么的
        logger.error("Could Not Confirm Whether Entered Main Page After Logout.")
        display.update("遇到未知异常，程序将终止。", color='red')
        raise UnknownError("Could Not Confirm Whether Entered Main Page After Logout.")

    # 非零申报返回主函数，主函数会抛异常给taskmgr
    for i in range(5):
        display.update(f"{corp.name}非“零申报”，请手动做账并申报，{5 - i} 秒后开始下一公司。", color='red')
        time.sleep(1)
    logger.info(f"{corp.name} Is Not 'Zero', Preparing To Next Corp.")
    return None


# ✅已完工，已测试，需优化（多于5条的处理逻辑需完善，目前只做了截图）
# 零申报的判断逻辑，目前只进行增值税和水利建设的零申报
def zero_declare(driver, corp, display, pos):
    # 开始申报零申报，目前只负责水利建设基金和增值税
    declare = ZeroDeclare()
    logger.info('Starting Declare...')
    display.update(f"开始零申报，正在检测需要申报的项目...")
    time.sleep(0.5)

    # 判断是否多于5行
    more_box = (re_coord(980), re_coord(560), re_coord(110), re_coord(30))
    if check_img('more_declared', more_box) or check_img('more_undeclared', more_box):
        logger.info('There Are More Than Five Items In My_To_Do.')
        more = True  # 标记是否有多于5行
    else:
        logger.info('There Has Only Five Or Less Items In My_To_Do.')
        more = False

    # 先检查我的待办前5行
    logger.info(f"Checking Undeclared Items In Front Five Lines...")
    for i in range(5):
        # 检查状态标签
        logger.info(f"Checking The Status Of Line {i + 1}")
        status_box = (re_coord(980), re_coord(343 + i * 43), re_coord(120), re_coord(43))
        status = recognize_status(status_box)
        # 状态标签为已申报
        if status == '已申报':
            # 已申报则跳过
            logger.info(f"Item In Line {i + 1} Declared.")
            continue
        # 状态标签为未申报
        elif status == '未申报':
            # 未申报则去判断是啥项目并做出对应处理
            logger.info(f"Line {i + 1} Undeclared, Recognizing Item...")
            item_box = (re_coord(548), re_coord(343 + i * 43), re_coord(320), re_coord(43))
            item = recognize_items(item_box)
            # 按项目判断的逻辑，后续可根据需求增删改，需要同步改recognize_items函数
            ########################################################################################################################
            # 目前只进行增值税和水利建设的零申报
            if item == '增值税及附加税费申报' or item == '水利建设基金':
                logger.info(f"Item Is {item}, Starting Zero Declare...")
                display.update(f"{item}未申报，准备开始零申报")
                time.sleep(1)
                click_box = position(re_coord(1118), re_coord(343 + i * 43), re_coord(32), re_coord(43))
                display.update("正在进入申报页面...")
                time.sleep(0.5)
                x, y = box_center(click_box)
                pyautogui.click(x, y)
                time.sleep(5)
                declare.mixdeclare_zzs_sljsjj(driver, display, pos, zzs=True, sljs=True)
                time.sleep(1)
                display.update(f"已返回主页，继续检测需要申报的项目...")
            # 遇到不认识的项目，暂时先跳过，后续可以添加上处理逻辑，比如截图该项目后做记录
            elif not item:
                logger.info(f"Item Not Supported, Skipping...")
                continue
            # 其余项目跳过
            else:
                logger.info(f"Item Is {item}, Skipping...")
                continue
        ########################################################################################################################
        # 未检测到状态标签（返回False）
        else:
            # 返回False说明没检测到状态标签，跳过这行（是否可以直接视为已全部完成？）
            logger.info(f"No Status Tag Found In Line {i + 1}")
            continue
        # 处理完一行，开始下一个循环
        logger.info(f"Line {i + 1} Completed, Starting The Next One...")

    # 我的待办前5行完成
    logger.info(f"Front Five Lines Completed.")
    display.update(f"前五项需要申报的项目均已申报")
    # 若需截图则刷新后截图
    if config.config["etax_home_ps"]:
        logger.info(f"Five Lines Completed, Refreshing status...")
        display.update(f"前五项需要申报的项目均已申报，正在刷新状态...")
        time.sleep(1)
        # 刷新并判断是否在主页了
        pyautogui.press('f5')
        time.sleep(1)
        judge_home_page(display)
        logger.info(f"Status Refreshed, Creating Screenshot...")
        display.update(f"状态已刷新，正在截图...")
        time.sleep(3)
        driver.save_screenshot(f'{config.screenshot_dir}电局申报完成我的待办截图01.png')
        logger.info(f"Screenshot Created.")
        display.update(f"截图完毕，请在运行结果文件夹中查看")

    # ❌若有超过五项则开始后面的(目前只做截图，后续有时间再完善)❌
    time.sleep(0.5)
    if more:
        logger.info(f"Processing More Items, Mouse Wheel Down...")
        pyautogui.click(re_coord(1086), re_coord(400))
        time.sleep(0.2)
        pyautogui.scroll(-100, x=re_coord(1086), y=re_coord(400))  # 鼠标滚轮向下
        time.sleep(1)
        # 若需截图则截图
        if config.config["etax_home_ps"]:
            driver.save_screenshot(f'{config.screenshot_dir}电局申报完成我的待办截图02.png')
            logger.info(f"Screenshot Created.")
            display.update(f"截图完毕，请在运行结果文件夹中查看")

    # 所有项目均完成
    logger.info(f"All Items Completed.")
    display.update(f"所有需要申报的项目均已申报")
    time.sleep(1)
    logger.info(f"zero_declare: Returning main...")
    return None


# 该函数未完工，可能会弃用
# 旧版本的零申报处理逻辑，先检测项目再判断状态，新版的是先判断状态再判断项目
def zero_declare_old(driver, corp, display):  # 未完成，那几个项目是否固定的？待研究，申报过程目前无截图素材，需获取
    # 开始零申报，两种逻辑，一种是根据项目找位置，另一种是在固定位置判断项目，目前是前者
    logger.info('Starting Declare...')
    display.update(f"开始申报")
    time.sleep(0.5)
    TO_DO_box = (re_coord(530), re_coord(344), re_coord(344), re_coord(242))
    # 处理水利建设基金
    logger.info("Checking 水利建设基金...")
    display.update(f"正在检查水利建设基金...")
    SLJSJJ = check_img('SLJSJJ', TO_DO_box, pos=True)
    if SLJSJJ is not False:
        SLJSJJ_status_box = (SLJSJJ.top + re_coord(422), SLJSJJ.left - re_coord(14), re_coord(116), re_coord(42))
        status = recognize_status(SLJSJJ_status_box)
        logger.info(f"水利建设基金{status}")
        display.update(f"水利建设基金{status}")
        if status == '未申报':
            logger.info('正在申报水利建设基金...')
            display.update(f"正在申报水利建设基金...")
            time.sleep(0.5)
            pyautogui.click(x=SLJSJJ.top + re_coord(578), y=SLJSJJ.left + re_coord(8))
            # declare()
            display.update(f"水利建设基金申报完毕")
    else:
        # 找不到水利建设基金怎么处理
        pass
    close_tab_return_home(driver, display)
    time.sleep(0.5)
    # 处理残疾人就业保障金
    logger.info("Checking 残疾人就业保障金...")
    display.update(f"正在检查残疾人就业保障金...")
    CJRJYBZJSB = check_img('CJRJYBZJSB', TO_DO_box, pos=True)
    if CJRJYBZJSB is not False:
        CJRJYBZJSB_status_box = (
            CJRJYBZJSB.top + re_coord(422), CJRJYBZJSB.left - re_coord(14), re_coord(116), re_coord(42))
        status = recognize_status(CJRJYBZJSB_status_box)
        logger.info(f"残疾人就业保障金{status}")
        display.update(f"残疾人就业保障金{status}")
        if status == '未申报':
            logger.info('正在申报残疾人就业保障金...')
            display.update(f"正在申报残疾人就业保障金...")
            time.sleep(0.5)
            pyautogui.click(x=CJRJYBZJSB.top + re_coord(578), y=CJRJYBZJSB.left + re_coord(8))
            # declare()
            display.update(f"残疾人就业保障金申报完毕，")
    else:
        # 找不到怎么处理
        pass
    close_tab_return_home(driver, display)
    time.sleep(0.5)
    # 处理增值税及附加税费
    logger.info("Checking 增值税及附加税费...")
    display.update(f"正在检查增值税及附加税费...")
    ZZSJFJSFSB = check_img('ZZSJFJSFSB', TO_DO_box, pos=True)
    if ZZSJFJSFSB is not False:
        ZZSJFJSFSB_status_box = (
            ZZSJFJSFSB.top + re_coord(422), ZZSJFJSFSB.left - re_coord(14), re_coord(116), re_coord(42))
        status = recognize_status(ZZSJFJSFSB_status_box)
        logger.info(f"增值税及附加税费{status}")
        display.update(f"增值税及附加税费{status}")
        if status == '未申报':
            logger.info('正在申报增值税及附加税费...')
            display.update(f"正在申报增值税及附加税费...")
            time.sleep(0.5)
            pyautogui.click(x=ZZSJFJSFSB.top + re_coord(578), y=ZZSJFJSFSB.left + re_coord(8))
            # declare()
            display.update(f"增值税及附加税费申报完毕")
    else:
        # 找不到怎么处理
        pass
    close_tab_return_home(driver, display)
    time.sleep(0.5)
    # 申报完成，截图记录
    logger.info('Creating ScreenShot...')
    display.update('正在截图...')
    sc_path = f"{config.screenshot_dir}电局零申报完成截图.png"
    driver.save_screenshot(sc_path)
    return None


# ✅已部分完工，已测试，可优化
# 零申报的执行逻辑，目前只开发了通过综合申报进行增值税和水利建设的零申报
class ZeroDeclare:
    """
    https://etax.shaanxi.chinatax.gov.cn/xxbg/view/zhsffw/#/mixDeclare/nav   # 简易确认式申报
    https://etax.shaanxi.chinatax.gov.cn/xxbg/view/zhsffw/#/relevanceDeclare/nav   # 综合关联式申报
    https://etax.shaanxi.chinatax.gov.cn/sbzx/view/lzsfjssb/#/declare/zzsxgmnsrsb?jyjkId=20   # 增值税及附加税费小规模申报
    https://etax.shaanxi.chinatax.gov.cn/sbzx/view/lzsfjssb/#/declare/zzsybnsrsb?jyjkId=10   # 增值税及附加税费一般申报
    https://etax.shaanxi.chinatax.gov.cn/sbzx/view/sdsfsgjssb/#/yyzx/tysb?ZsxmDm=30221   # 水利建设基金申报
    """

    @staticmethod
    # ✅通过简易确认式申报进行增值税和水利建设基金的零申报，已完工，已测试
    def mixdeclare_zzs_sljsjj(driver, display, pos, zzs=False, sljs=False):
        """
        通过综合申报进行增值税和水利建设基金的申报
        通过URL还是点击申报需探究，是否每次都能进综合申报（是有多个未申报项目时才能进还是每次都能？）
        综合关联式申报和简易确认式申报有啥区别
        """
        # 打开增值税小规模纳税人申报页面并切换浏览器焦点
        # driver.execute_script(
        #     "window.open('https://etax.shaanxi.chinatax.gov.cn/xxbg/view/zhsffw/#/mixDeclare/nav', '_blank');")
        # all_handles = driver.window_handles
        # driver.switch_to.window(all_handles[1])  # 切换到增值税小规模纳税人申报

        # 检测是否有取消全选按钮，判断是否已进入mixDeclare或relevanceDeclare页面，若进入则点击取消全选按钮
        timer = 60
        deselect_all_box = (re_coord(188), re_coord(400), re_coord(120), re_coord(130))
        for i in range(timer):
            # 若检测到取消全选则点击后跳出循环
            deselect_all = check_img('deselect_all', deselect_all_box, pos=True)
            if deselect_all:
                logger.info('Entered mixdeclare, deselecting all...')
                display.update('已进入综合申报，正在取消全选...')
                time.sleep(1)
                x, y = box_center(deselect_all)
                pyautogui.click(x, y)
                time.sleep(1)
                break
            # 检查404
            check_404(pos, 'Checking Whether Entered Mix Declare.', display)
            time.sleep(1)
            timer -= 1
            # 30秒后刷新一次
            if timer == 30:
                logger.info('Waited 30s, Refreshing...')
                pyautogui.press('f5')
                time.sleep(1)
            # 超时处理逻辑
            if timer == 0:
                logger.error('Timeout While Entering Mix Declare And Seems Not 404.')
                raise TimeOutError('Timeout While Entering Mix Declare And Seems Not 404.')

        # 取消全选后，向下翻页一次，然后选择增值税和水利建设基金（若有）
        logger.info('Selecting items...')
        display.update('已进入综合申报，正在选择申报项目...')
        pyautogui.press('pagedown')
        time.sleep(1)
        selected_flag = False
        item_box = (re_coord(82), re_coord(326), re_coord(1010), re_coord(364))
        if zzs:
            # 选择增值税
            declare_zzs = check_img('declare_zzs', item_box, pos=True)
            if declare_zzs:
                # 检测到增值税，点击并选择
                logger.info('Found 增值税 While Selecting items, Selecting...')
                display.update('已进入综合申报，正在选择申报项目，已选择增值税。')
                x, y = box_center(declare_zzs)
                time.sleep(0.5)
                pyautogui.click(x, y)
                time.sleep(0.5)
                selected_flag = True
            else:
                # 找不到则跳过
                logger.info('Could Not Found 增值税 While Selecting items.')
                pass
        if sljs:
            # 选择水利建设
            declare_sljs = check_img('declare_sljs', item_box, pos=True)
            if declare_sljs:
                # 检测到水利建设，点击并选择
                logger.info('Found 水利建设 While Selecting items, Selecting...')
                display.update('已进入综合申报，正在选择申报项目，已选择水利建设基金。')
                x, y = box_center(declare_sljs)
                time.sleep(0.5)
                pyautogui.click(x, y)
                time.sleep(0.5)
                selected_flag = True
            else:
                # 找不到则跳过
                logger.info('Could Not Found 水利建设 While Selecting items.')
                pass

        # 若没选项目则返回，选择完项目后点击"去办理"进入申报页面（税费试算页面）
        if not selected_flag:
            logger.info("Selected Nothing, Returning Home Page...")
            display.update('未找到可选项目，正在返回主页...')
            close_tab_return_home(driver, display)
            return None
        logger.info("Selecting Completed, Clicking 去办理 Button...")
        display.update('项目选择完毕，正在进入税费试算页面...')
        pyautogui.click(re_coord(636), re_coord(710))   # 点击去办理按钮

        # 检测是否进入税费试算页面，若进入则按三次下翻页到页面最底部填0(若需)
        timer = 60
        tax_calculation_box = (re_coord(190), re_coord(258), re_coord(330), re_coord(122))
        for i in range(timer):
            # 若检测到税费试算则按三次下翻页到页面最底部填0(若需)后跳出循环
            if check_img('tax_calculation', tax_calculation_box):
                logger.info('Entered tax calculation.')
                display.update('已进入税费试算，正在向下翻页...')
                time.sleep(1)
                pyautogui.press('pagedown')
                time.sleep(0.5)
                pyautogui.press('pagedown')
                time.sleep(0.5)
                pyautogui.press('pagedown')
                time.sleep(1)
                # 翻到最后需要填写水利建设的应税项
                if sljs:
                    logger.info("Inputting 0 Of 水利建设的应税项...")
                    display.update('正在填写 0 ...')
                    pyautogui.click(re_coord(436), re_coord(612))  # 点击文本框
                    time.sleep(0.5)
                    pyautogui.typewrite('0')  # 输入0
                    time.sleep(0.5)
                    logger.info("Clicking next_step_button...")
                    display.update('正在进入申报确认页面...')
                    pyautogui.click(re_coord(696), re_coord(736))  # 点击下一步
                    time.sleep(1)
                break
            # 检查404
            check_404(pos, 'Checking Whether Entered Mix Declare.', display)
            time.sleep(1)
            timer -= 1
            # 30秒后刷新一次
            if timer == 30:
                logger.info('Waited 30s, Refreshing...')
                pyautogui.press('f5')
                time.sleep(1)
            # 超时处理逻辑
            if timer == 0:
                logger.error('Timeout While Entering tax_calculation And Seems Not 404.')
                raise TimeOutError(
                    'Timeout While Entering tax_calculation And Seems Not 404.')

        # 检测是否进入申报确认页面（检测提交申报按钮）
        timer = 60
        submit_declaration_button_box = (re_coord(640), re_coord(704), re_coord(134), re_coord(66))
        for i in range(timer):
            # 检测到下一步按钮则点击
            if check_img('next_step_button', submit_declaration_button_box):
                logger.info('next_step_button Detected, Checking Again...')
                pyautogui.click(re_coord(696), re_coord(736))
                time.sleep(0.5)
            # 检测提交申报按钮(用于判断是否进入申报确认界面)
            pyautogui.moveTo(re_coord(750), re_coord(600))
            submit_declaration_button = check_img('submit_declaration_button', submit_declaration_button_box, pos=True)
            if submit_declaration_button:
                logger.info("Entered 申报确认页面, Submitting Declaration...")
                display.update('已进入申报确认页面， 正在提交申报...')
                time.sleep(0.5)
                # 点击提交申报按钮
                x, y = box_center(submit_declaration_button)
                pyautogui.click(x, y)
                time.sleep(3)
                # 点击"真实责任"随后确定（都在函数里）
                logger.info('Clicking 真实责任...')
                click_zszr()
                break
            # 检查404
            check_404(pos, 'Checking Whether Entered zzsybnsrsb.', display)
            time.sleep(1)
            timer -= 1
            # 30秒后刷新一次
            if timer == 30:
                logger.info('Waited 30s, Refreshing...')
                pyautogui.press('f5')
                time.sleep(1)
            # 超时处理逻辑
            if timer == 0:
                logger.error('Timeout While Checking Entered 申报确认页面 And Seems Not 404.')
                raise TimeOutError('Timeout While Checking Entered 申报确认页面 And Seems Not 404.')

        # 检测是否申报成功，若成功则截图
        timer = 60
        declare_success_box = (re_coord(366), re_coord(274), re_coord(556), re_coord(336))
        for i in range(timer):
            # 检测是否申报成功，若成功则截图后回主页
            if check_img('declare_success', declare_success_box):
                logger.info('declare_success Confirmed.')
                display.update('申报成功')
                time.sleep(0.5)
                # 若设置需要截图则截图
                if config.config["etax_item_ps"]:
                    logger.info('申报成功，正在截图...')
                    # 设置截图文件名
                    if sljs and zzs:
                        img_name = '电局增值税和水利建设申报成功截图'
                    elif sljs:
                        img_name = '电局水利建设申报成功截图'
                    else:
                        img_name = '电局增值税申报成功截图'
                    # 截图
                    driver.save_screenshot(f'{config.screenshot_dir}{img_name}.png')
                    time.sleep(0.3)
                    logger.info(f'截图完毕，保存路径：{config.screenshot_dir}{img_name}.png')
                # 处理完后关闭标签页返回主页
                time.sleep(1)
                close_tab_return_home(driver, display)
                return None
            # 30秒后刷新一次
            if timer == 30:
                logger.info('Waited 30s, Refreshing...')
                pyautogui.press('f5')
                time.sleep(1)
            # 检查404
            check_404(pos, 'Checking Whether declare_success.', display)
            time.sleep(1)
            timer -= 1
            # 超时处理逻辑
            if timer == 0:
                logger.error('Timeout While Checking Whether declare_success And Seems Not 404.')
                raise TimeOutError(
                    'Timeout While Checking Whether declare_success And Seems Not 404.')

    @staticmethod
    # ❌增值税小规模纳税人申报，单独申报，未完工，未测试，可能弃用
    def zzsxgmnsrsb(driver, corp, display, pos):
        # 打开增值税小规模纳税人申报页面并切换浏览器焦点
        driver.execute_script(
            "window.open('https://etax.shaanxi.chinatax.gov.cn/sbzx/view/lzsfjssb/#/declare/zzsxgmnsrsb?jyjkId=20', '_blank');")
        all_handles = driver.window_handles
        driver.switch_to.window(all_handles[1])  # 切换到增值税小规模纳税人申报
        # 判断是否进入增值税小规模纳税人申报页面，若进入则点击申报按钮
        timer = 30
        zzsxgmnsrsb_box = (re_coord(180), re_coord(194), re_coord(308), re_coord(60))
        submit_declaration_button_box = (re_coord(640), re_coord(704), re_coord(134), re_coord(66))
        for i in range(30):
            # 若检测到发票业务则点击后跳出循环
            if check_img('zzsxgmnsrsb', zzsxgmnsrsb_box):
                logger.info('Entered zzsxgmnsrsb, declaring...')
                display.update('已进入增值税小规模纳税人申报，正在申报...')
                time.sleep(1)
                submit_declaration_button = check_img('submit_declaration_button', submit_declaration_button_box,
                                                      pos=True)
                if submit_declaration_button:
                    x, y = box_center(submit_declaration_button)
                    pyautogui.click(x, y)
                    time.sleep(3)
                    break
                else:
                    pass
            # 检查404
            check_404(pos, 'Checking Whether Entered zzsxgmnsrsb.', display)
            time.sleep(1)
            timer -= 1
            # 超时处理逻辑
            if timer == 0:
                logger.error('Timeout While Declaring zzsxgmnsrsb And Seems Not 404.')
                raise TimeOutError(
                    'Timeout While Declaring zzsxgmnsrsb And Seems Not 404.')
            # 检测是否有通知弹窗（小规模可能有一个）

        # 点击真实责任
        click_zszr()
        timer = 30

    @staticmethod
    # ❌增值税一般纳税人申报，单独申报，未完工，未测试，可能弃用
    def zzsybnsrsb(driver, corp, display, pos):  # 未完工，未测试
        # 打开增值税小规模纳税人申报页面并切换浏览器焦点
        driver.execute_script(
            "window.open('https://etax.shaanxi.chinatax.gov.cn/sbzx/view/lzsfjssb/#/declare/zzsybnsrsb?jyjkId=10', '_blank');")
        all_handles = driver.window_handles
        driver.switch_to.window(all_handles[1])  # 切换浏览器焦点
        # 判断是否进入页面，若进入则点击申报按钮
        timer = 30
        zzsybnsrsb_box = (re_coord(180), re_coord(194), re_coord(308), re_coord(60))
        submit_declaration_button_box = (re_coord(640), re_coord(704), re_coord(134), re_coord(66))
        for i in range(30):
            # 判断是否进入页面，若进入则点击申报按钮
            if check_img('zzsybnsrsb', zzsybnsrsb_box):
                logger.info('Entered zzsybnsrsb, declaring...')
                display.update('已进入增值税一般纳税人申报，正在申报...')
                time.sleep(1)
                submit_declaration_button = check_img('submit_declaration_button', submit_declaration_button_box,
                                                      pos=True)
                if submit_declaration_button:
                    x, y = box_center(submit_declaration_button)
                    pyautogui.click(x, y)
                    time.sleep(3)
                    break
                else:
                    pass
            # 检查404
            check_404(pos, 'Checking Whether Entered zzsybnsrsb.', display)
            time.sleep(1)
            timer -= 1
            # 超时处理逻辑
            if timer == 0:
                logger.error('Timeout While Declaring zzsybnsrsb And Seems Not 404.')
                raise TimeOutError(
                    'Timeout While Declaring zzsybnsrsb And Seems Not 404.')
        # 点击真实责任
        click_zszr()
        timer = 30


# 该函数未完工，可能会弃用
# 未开发，driver用于截图，Corp和display用于显示，item用于判断申报项目
def declare1(driver, corp, display, item, click_box):
    # 点击按钮进入申报界面
    logger.info("Entered function declare, clicking button '填写申报表'")
    display.update("正在进入申报页面...")
    time.sleep(0.5)
    x, y = box_center(click_box)
    pyautogui.click(x, y)
    pass


# ✅已完工，已测试，可优化
# 该函数功能：全量发票查询中，根据当前日期和纳税人类型来选择日期
def select_date(corp, display):
    # 开始选择日期，设置变量
    logger.info(f"Starting select_date...")
    c_year = int(get_current_date(y=True))
    c_month = int(get_current_date(m=True))
    month_select_left_box = (re_coord(392), re_coord(292), re_coord(98), re_coord(50))
    month_select_right_box = (re_coord(678), re_coord(292), re_coord(86), re_coord(50))

    # 根据公司类型继续设置变量
    if corp.type == '小规模纳税人':
        # 小规模纳税人4, 7, 10, 1月才季度申报，对应日期范围: 0101-0331，0401-0630，0701-0930，1001-1231
        logger.info(f"{corp.name}是{corp.type}，正在查询季度发票记录")
        display.update(f"{corp.name}是{corp.type}，正在查询季度发票记录...")
        month_minus_times = 3
        if c_month == 1:
            s_month = c_month + 9
            s_year = c_year - 1
        else:
            s_month = c_month - month_minus_times
            s_year = c_year
        time.sleep(0.5)
    else:
        # 一般纳税人每月申报，日期范围是上个月起止
        logger.info(f"{corp.name}是{corp.type}，正在查询月度发票记录")
        display.update(f"{corp.name}是{corp.type}，正在查询月度发票记录...")
        month_minus_times = 1
        if c_month == 1:
            s_month = c_month + 11
            s_year = c_year - 1
        else:
            s_month = c_month - month_minus_times
            s_year = c_year
        time.sleep(0.5)

    # 变量设置完毕，开始设置左边的开始日期
    logger.info('Clicking Left Date Box...')
    click(268, 590)  # 点击左边日期框
    time.sleep(1)
    # 识别并跳转至当前月份
    month_current = check_img('month_current', month_select_left_box, pos=True)
    if month_current:
        logger.info('Clicking Left Current_Month Button...')
        x, y = box_center(month_current)
        pyautogui.click(x, y)
        time.sleep(0.5)
        # 跳转至当前月份后，识别并点击月份减
        month_previous = check_img('month_previous', month_select_left_box, pos=True)
        if month_previous:
            logger.info(f'Clicking Left Previous Month Button {month_minus_times} Times...')
            x, y = box_center(month_previous)
            for i in range(month_minus_times):
                pyautogui.click(x, y)
                time.sleep(0.5)
            time.sleep(0.5)
            # 到目标月份后，识别并点击1号
            logger.info(f'Clicking Left Date 1 ...')
            date_clicker(s_year, s_month, 1)
            time.sleep(0.5)
        else:
            # 若找不到月份减按钮则返回False，会再试一次
            logger.warning('Could Not Find Left Previous Month Button, select_date Returning False')
            return False
    else:
        # 若找不到当前月份按钮则返回False，会再试一次
        logger.warning('Could Not Find Left Current Month Button, select_date Returning False')
        return False

    # 设置完左边的开始日期，设置右边的结束日期
    logger.info('Clicking Right Date Box...')
    click(550, 590)  # 点击右边日期框
    time.sleep(1)
    # 识别并跳转至当前月份
    month_current = check_img('month_current', month_select_right_box, pos=True)
    if month_current:
        logger.info('Clicking Right Current Month Button...')
        x, y = box_center(month_current)
        pyautogui.click(x, y)
        time.sleep(0.5)
        # 跳转至当前月份后，识别并点击月份减
        month_previous = check_img('month_previous', month_select_right_box, pos=True)
        if month_previous:
            logger.info('Clicking Right Previous Month Button 1 Time...')
            x, y = box_center(month_previous)
            pyautogui.click(x, y)
            time.sleep(0.5)
            # 判断、识别并点击对应日期
            img_date = date_selector(c_month - 1)  # 判断前一个月的天数
            logger.info(f'Clicking Right Date {img_date}...')
            date_clicker(s_year, s_month + month_minus_times - 1, img_date, right=True)
            time.sleep(0.5)
        else:
            # 若找不到月份减按钮则返回False，会再试一次
            logger.warning('Could Not Find Right Previous Month Button, select_date Returning False')
            return False
    else:
        # 若找不到当前月份按钮则返回False，会再试一次
        logger.warning('Could Not Find Right Current Month Button, select_date Returning False')
        return False
    # 过程无异常则返回True
    return True


# ✅已完工，已测试，逻辑不太好，可优化
# 该函数功能：全量发票查询中，选择完日期后，由该函数查询并判断是否有发票记录，有则返回True（若无则为零申报）
def inquiry_and_judge(corp, display):  # 判断有无发票，已完工，已测试，待优化（对于无发票的判断）
    logger.info('Starting inquiry_and_judge, Clicking Inquiry Button...')
    click(1072, 592)  # 点击查询
    time.sleep(3)
    export_button_box = (re_coord(44), re_coord(262), re_coord(206), re_coord(508))
    no_data_box = (re_coord(562), re_coord(262), re_coord(206), re_coord(508))
    timer = 10
    for i in range(timer):
        # 检测到导出按钮则说明有发票
        if timer > 5:
            if check_img('export_button', export_button_box):
                logger.info(f'{corp.name}检测到发票记录')
                display.update("检测到存在发票，非零申报，请手动做账并申报！", color='red')
                time.sleep(0.5)
                return True
        # 5秒后如果没有导出按钮则下去确认是否暂无数据
        if timer <= 5:
            if timer == 5:
                pyautogui.press('pagedown')
                time.sleep(0.3)
                pyautogui.press('pagedown')
                time.sleep(0.5)
            if check_img('no_data', no_data_box):
                # 检测到暂无数据则说明无发票
                logger.info(f'{corp.name}无发票记录')
                display.update("无发票记录")
                pyautogui.press('pageup')
                time.sleep(0.3)
                pyautogui.press('pageup')
                time.sleep(0.5)
                return False
            else:
                # 如果既没导出按钮也没暂无数据则有可能是还在加载，小概率，待开发
                pass
        timer -= 1
        time.sleep(1)


# ✅已完工，已测试，可优化
# 该函数功能：全量发票查询中，判断完成后，若需截图，则由该函数对页面进行缩小后截图再复原
def zoom_and_ps(driver, img_name):  # 已完工，已测试
    # 切换浏览器焦点到全量发票查询
    all_handles = driver.window_handles  # 获取所有标签页句柄
    # driver.switch_to.window(all_handles[2])  # 切换到最后一个标签页
    for handle in all_handles:
        driver.switch_to.window(handle)
        if "全量发票查询" in driver.title:
            break

    # 缩小页面
    logger.info('准备截图，正在缩小页面...')
    for i in range(4):
        pyautogui.hotkey('ctrl', '-')
        time.sleep(0.5)

    # 检查截图文件夹，若无则创建，随后截图
    mkdir(config.screenshot_dir)
    time.sleep(1)
    logger.info('正在截图')
    driver.save_screenshot(f'{config.screenshot_dir}{img_name}.png')
    time.sleep(0.3)
    logger.info(f'截图完毕，保存路径：{config.screenshot_dir}{img_name}.png')

    # 复原页面缩放比例
    pyautogui.hotkey('ctrl', '0')
    time.sleep(1)
    return None


# ✅已完工，已测试，可优化
# 该函数功能：关闭所有标签页并返回电局主页
def close_tab_return_home(driver, display):  # 已完工，已测试
    logger.info('Closing Tabs Expect The First Tab (Home Page).')
    # 获取所有标签页的句柄
    window_handles = driver.window_handles
    # 关闭除了主页的所有标签页
    if len(window_handles) != 1:
        driver.switch_to.window(window_handles[0])
        for handle in window_handles[1:]:
            driver.switch_to.window(handle)
            driver.close()
        driver.switch_to.window(window_handles[0])  # 最后再切回第一个标签页（一定要切回去）
    # 判断是否在主页了
    judge_home_page(display)
    time.sleep(0.2)
    # 刷新主页
    pyautogui.press('f5')
    time.sleep(0.2)
    judge_home_page(display)
    return None


# ✅已完工，已测试
# 判断是否在电局主页，通过检查有没有"我的待办"栏目，检测到则返回，超时抛异常
def judge_home_page(display):
    # 判断是否在主页
    timer = 60
    my_to_do_box = (re_coord(530), re_coord(210), re_coord(116), re_coord(56))
    for i in range(timer):
        # 检查我的待办
        if check_img('my_to_do', my_to_do_box):
            logger.info('Home Page Returning Confirmed After Other Tabs Closed.')
            time.sleep(0.5)
            return None
        # 30秒后刷新一次
        if timer == 30:
            logger.info('Waited 30s, Refreshing...')
            pyautogui.press('f5')
            time.sleep(1)
        # 超时处理逻辑
        if timer == 0:
            # 关标签页后检测主页超时（极小概率），抛异常，后续可加处理逻辑
            display.update("遇到未知异常，程序将终止。", color='red')
            time.sleep(3)
            logger.error("Could Not Confirm Whether Entered Home Page, Timeout.")
            raise UnknownError("Could Not Confirm Whether Entered Home Page, Timeout.")
        timer -= 1
        time.sleep(1)


# ✅已完工，已测试
# 在电局主页后，登出当前纳税人并检测是否成功登出（检测登录按钮），成功则返回True
def logout():
    logger.info('Logging Out...')
    # 点击头像打开菜单
    click(1192, 162)
    time.sleep(1)
    # 识别退出登录按钮并点击
    logout_button_box = (re_coord(840), re_coord(194), re_coord(418), re_coord(176))
    timer = 10
    for i in range(timer):
        logout_button = check_img('logout_button', logout_button_box, pos=True)
        if logout_button:
            time.sleep(0.5)
            x, y = box_center(logout_button)
            logger.info('Clicking Logout Button...')
            pyautogui.click(x, y)
            break
        # 超时处理逻辑，小概率，待完善
        if timer == 0:
            pass
        timer -= 1
        time.sleep(1)

    # 判断是否成功登出，检测登录按钮
    timer = 60
    login_button_box = (re_coord(1152), re_coord(138), re_coord(108), re_coord(56))
    for i in range(timer):
        if check_img('login_button', login_button_box):
            time.sleep(0.5)
            logger.info('Logout Confirmed.')
            return True
        # 30秒后刷新一次
        if timer == 30:
            logger.info('Waited 30s, Refreshing...')
            pyautogui.press('f5')
            time.sleep(1)
        # 超时处理逻辑
        if timer == 0:
            logger.error("Could Not Find Login Button After Logout.")
            return False
        timer -= 1
        time.sleep(1)


# ✅已完工，已测试，需要不断扩充
# 给定一个box，识别该box内的税种
def recognize_items(box):  # 已完工，但是项目需补充更多
    if check_img('QYSDSNDSB', box):
        result = '企业所得税年度申报'
    elif check_img('SLJSJJ', box):  # 这个是通用申报，可能不一定是水利建设基金
        result = '水利建设基金'
    elif check_img('CJRJYBZJSB', box):
        result = '残疾人就业保障金申报'
    elif check_img('ZZSJFJSFSB', box):
        result = '增值税及附加税费申报'
    else:
        result = False
    return result


# ✅已完工，已测试
# 判断已申报还是未申报
def recognize_status(box):  # 已完工，未测试
    if check_img('declared', box):
        result = '已申报'
    elif check_img('undeclared', box):
        result = '未申报'
    else:
        result = False   # 既没检查到已申报又没未申报
    return result


# ✅已完工，已测试，待优化（闰年）
# 根据月份，返回月份的最后一天
def date_selector(month):
    if month in {1, 3, 5, 7, 8, 10, 12}:
        return 31
    elif month in {4, 6, 9, 11}:
        return 30
    elif month == 2:  # 二月没分闰年
        return 28


# ✅已完工，已测试，可优化
# 给出年月日，判断该日期在日历的哪行哪列，并点击
def date_clicker(year, month, day, right=False):
    # 设置变量
    # first_day = datetime.date(year, month, 1).weekday()
    # day_of_month = datetime.date(year, month, day).day
    first_day = int(time.strftime("%w", time.strptime(f"{year}-{month}-01", "%Y-%m-%d"))) - 1
    if first_day == -1:  # 处理周日的情况
        first_day = 6
    day_of_month = day

    # 计算日期在日历中的行列（从零开始）
    row = (first_day + day_of_month - 1) // 7
    col = (first_day + day_of_month - 1) % 7
    # print(f"{year}-{month}-{day} is {row}-{col}")

    # 根据是左边日期框还是右边日期框去点击日期
    if right:
        left, top = 520, 390  # 高宽都是222，155
    else:
        left, top = 240, 390
    x = left + col * 37
    y = top + row * 31
    click(x, y)
    logger.info(f"date_clicker: {year}-{month}-{day} Clicked.")
    return None


# ✅已完工，已测试，可优化
# 提交税务申报时，点击真实责任并确定
def click_zszr():
    # 点击真实责任并确定
    logger.info("Clicking zszr...")
    pyautogui.click(re_coord(458), re_coord(500))
    time.sleep(0.5)
    pyautogui.click(re_coord(618), re_coord(526))
    time.sleep(0.5)
    pyautogui.click(re_coord(748), re_coord(508))
    time.sleep(0.5)
    pyautogui.click(re_coord(842), re_coord(524))
    time.sleep(0.5)
    # 点击确定
    pyautogui.click(re_coord(932), re_coord(606))


# ✅已完工，已测试，可优化
# 主函数
def main(driver, corp, display):
    # 设置浏览器窗口(可以放全局变量里)
    logger.debug('Entered declare.main')
    display.update("准备判断是否零申报...")
    pos = position(left=re_coord(22), top=re_coord(136), width=re_coord(1256), height=re_coord(630))
    time.sleep(0.5)

    # 判断零申报
    if not judge_zero(driver, corp, display, pos):
        # 非零申报，非零处理后抛异常回到Panel
        time.sleep(0.5)
        logger.info(f"{corp.name} Is Not 'Zero', Not Zero Processing...")
        display.update(f"{corp.name}非零申报，准备返回主页...")
        # not_zero_process(driver, corp, display)
        display.update('正在读取任务列表...')
        raise SkipError(f"{corp.name} 非零申报。")
    else:
        # 零申报，关闭多的标签页返回主页开始零申报
        time.sleep(0.5)
        logger.info(f"{corp.name} Is 'Zero', Closing Tabs And Returning Home...")
        display.update(f"{corp.name}是零申报，准备返回主页，准备开始申报...")
        # 关多的标签页并回主页
        close_tab_return_home(driver, display)
        # 回到主页，准备开始零申报
        logger.info(f"Preparing To Declare...")
        display.update(f"已返回主页，准备开始申报...")
        time.sleep(0.5)
        # 零申报函数
        zero_declare(driver, corp, display, pos)  # 零申报如果有异常，抛异常or错误码？

    # 零申报完成，退出登录
    logger.info(f"{corp.name} Declaration Finished...")
    display.update(f"{corp.name}零申报完成，正在退出登录...")
    time.sleep(0.5)
    if not logout():
        # 登出后检测主页登录按钮超时（小概率），抛异常
        logger.error("Could Not Confirm Whether Entered Main Page After Logout.")
        display.update("遇到未知异常，程序将终止。", color='red')
        raise UnknownError("Could Not Confirm Whether Entered Main Page After Logout.")

    # 过程无异常则函数返回
    logger.info('declare.main: Declare Completed, Returning Panel...')
    return None


if __name__ == '__main__':
    pop_error_window('此文件不可单独运行！！')
