import etax_declare as declare
import etax_login as login
import pandas as pd
from common import *
import threading
import config
import panel
import ptax
import time
import sys


# ✅已完工，已测试，可优化
# 电局申报任务管理、异常处理
def etax_mgr(task_list, panel_display):
    # 更新panel显示器
    time.sleep(1)
    panel_display.update(f"任务类型：电局申报\n任务状态：开始运行\n任务进度：0 / {len(task_list)}\n如需终止任务请按F8")
    # 创建状态显示器
    logger.info('Starting Task...')
    status_display = StatusDisplay("程序已启动", position=(re_coord(300), re_coord(100)))
    # 创建运行结果报告Excel
    create_result_table(task='电局申报')
    # 启动浏览器
    logger.info("Starting Edge...")
    status_display.update("正在启动浏览器...")
    driver = start_browser()
    config.driver = driver
    # 读取Excel信息
    logger.info("Reading Excel Info...")
    status_display.update("正在读取Excel...")
    df = pd.read_excel(config.excel_path)
    # 开始执行任务
    error_flag = False
    for i in range(0, len(task_list)):
        # 读取当前任务纳税人信息，后续可以在配置文件里面也写入，如果配置文件有则不读表格了（比如只给某家公司自己使用）
        logger.info("Reading Task Info...")
        status_display.update("正在读取任务信息...")
        index = task_list[i] + 1
        corp = corpInfo(serial=str(df.iloc[index, 0]),  # 序号
                        name=str(df.iloc[index, 1]),  # 纳税人名称
                        type=str(df.iloc[index, 2]),  # 纳税人类型
                        taxid=str(df.iloc[index, 3]),  # 纳税人识别号
                        loginid=str(df.iloc[index, 4]),  # 登录手机号或身份证号
                        taxpwd=str(df.iloc[index, 5]),  # 电局登录密码
                        ptaxpwd=str(df.iloc[index, 6]))  # 个税登录密码

        # 根据公司类型判断本月是否需要申报
        logger.info(f"Current Task: 电局申报，{corp.name}")
        if corp.type == '小规模纳税人':
            month = int(get_current_time(m=True))
            if month != 1 and month != 4 and month != 7 and month != 10:
                logger.info(f"{corp.name}是小规模纳税人，{month}月无需申报！正在写入结果。")
                result = f'未申报，原因：该纳税人为小规模纳税人，本月（{month}月）无需申报'
                write_result_table(i + 1, corp.name, corp.type, result, red=True)
                # for i in range(5):
                #     status_display.update(f"{corp.name}是小规模纳税人，{month}月无需申报！{5 - i} 秒后开始下一公司！", 'red')
                #     time.sleep(1)
                i += 1
                config.task_fail += 1
                continue

        # 设置截图文件夹路径变量并创建文件夹
        logger.info("Creating Screenshot dir...")
        config.screenshot_dir = f'{config.result_dir}截图/{corp.serial} {corp.name}/'
        mkdir(config.screenshot_dir)

        # 开始登录
        panel_display.update(
            f"任务类型：电局申报\n任务状态：正在登录\n任务进度：{i} / {len(task_list)}\n如需终止任务请按F8")
        try:
            if i == 0:  # 初次登录
                login.main(driver, corp, status_display, first=True)
            elif error_flag:  # 出错重试登录，跳过预处理任务
                login.main(driver, corp, status_display, skip=True)
            else:  # 前置任务正常结束，新公司登录
                login.main(driver, corp, status_display)

        # 登录异常处理：登录信息错误（跳过，开始下一个）
        except LoginInfoError:
            panel_display.update(f"任务类型：电局申报\n任务状态：信息错误\n任务进度：Null\n如需终止任务请按F8")
            i += 1
            error_flag = True
            # 记录到Excel
            logger.info(f"{corp.name}登录信息错误，正在写入结果。\n")
            result = f'未申报，原因：该纳税人登录信息错误'
            write_result_table(i + 1, corp.name, corp.type, result, red=True)
            config.task_fail += 1
            continue

        # 登录异常处理：微信联系人未找到（跳过，开始下一个）
        except ContactNotFount:
            panel_display.update(f"任务类型：电局申报\n任务状态：无联系人\n任务进度：Null\n如需终止任务请按F8")
            i += 1
            error_flag = True
            # 记录到Excel
            logger.info(f"{corp.name}在微信中找不到联系人，正在写入结果。\n")
            result = f'未申报，原因：在微信中找不到该纳税人联系人'
            write_result_table(i + 1, corp.name, corp.type, result, red=True)
            config.task_fail += 1
            continue

        # 登录异常处理：识图错误（再试一次，还出错会因为未捕获异常而退出）。
        except CheckImgError:
            panel_display.update(f"任务类型：电局申报\n任务状态：识图错误\n任务进度：Null\n如需终止任务请按F8")
            login.main(driver, corp, status_display, first=True)

        # 登录异常处理：获取验证码异常（跳过，开始下一家）
        except GetCodeError:
            panel_display.update(f"任务类型：电局申报\n任务状态：无法获取验证码\n任务进度：Null\n如需终止任务请按F8")
            i += 1
            error_flag = True
            # 记录到Excel
            logger.info(f"{corp.name}无法获取验证码，正在写入结果。")
            result = f'未申报，原因：无法获取验证码'
            write_result_table(i + 1, corp.name, corp.type, result, red=True)
            config.task_fail += 1
            continue

        # 登录异常处理：微信异常（弹窗报错后结束任务）
        except WeChatError:  # 这里未来可以加上让用户打开微信后再检测的逻辑，或者放到总控制面板步骤（启动这个脚本前检测）
            panel_display.update(f"任务类型：电局申报\n任务状态：微信错误\n任务进度：Null\n如需终止任务请按F8")
            pop_error_window(f"请确保微信已打开并且已登录！\n请点击确定，任务会自动终止！",
                             title=f"{get_current_time()} 微信异常！")
            abort_etask(driver, status_display, panel_display)

        # 登录异常处理：网络异常（弹窗报错后结束任务）
        except NetworkError:
            panel_display.update(f"任务类型：电局申报\n任务状态：网络错误\n任务进度：Null\n任务终止！即将返回主界面！")
            pop_error_window(f"请检查网络设置后重新运行本程序！\n请点击确定，任务会自动终止！",
                             title=f"{get_current_time()} 页面加载失败！")
            abort_etask(driver, status_display, panel_display)

        # 登录异常处理：其它未知异常（弹窗报错后结束任务）
        except UnknownError:
            panel_display.update(f"任务类型：电局申报\n任务状态：未知错误\n任务进度：Null\n任务终止！即将返回主界面！")
            pop_error_window(
                f"出现未知错误！\n请检查电脑配置、Excel文件格式以及内容是否有误！\n请点击确定，任务会自动终止！",
                title=f"{get_current_time()} 未知错误！")
            abort_etask(driver, status_display, panel_display)

        # 登录成功，进入电局主界面，开始进行税务申报
        logger.info('Starting Tax Declaration...')
        status_display.update("开始进行税务申报。")
        panel_display.update(
            f"任务类型：电局申报\n任务状态：正在申报\n任务进度：{i} / {len(task_list)}\n如需终止任务请按F8")
        try:
            declare.main(driver, corp, status_display)

        # 申报异常处理：超时异常（通常是页面没加载出来）（再试一次）
        except TimeOutError:
            logger.info('Processing TimeOutError...')
            panel_display.update(f"任务类型：电局申报\n任务状态：超时重试\n任务进度：Null\n如需终止任务请按F8")
            # 关标签页返回主页
            declare.close_tab_return_home(driver, status_display)
            time.sleep(1)
            # 重新尝试一次，还有错会Abort
            declare.main(driver, corp, status_display)

        # 申报异常处理：网络异常（弹窗报错后结束任务）
        except NetworkError:
            panel_display.update(f"任务类型：电局申报\n任务状态：网络错误\n任务进度：Null\n任务终止！即将返回主界面！")
            pop_error_window(f"请检查网络设置后重新运行本程序！\n请点击确定，任务会自动终止！",
                             title=f"{get_current_time()} 页面加载失败！")
            abort_etask(driver, status_display, panel_display)

        # 申报异常处理：跳过任务（非零申报或者无权操作）
        except SkipError as e:
            logger.info('Processing SkipError...')
            panel_display.update(f"任务类型：电局申报\n任务状态：跳过任务\n任务进度：Null\n如需终止任务请按F8")
            # 关标签页返回主页并登出
            declare.close_tab_return_home(driver, status_display)
            time.sleep(1)
            declare.logout()
            # 记录到Excel
            logger.info(f"正在写入结果。\n")
            result = f'未申报，原因：{e}'
            write_result_table(i + 1, corp.name, corp.type, result, red=True)
            config.task_fail += 1
            i += 1
            continue

        # 申报异常处理：需要重试（重试一次）
        except RetryError:
            logger.info('Processing RetryError...')
            panel_display.update(f"任务类型：电局申报\n任务状态：重新尝试\n任务进度：Null\n如需终止任务请按F8")
            # 关标签页返回主页
            declare.close_tab_return_home(driver, status_display)
            time.sleep(1)
            # 重新尝试一次，还有错会Abort
            declare.main(driver, corp, status_display)

        # 申报异常处理：其它未知异常（弹窗报错后结束任务）
        except UnknownError:
            panel_display.update(f"任务类型：电局申报\n任务状态：未知错误\n任务进度：Null\n任务终止！即将返回主界面！")
            pop_error_window(
                f"出现未知错误！\n请检查电脑配置、Excel文件格式以及内容是否有误！\n请点击确定，任务会自动终止！",
                title=f"{get_current_time()} 未知错误！")
            abort_etask(driver, status_display, panel_display)

        # 成功申报，写入结果Excel
        logger.info(f"{corp.name}成功申报，正在写入结果。")
        result = f'已申报，运行正常'
        write_result_table(i + 1, corp.name, corp.type, result)
        # insert_checkmark(index + 2, 9)
        config.task_succeed += 1
        i += 1  # 开始下一个纳税人

    # 全部申报完成的后处理
    logger.info('All Task Finished, Exiting Edge & Closing Display...')
    status_display.update(f"关闭浏览器")
    driver.quit()
    status_display.close()
    return None


# ✅已完工，已测试
# 个税申报任务管理、异常处理
def ptax_mgr(task_list, panel_display):
    # 更新panel显示器
    time.sleep(1)
    panel_display.update(f"任务类型：个税申报\n任务状态：开始运行\n任务进度：0 / {len(task_list)}\n如需终止任务请按F8")
    # 创建状态显示器
    logger.info('Starting Task...')
    status_display = StatusDisplay("程序已启动", position=(re_coord(466), re_coord(56)))
    # 创建运行结果报告Excel
    create_result_table(task='个税申报')
    # 读取Excel信息
    logger.info("Reading Excel...")
    status_display.update("正在读取Excel信息...")
    df = pd.read_excel(config.excel_path)
    # 开始循环任务
    for i in range(0, len(task_list)):
        # 读取信息
        logger.info("Reading Task Info...")
        status_display.update("正在读取任务信息...")
        index = task_list[i] + 1
        corp = corpInfo(serial=int(df.iloc[index, 0]),
                        name=str(df.iloc[index, 1]),
                        type=str(df.iloc[index, 2]),
                        taxid=str(df.iloc[index, 3]),
                        loginid=str(df.iloc[index, 4]),
                        taxpwd=str(df.iloc[index, 5]),
                        ptaxpwd=str(df.iloc[index, 6]))
        logger.info(f"Current Task: 个税申报，{corp.name}")

        # 设置截图文件夹路径变量并创建文件夹
        logger.info("Creating Screenshot dir...")
        config.screenshot_dir = f'{config.result_dir}截图/{corp.serial} {corp.name}/'
        logger.info(f"config.screenshot_dir: {config.screenshot_dir}")
        mkdir(config.screenshot_dir)

        # 开始单个任务
        panel_display.update(
            f"任务类型：个税申报\n任务状态：开始运行\n任务进度：{i} / {len(task_list)}\n如需终止任务请按F8")
        try:
            if i == 0:
                ptax.main(corp, status_display, first=True)
            else:
                ptax.main(corp, status_display)

        except SkipError as e:
            logger.info('Processing SkipError...')
            panel_display.update(f"任务类型：个税申报\n任务状态：跳过任务\n任务进度：Null\n如需终止任务请按F8")
            # 记录到Excel
            logger.info(f"正在写入结果。\n")
            result = f'未申报，原因：{e}'
            write_result_table(i + 1, corp.name, corp.type, result, red=True)
            status_display.update("开始下一个纳税人")
            config.task_fail += 1
            i += 1
            continue

        except UnknownError:
            logger.info('Processing UnknownError...')
            panel_display.close()
            status_display.close()
            config.abort = True
            time.sleep(1)

        # 成功申报，写入结果Excel
        logger.info(f"{corp.name}成功申报，正在写入结果。")
        result = f'已申报，运行正常'
        write_result_table(i + 1, corp.name, corp.type, result)
        # insert_checkmark(index + 2, 9)
        config.task_succeed += 1
        status_display.update("开始下一个纳税人")
        i += 1  # 开始下一个纳税人

    # 全部申报完成的后处理
    logger.info('All Task Finished, Closing Display...')
    status_display.update("所有任务已完成，即将退出程序...")
    time.sleep(3)
    status_display.close()
    return None


# ✅已完工，已测试
# 从Excel读取任务标志并创建任务列表
def read_task_list():
    logger.info('Reading Task List...')
    # 根据任务类型设置标志索引
    if config.task == 'etax':
        column_index = 7
    else:
        column_index = 8

    # 读取Excel文件并创建任务列表
    df = pd.read_excel(config.excel_path)
    task_list = []

    # 若序号和纳税人名称都非空则读取任务标记,标记非空则将序号加入任务列表
    for i in range(df.shape[0] - 3):
        if pd.notna(df.iloc[i + 2, 0]) and pd.notna(df.iloc[i + 2, 1]) and pd.notna(df.iloc[i + 2, column_index]):
            task_list.append(df.iloc[i + 2, 0])
    logger.info(f'task_list: {task_list}')
    return task_list


# ✅已完工，已测试
# 个税任务中，若用户在弹窗确认信息有无变动中选择有变动，则弹窗提示手动申报并退出程序
def notice_and_abort():
    pop_error_window('人员和工资金额有变动！\n请手动申报！', title='需手动申报')
    logger.info('人员和工资金额有变动!已提醒手动申报！程序即将终止！')
    config.abort = True
    time.sleep(0.1)
    return None


# ✅已完工，已测试
# 电局申报终止的处理函数
def abort_etask(driver, status_display, panel_display):
    logger.info('Aborting Task, Exiting Edge & Closing Display...')
    status_display.update(f"关闭浏览器")
    driver.quit()
    status_display.close()
    panel_display.close()
    config.abort = True
    time.sleep(0.1)
    return None


# ✅已完工，已测试
# 主函数
def main():
    # 获取任务列表（从Excel中读取）
    logger.info('Entered taskmgr.mian, Reading Task List...')
    task_list = read_task_list()
    if len(task_list) == 0:
        # 若任务列表为空则弹窗报错并退出
        logger.info('Task List is Empty, Programme Aborting...')
        pop_error_window('任务列表为空！\n请在Excel文件中将需要申报的纳税人打勾！\n点击确定程序将会终止！')
        config.abort = True
        time.sleep(0.1)

    # 任务为电局申报
    if config.task == 'etax':
        # 创建panel显示器
        panel_display = StatusDisplay("程序开始运行", position=(re_coord(300), re_coord(100)))
        panel_display.panel_style(re_coord(1016), re_coord(11), re_coord(266), re_coord(80))
        panel_display.update(f"任务类型：税务申报\n任务状态：开始运行\n任务进度：Null\n如需终止任务请按F8")
        # 电局申报任务开始
        etax_mgr(task_list, panel_display)
        # 电局申报任务结束
        logger.info("Task Function Returned.")
        panel_display.update(
            f"任务类型：税务申报\n任务状态：运行结束\n任务进度：{len(task_list)} / {len(task_list)}\n任务结束！")
        panel_display.close()

    # 任务为个税申报
    elif config.task == 'ptax':
        # 弹窗询问信息
        title = '需要确认信息'
        text = '\n人员和工资金额是否有变动？'
        show_yn_popup(title, text, y_func=notice_and_abort, y_text='有变动', n_text='无变动')
        # 创建panel显示器
        panel_display = StatusDisplay("程序开始运行", position=(re_coord(1300), re_coord(20)))  # 启动坐标
        panel_display.panel_style(re_coord(1300), re_coord(20), re_coord(266), re_coord(80))  # 尺寸
        panel_display.update(f"任务类型：个税申报\n任务状态：开始运行\n任务进度：Null\n如需终止任务请按F8")
        # 个税申报任务开始
        ptax_mgr(task_list, panel_display)
        # 个税申报任务结束
        logger.info("Task End.")
        panel_display.update(
            f"任务类型：个税申报\n任务状态：运行结束\n任务进度：{len(task_list)} / {len(task_list)}\n任务结束！")
        panel_display.close()

    # 任务正常结束，弹窗
    pop_error_window(f"任务结束！运行结果在：\n{config.result_dir}\n点击确定将退出程序！",
                     f"{get_current_time()} 任务结束")
    config.abort = True
    time.sleep(0.1)
    return None


if __name__ == "__main__":
    config.device = check_device()
    config.scale = config.device.scale_factor
    config.img_dir = set_img_dir(step=config.task)
    main()
