from common import *
import pandas as pd
import threading
import declare
import config
import login
import time


def shuiwu(task_list, panel_display):
    # 更新panel显示器
    time.sleep(1)
    panel_display.update(f"任务类型：税务申报\n任务状态：开始运行\n任务进度：0 / {len(task_list)}\n如需终止任务请按F4")

    # 创建状态显示器
    logger.info('Starting Task...')
    status_display = StatusDisplay("程序已启动", position=(re_coord(300), re_coord(100)))

    # 启动浏览器
    logger.info("Starting Edge...")
    status_display.update("正在启动浏览器...")
    driver = start_browser()

    # 读取Excel信息
    logger.info("Reading Excel Info...")
    status_display.update("正在读取Excel信息...")
    df = pd.read_excel(config.excel_dir)
    error_flag = False
    for i in range(0, len(task_list)):

        # 读取当前任务信息
        index = task_list[i]
        corp = corpInfo(serial=index,
                        name=str(df.iloc[index - 1, 1]),
                        shibiehao=str(df.iloc[index - 1, 2]),
                        shoujihao=str(df.iloc[index - 1, 3]),
                        password=str(df.iloc[index - 1, 4]))
        # 开始登录
        panel_display.update(f"任务类型：税务申报\n任务状态：运行中\n任务进度：{i} / {len(task_list)}\n如需终止任务请按F4")
        try:
            if i == 0:
                login.main(driver, corp, status_display, first=True)
            elif error_flag:
                login.main(driver, corp, status_display, skip=True)
            else:
                login.main(driver, corp, status_display)
        except LoginInfoError:
            panel_display.update(f"任务类型：税务申报\n任务状态：信息错误\n任务进度：Null\n如需终止任务请按F4")
            i += 1
            error_flag = True
            # 记录到Excel
            continue
        except ContactNotFount:
            panel_display.update(f"任务类型：税务申报\n任务状态：无联系人\n任务进度：Null\n如需终止任务请按F4")
            i += 1
            error_flag = True
            # 记录到Excel
        except CheckImgError:
            panel_display.update(f"任务类型：税务申报\n任务状态：识图错误\n任务进度：Null\n如需终止任务请按F4")
            login.main(driver, corp, status_display, first=True)
        except WeChatError:
            panel_display.update(f"任务类型：税务申报\n任务状态：微信错误\n任务进度：Null\n如需终止任务请按F4")
            pop_error_window(f"请确保微信已打开并且已登录！\n请点击确定，任务会自动终止！",
                             title=f"{get_current_time()} 微信异常！")
            logger.info('Exiting Edge...')
            status_display.update(f"关闭浏览器")
            status_display.close()
            driver.quit()
            return  # 这里未来可以加上让用户打开微信后再检测的逻辑，或者放到总控制面板步骤（启动这个脚本前检测）
        except GetCodeError:
            panel_display.update(f"任务类型：税务申报\n任务状态：无法获取验证码\n任务进度：Null\n如需终止任务请按F4")
            i += 1
            error_flag = True
            # 记录到Excel
        except NetworkError:
            panel_display.update(f"任务类型：税务申报\n任务状态：网络错误\n任务进度：Null\n任务终止！即将返回主界面！")
            pop_error_window(f"请检查网络设置后重新运行本程序！\n请点击确定，任务会自动终止！",
                             title=f"{get_current_time()} 页面加载失败！")
            logger.info('Exiting Edge...')
            status_display.update(f"关闭浏览器")
            status_display.close()
            driver.quit()
            return
        except UnknownError:
            panel_display.update(f"任务类型：税务申报\n任务状态：未知错误\n任务进度：Null\n任务终止！即将返回主界面！")
            pop_error_window(
                f"出现未知错误！\n请检查电脑配置、Excel文件格式以及内容是否有误！\n请点击确定，任务会自动终止！",
                title=f"{get_current_time()} 未知错误！")
            logger.info('Exiting Edge...')
            status_display.update(f"关闭浏览器")
            status_display.close()
            driver.quit()
            return

        # 登录成功，进入主界面，开始进行税务申报
        logger.info('Starting Tax Declaration.')
        status_display.update("开始进行税务申报。")
        try:
            # declare.main(corp, status_display)
            pass
        except CheckImgError:
            pass
        time.sleep(3000)


def start_control_panel():
    # 创建panel显示器
    panel_display = StatusDisplay("程序开始运行", position=(re_coord(300), re_coord(100)))
    panel_display.panel_style(re_coord(1016), re_coord(11), re_coord(266), re_coord(80))
    panel_display.update(f"任务类型：税务申报\n任务状态：开始运行\n任务进度：Null\n如需终止任务请按F4")
    task_list = [1, 3, 5, 7, 12]  # 后期需要自主选择
    thread = threading.Thread(target=shuiwu(task_list, panel_display))
    thread.start()
    thread.join()
    logger.info("Task End.")
    panel_display.update(
        f"任务类型：税务申报\n任务状态：运行结束\n任务进度：{len(task_list)} / {len(task_list)}\n任务结束！即将返回主界面！")
    pop_error_window("任务结束！\n运行结果在...\n点击确定将返回主界面！", f"{get_current_time()} 任务结束")
    panel_display.close()
    return True


if __name__ == "__main__":
    device = device_check()
    config.scale = device.scale_factor
    config.img_dir = set_img_dir(step='shuiwu')
    start_control_panel()
