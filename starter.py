from panel import *
import taskmgr
import keyboard


def start_program():
    # 启动任务面板
    logger.info("Starting Panel Thread...")
    panel_thread = threading.Thread(target=start_panel, daemon=True)
    panel_thread.start()
    panel_thread.join()
    logger.info("Panel Thread Exited, Checking Run Flag...")
    # 若运行标志为真，则设置全局变量并启动任务调度器线程和监视器线程
    if config.run:
        # 设置全局变量
        logger.info("Run Flag Is True, Setting Config...")
        config.device = device
        config.scale = device.scale_factor
        config.base_dir = os.getcwd()
        config.screen_box = (0, 0, config.device.phys_w - 1, config.device.phys_w - 1)
        config.result_dir = f'{config.result_base_dir}/{get_current_date(y=True)}/{get_current_date(m=True)}/'
        config.img_dir = set_img_dir(step=config.task)
        # 启动监视器线程
        threading.Thread(target=monitor).start()
        # 启动任务线程
        task_thread.start()
        # taskmgr.main()
    else:
        logger.info("Run Flag Is False, Aborting...\n")
        sys.exit()   # 未启动任务，程序退出


def monitor():
    logger.info("Monitor Started...")
    while True:
        # 检测终止标志
        if config.abort:
            logger.info("Abort Flag Acknowledged, Aborting...")
            break
        # 检测是否按下F8
        if keyboard.is_pressed('F8'):
            logger.info("F8 Pressed. Terminating The Task...")
            config.abort = True
            config.run = False
            break
        # 50毫秒轮询一次
        time.sleep(0.05)

    # 终止程序的步骤
    logger.info("Monitor Stopped, Aborting...\n")
    if config.driver is not None:
        config.driver.quit()
    sys.exit()   # 检测到终止标志或F8被按下，程序退出


if __name__ == "__main__":
    run_as_admin()
    logger.info("Starter Activated.")
    # 程序自检，检查项目文件夹路径是否存在等（暂不开发）
    logger.info("Self Checking...")
    # 设备检测，检查系统分辨率是否符合最低要求
    logger.info("Device Checking...")
    device = check_device()
    if device.available is not True:
        # 不符合则弹窗提示后退出
        logger.error(f"Device is not available, {device}, Aborting...\n")
        pop_error_window(f"当前屏幕逻辑分辨率为{device.logic_w} x {device.logic_h}，"
                         f"当前缩放比例为 {int(device.scale_factor * 100)}%\n"
                         f"请在系统设置将屏幕缩放比例设置为100% ！\n"
                         f"请点击确定，程序会自动终止！", f"{get_current_time()} 设备不满足最低运行条件！")
        """
        pop_error_window(f"当前屏幕逻辑分辨率为{device.logic_w}x{device.logic_h}，最低要求为1280x768\n"
                         f"当前缩放比例为 {int(device.scale_factor * 100)}%，请在系统设置降低屏幕缩放比例！\n"
                         f"请点击确定，程序会自动终止！", f"{get_current_time()} 设备不满足最低运行条件！")
        """
        sys.exit()   # 分辨率不达标，程序退出
    elif device.available:
        # 若设备可运行则启动程序
        logger.info(f"Device Available, Starting program...")
        task_thread = threading.Thread(target=taskmgr.main, daemon=True)
        start_program()
