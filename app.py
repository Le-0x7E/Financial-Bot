from controlpanel import *


def start_program():
    # 程序自检，检查文件夹路径是否存在等
    logger.info("Starting program...")
    # 环境检测，检查系统分辨率是否符合最低要求
    logger.info("Checking device...")
    device = device_check()
    print(config.scale)
    if device.available is not True:
        logger.error(f"Device is not available, {device}\n")
        pop_error_window(f"当前屏幕逻辑分辨率为{device.logic_w}x{device.logic_h}，最低要求为1280x768\n"
                         f"当前缩放比例为 {int(device.scale_factor * 100)}%，请在系统设置降低屏幕缩放比例！\n"
                         f"请点击确定，程序会自动终止！", f"{get_current_time()} 设备不满足最低运行条件！")
        sys.exit()
    # 设置全局变量
    config.scale = device.scale_factor
    config.img_dir = set_img_dir(step='shuiwu')
    # 启动程序控制窗口
    logger.info("Starting Control Panel...")
    start_control_panel()


if __name__ == "__main__":
    start_program()
