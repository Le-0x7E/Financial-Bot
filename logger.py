import logging
from colorama import Fore, Style, init

# 初始化 colorama
init(autoreset=True)


# 自定义颜色格式器
class ColoredFormatter(logging.Formatter):
    # 定义日志级别对应的颜色
    COLORS = {
        logging.DEBUG: Fore.CYAN,
        logging.INFO: Fore.WHITE,
        logging.WARNING: Fore.YELLOW,
        logging.ERROR: Fore.RED,
        logging.CRITICAL: Fore.RED + Style.BRIGHT,
    }

    def format(self, record):
        # 获取日志级别对应的颜色
        color = self.COLORS.get(record.levelno, Fore.WHITE)
        # 设置日志格式
        formatter = logging.Formatter(
            f"{color}[%(levelname)s][%(asctime)s] - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        return formatter.format(record)


# 创建 logger
logger = logging.getLogger('my_app_logger')
logger.setLevel(logging.DEBUG)  # 设置日志级别

# 创建一个控制台处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)  # 控制台只输出 INFO 及以上级别的日志
console_handler.setFormatter(ColoredFormatter())  # 使用自定义颜色格式器

# 创建一个文件处理器
file_handler = logging.FileHandler('run.log', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)  # 文件记录 DEBUG 及以上级别的日志
file_handler.setFormatter(
    logging.Formatter(
        "[%(levelname)s][%(asctime)s][%(filename)s:%(lineno)d] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
)

# 将处理器添加到 logger
logger.addHandler(console_handler)
logger.addHandler(file_handler)

# 禁止根记录器输出日志
logging.getLogger().setLevel(logging.CRITICAL)
