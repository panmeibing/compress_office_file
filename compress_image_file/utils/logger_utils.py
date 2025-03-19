import logging
import os

from utils.singleton_utils import singleton


@singleton
def get_logger(save_path: str, file_name: str = "uname"):
    logger = logging.getLogger("MultiDownloader")
    logger.setLevel(logging.INFO)
    print_fmt = logging.Formatter("%(asctime)s-%(levelname)s-%(process)s: %(message)s")
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(print_fmt)
    if not os.path.exists(save_path):
        os.mkdir(save_path)
    logger_name = file_name if file_name and file_name.endswith(".log") else f"{file_name}.log"
    file_handler = logging.FileHandler(os.path.join(save_path, logger_name), encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_fmt = logging.Formatter("%(asctime)s-%(filename)s-line:%(lineno)d-%(levelname)s-%(process)s: %(message)s")
    file_handler.setFormatter(file_fmt)
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    return logger
