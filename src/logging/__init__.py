import logging
import os
from pathlib import Path


def setup_logger(config):
    """
    初始化日志记录器
    
    Args:
        config: 配置字典
        
    Returns:
        日志记录器实例
    """
    # 确保日志目录存在
    log_dir = Path(config.get('log_dir', 'logs'))
    log_dir.mkdir(exist_ok=True)
    
    # 确保TEST文件夹存在
    test_dir = Path('TEST')
    test_dir.mkdir(exist_ok=True)
    
    # 配置日志格式
    log_format = '%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s'
    log_level = getattr(logging, config.get('log_level', 'DEBUG').upper(), logging.DEBUG)
    
    # 创建日志记录器
    logger = logging.getLogger('bank_statement_aggregator')
    logger.setLevel(log_level)
    
    # 清除现有处理器
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_formatter = logging.Formatter(log_format)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # 创建文件处理器
    log_file = log_dir / f"bank_statement_aggregator_{os.getpid()}.log"
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(log_level)
    file_formatter = logging.Formatter(log_format)
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)
    
    # 在TEST文件夹中也创建日志文件
    test_log_file = test_dir / f"bank_statement_aggregator_{os.getpid()}.log"
    test_file_handler = logging.FileHandler(test_log_file, encoding='utf-8')
    test_file_handler.setLevel(log_level)
    test_file_formatter = logging.Formatter(log_format)
    test_file_handler.setFormatter(test_file_formatter)
    logger.addHandler(test_file_handler)
    
    return logger
