"""
日志系统配置模块
功能：配置日志系统，将所有日志输出到logs目录
"""

import os
import logging
from pathlib import Path
from datetime import datetime
from logging.handlers import RotatingFileHandler


def setup_logger(name: str = "document_converter", log_dir: str = "logs") -> logging.Logger:
    """
    设置日志系统
    
    Args:
        name: 日志记录器名称
        log_dir: 日志目录
        
    Returns:
        配置好的日志记录器
    """
    # 创建logs目录
    log_path = Path(log_dir)
    log_path.mkdir(exist_ok=True)
    
    # 创建日志记录器
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # 避免重复添加处理器
    if logger.handlers:
        return logger
    
    # 日志格式
    log_format = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # 文件处理器 - 按日期命名
    log_filename = log_path / f"converter_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = RotatingFileHandler(
        log_filename,
        maxBytes=10 * 1024 * 1024,  # 10MB
        backupCount=5,  # 保留5个备份文件
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(log_format)
    
    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(log_format)
    
    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


def get_logger(name: str = "document_converter") -> logging.Logger:
    """
    获取日志记录器（如果已存在则返回，否则创建新的）
    
    Args:
        name: 日志记录器名称
        
    Returns:
        日志记录器
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        return setup_logger(name)
    return logger

