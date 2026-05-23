#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
银行流水汇总程序主入口
"""
import argparse
import os
import sys
from pathlib import Path

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import load_config, init_bank_alias_mapping
from src.logging import setup_logger
from src.core.bank_statement_aggregator import BankStatementAggregator


def main():
    """
    程序主入口
    """
    # 命令行参数解析
    parser = argparse.ArgumentParser(description='银行流水处理程序')
    parser.add_argument('-c', '--config', type=str, default='config.yaml', help='配置文件路径')
    parser.add_argument('-d', '--directory', type=str, help='指定流水文件目录，覆盖配置文件中的base_dir')
    parser.add_argument('-o', '--output', type=str, help='指定输出文件路径，覆盖配置文件中的output_file')
    parser.add_argument('--disable-parallel', action='store_true', help='禁用并行处理')
    args = parser.parse_args()

    # 加载配置
    CONFIG = load_config(args.config)

    # 应用命令行参数覆盖
    if args.directory:
        CONFIG['base_dir'] = args.directory
    if args.output:
        CONFIG['output_file'] = args.output
    if args.disable_parallel:
        CONFIG['parallel']['enabled'] = False

    # 执行银行标准化初始化
    init_bank_alias_mapping(CONFIG)

    # 初始化日志
    logger = setup_logger(CONFIG)

    # 初始化并运行银行流水汇总程序
    aggregator = BankStatementAggregator(CONFIG)
    aggregator.run()


if __name__ == "__main__":
    main()