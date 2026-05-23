# 银行流水汇总程序

这是一个用于汇总和处理银行流水的Python程序。

## 功能说明

- 支持多种银行格式的流水文件解析
- 自动识别和处理多个账户
- 生成汇总报告
- 支持PDF和Excel格式的流水文件

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

运行主程序：

```bash
python bank_statement_aggregator.py
```

## 项目结构

```
银行流水汇总程序1.0/
├── bank_statement_aggregator.py  # 主程序入口
├── config.yaml                   # 配置文件
├── update_mappings.py            # 映射更新脚本
├── requirements.txt              # 依赖列表
├── CODE_WIKI.md                  # 代码文档
├── OPTIMIZATION_SUGGESTIONS.md   # 优化建议
├── src/
│   ├── config/                   # 配置模块
│   ├── core/                     # 核心功能模块
│   ├── file_processing/          # 文件处理模块
│   ├── logging/                  # 日志模块
│   └── utils/                    # 工具模块
└── README.md                     # 本文件
```
