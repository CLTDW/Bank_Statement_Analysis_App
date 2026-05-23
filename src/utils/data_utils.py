import pandas as pd
import numpy as np
import re
from typing import Dict, Union, Optional


def optimize_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    优化列名去重：保留非空值占比最高的列
    """
    if not df.columns.duplicated().any():
        return df

    new_columns = []
    seen = {}
    for col in df.columns:
        if col not in seen:
            seen[col] = [col]
            new_columns.append(col)
        else:
            seen[col].append(col + f"_dup_{len(seen[col])}")
            new_columns.append(seen[col][-1])

    df.columns = new_columns
    # 合并重复列（保留非空值多的）
    for orig_col, dup_cols in seen.items():
        if len(dup_cols) > 1:
            # 计算各重复列的非空占比
            ratios = {}
            for col in dup_cols:
                try:
                    # 简化计算，直接统计非空值数量
                    ratios[col] = df[col].notna().sum()
                except Exception as e:
                    ratios[col] = 0
            
            best_col = max(ratios, key=ratios.get)
            # 将其他列的非空值填充到最佳列
            for col in dup_cols:
                if col != best_col:
                    try:
                        df[best_col] = df[best_col].fillna(df[col])
                        df = df.drop(col, axis=1)
                    except Exception as e:
                        pass
    return df


def standardize_header(header: Union[str, None]) -> str:
    """
    表头标准化
    
    Args:
        header: 原始表头值
        
    Returns:
        标准化后的表头值
    """
    if pd.isna(header) or header is None:
        return ""
    # 1. 强制转字符串，去除首尾空白
    header_str = str(header).strip()
    if not header_str:
        return ""
    # 2. 去除不可见字符
    invisible_chars = r'[\u0000-\u001F\u007F-\u009F\u00A0\u200b\u2028\u2029]'
    header_str = re.sub(invisible_chars, '', header_str)
    # 3. 特殊处理：合并表格、制表符、换行符
    # 去除换行符、制表符、回车符
    header_str = re.sub(r'[\n\t\r]', '', header_str)
    # 合并多余空格
    header_str = re.sub(r'\s{2,}', ' ', header_str)
    # 处理合并表格产生的特殊标记
    header_str = re.sub(r'\[合并\]|\(合并\)|合并表格', '', header_str)

    # 4. 全半角统一
    def full2half(char):
        code = ord(char)
        if code == 0x3000:
            return chr(0x20)
        elif 0xFF01 <= code <= 0xFF5E:
            return chr(code - 0xFEE0)
        else:
            return char

    header_str = ''.join([full2half(c) for c in header_str])
    
    # 5. MAC地址和IP地址标准化：将字母转换为大写
    # 检测是否包含MAC地址或IP地址相关内容
    mac_ip_patterns = [
        r'MAC', r'IP', r'ip', r'mac'
    ]
    
    # 检查表头是否包含MAC或IP相关内容
    contains_mac_ip = any(pattern in header_str for pattern in mac_ip_patterns)
    
    if contains_mac_ip:
        # 将字母转换为大写
        header_str = header_str.upper()
    
    # 6. 保留有意义的空白字符，只去除多余空格
    header_str = re.sub(r'\s{2,}', ' ', header_str)
    header_str = header_str.strip()
    return header_str









