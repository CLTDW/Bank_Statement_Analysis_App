import pdfplumber
import pandas as pd
import logging
import re
from typing import List, Dict, Optional

logger = logging.getLogger('bank_statement_aggregator')
# 设置PDF处理模块的日志级别为WARNING
logger.setLevel(logging.WARNING)

def process_pdf_file(file_path: str, config: Dict) -> Optional[pd.DataFrame]:
    """
    处理PDF文件，提取表格数据并进行标准化
    
    Args:
        file_path: PDF文件路径
        config: 配置字典
        
    Returns:
        处理后的DataFrame或None
    """
    try:
        logger.info(f"【PDF处理】开始处理PDF文件：{file_path}")
        
        # 打开PDF文件
        with pdfplumber.open(file_path) as pdf:
            all_tables = []
            
            # 遍历所有页面
            for page_num, page in enumerate(pdf.pages, 1):
                logger.info(f"【PDF处理】处理第{page_num}页")
                
                # 提取文本用于银行类型检测
                page_text = page.extract_text()
                bank_type = detect_bank_type(page_text)
                logger.info(f"【PDF处理】检测到银行类型：{bank_type}")
                
                # 多策略表格提取
                table_extracted = False
                
                # 策略1：默认表格提取
                tables = page.extract_tables()
                logger.info(f"【PDF处理】第{page_num}页默认设置提取到{len(tables)}个表格")
                
                for table_idx, table in enumerate(tables):
                    if table and len(table) > 1:
                        try:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            logger.info(f"【PDF处理】第{page_num}页表格{table_idx+1}：{len(df)}行数据")
                            all_tables.append(df)
                            table_extracted = True
                        except Exception as e:
                            logger.debug(f"【PDF处理】第{page_num}页表格{table_idx+1}转换失败：{str(e)}")
                            try:
                                df = pd.DataFrame(table[1:])
                                logger.info(f"【PDF处理】第{page_num}页表格{table_idx+1}（使用默认列名）：{len(df)}行数据")
                                all_tables.append(df)
                                table_extracted = True
                            except Exception as e2:
                                logger.debug(f"【PDF处理】第{page_num}页表格{table_idx+1}完全失败：{str(e2)}")
                
                # 策略2：使用自定义表格设置
                if not table_extracted:
                    logger.info(f"【PDF处理】第{page_num}页使用自定义表格设置")
                    # 为不同银行类型使用不同的表格设置
                    table_settings = get_table_settings(bank_type)
                    tables = page.extract_tables(table_settings=table_settings)
                    logger.info(f"【PDF处理】第{page_num}页使用自定义设置提取到{len(tables)}个表格")
                    
                    for table_idx, table in enumerate(tables):
                        if table and len(table) > 1:
                            try:
                                df = pd.DataFrame(table[1:], columns=table[0])
                                logger.info(f"【PDF处理】第{page_num}页表格{table_idx+1}：{len(df)}行数据")
                                all_tables.append(df)
                                table_extracted = True
                            except Exception as e:
                                logger.debug(f"【PDF处理】第{page_num}页表格{table_idx+1}转换失败：{str(e)}")
                                try:
                                    df = pd.DataFrame(table[1:])
                                    logger.info(f"【PDF处理】第{page_num}页表格{table_idx+1}（使用默认列名）：{len(df)}行数据")
                                    all_tables.append(df)
                                    table_extracted = True
                                except Exception as e2:
                                    logger.debug(f"【PDF处理】第{page_num}页表格{table_idx+1}完全失败：{str(e2)}")
                
                # 策略3：使用更宽松的表格设置
                if not table_extracted:
                    logger.info(f"【PDF处理】第{page_num}页使用宽松表格设置")
                    loose_settings = {
                        'vertical_strategy': 'text',
                        'horizontal_strategy': 'text',
                        'snap_tolerance': 10,
                        'join_tolerance': 5,
                        'edge_min_length': 3,
                        'min_words_vertical': 2,
                        'min_words_horizontal': 1,
                        'text_tolerance': 5,
                        'text_x_tolerance': 5,
                        'text_y_tolerance': 5,
                        'intersection_tolerance': 5
                    }
                    tables = page.extract_tables(table_settings=loose_settings)
                    logger.info(f"【PDF处理】第{page_num}页使用宽松设置提取到{len(tables)}个表格")
                    
                    for table_idx, table in enumerate(tables):
                        if table and len(table) > 1:
                            try:
                                df = pd.DataFrame(table[1:], columns=table[0])
                                logger.info(f"【PDF处理】第{page_num}页表格{table_idx+1}：{len(df)}行数据")
                                all_tables.append(df)
                                table_extracted = True
                            except Exception as e:
                                logger.debug(f"【PDF处理】第{page_num}页表格{table_idx+1}转换失败：{str(e)}")
                                try:
                                    df = pd.DataFrame(table[1:])
                                    logger.info(f"【PDF处理】第{page_num}页表格{table_idx+1}（使用默认列名）：{len(df)}行数据")
                                    all_tables.append(df)
                                    table_extracted = True
                                except Exception as e2:
                                    logger.debug(f"【PDF处理】第{page_num}页表格{table_idx+1}完全失败：{str(e2)}")
                
                # 策略4：从文本中提取数据
                if not table_extracted:
                    logger.info(f"【PDF处理】第{page_num}页尝试从文本中提取数据")
                    if page_text:
                        extracted_data = extract_data_from_text(page_text, bank_type)
                        if extracted_data:
                            df = pd.DataFrame(extracted_data)
                            logger.info(f"【PDF处理】第{page_num}页从文本中提取到{len(df)}行数据")
                            all_tables.append(df)
                            table_extracted = True
            
            if not all_tables:
                logger.warning("【PDF处理】未提取到有效表格")
                return None
            
            # 合并所有表格
            try:
                combined_df = pd.DataFrame()
                
                for i, table_df in enumerate(all_tables):
                    try:
                        table_df['_table_id'] = i
                        table_df = table_df.reset_index(drop=True)
                        combined_df = pd.concat([combined_df, table_df], ignore_index=True)
                    except Exception as e2:
                        logger.debug(f"【PDF处理】跳过有问题的表格{i+1}：{str(e2)}")
                
                if not combined_df.empty:
                    logger.info(f"【PDF处理】合并后共{len(combined_df)}行数据")
                else:
                    logger.warning("【PDF处理】所有表格都有问题，无法合并")
                    return None
            except Exception as e:
                logger.error(f"【PDF处理】合并表格失败：{str(e)}")
                if all_tables:
                    combined_df = all_tables[0].copy()
                    logger.info(f"【PDF处理】使用第一个表格，共{len(combined_df)}行数据")
                else:
                    logger.warning("【PDF处理】未提取到有效表格")
                    return None
            
            # 清理和标准化数据
            cleaned_df = clean_pdf_data(combined_df)
            
            if cleaned_df.empty:
                logger.warning("【PDF处理】清理后无有效数据")
                return None
            
            return cleaned_df
            
    except Exception as e:
        logger.error(f"【PDF处理异常】{str(e)}", exc_info=True)
        return None

def clean_pdf_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    清理和标准化PDF提取的数据
    
    Args:
        df: 原始提取的DataFrame
        
    Returns:
        清理后的DataFrame
    """
    # 移除空行
    df = df.dropna(how='all')
    
    # 移除空列
    df = df.dropna(axis=1, how='all')
    
    # 标准化列名
    if not df.empty:
        # 如果列名是数字（默认列名），尝试从第一行获取列名
        if all(isinstance(col, int) for col in df.columns):
            try:
                # 使用第一行作为列名
                new_columns = [str(df.iloc[0, i]).strip() if pd.notna(df.iloc[0, i]) else f'列{i+1}' for i in range(len(df.columns))]
                df.columns = new_columns
                # 移除第一行（已作为列名）
                df = df.iloc[1:].reset_index(drop=True)
            except Exception as e:
                logger.debug(f"【PDF处理】无法从第一行获取列名：{str(e)}")
                # 使用默认列名
                df.columns = [f'列{i+1}' for i in range(len(df.columns))]
        else:
            # 标准化现有列名
            df.columns = [str(col).strip() if col is not None else f'列{i+1}' for i, col in enumerate(df.columns)]
    
    # 清理单元格数据
    for col in df.columns:
        df[col] = df[col].apply(lambda x: str(x).strip() if x is not None else '')
    
    # 移除完全重复的行
    df = df.drop_duplicates()
    
    return df

def detect_bank_type(pdf_text: str) -> str:
    """
    根据PDF文本检测银行类型
    
    Args:
        pdf_text: PDF文本内容
        
    Returns:
        银行类型
    """
    bank_keywords = {
        '中国工商银行': ['工商银行', 'ICBC'],
        '中国建设银行': ['建设银行', 'CCB'],
        '中国农业银行': ['农业银行', 'ABC'],
        '中国银行': ['中国银行', 'BOC'],
        '交通银行': ['交通银行', 'BOCOM'],
        '招商银行': ['招商银行', 'CMB'],
        '中信银行': ['中信银行', 'CITIC'],
        '浦发银行': ['浦发银行', 'SPDB'],
        '民生银行': ['民生银行', 'CMBC'],
        '平安银行': ['平安银行', 'PING AN'],
        '微信支付': ['微信支付', 'WeChat Pay'],
        '支付宝': ['支付宝', 'Alipay']
    }
    
    for bank, keywords in bank_keywords.items():
        for keyword in keywords:
            if keyword in pdf_text:
                return bank
    
    return '未知银行'

def get_table_settings(bank_type: str) -> Dict:
    """
    根据银行类型返回对应的表格提取设置
    
    Args:
        bank_type: 银行类型
        
    Returns:
        表格提取设置字典
    """
    # 针对不同银行的表格设置
    bank_settings = {
        '中国工商银行': {
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'snap_tolerance': 4,
            'join_tolerance': 2,
            'edge_min_length': 4,
            'min_words_vertical': 3,
            'min_words_horizontal': 1,
            'text_tolerance': 2,
            'text_x_tolerance': 2,
            'text_y_tolerance': 2,
            'intersection_tolerance': 2
        },
        '中国建设银行': {
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'snap_tolerance': 5,
            'join_tolerance': 3,
            'edge_min_length': 5,
            'min_words_vertical': 3,
            'min_words_horizontal': 1,
            'text_tolerance': 3,
            'text_x_tolerance': 3,
            'text_y_tolerance': 3,
            'intersection_tolerance': 3
        },
        '中国农业银行': {
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'snap_tolerance': 6,
            'join_tolerance': 4,
            'edge_min_length': 6,
            'min_words_vertical': 3,
            'min_words_horizontal': 1,
            'text_tolerance': 4,
            'text_x_tolerance': 4,
            'text_y_tolerance': 4,
            'intersection_tolerance': 4
        },
        '中国银行': {
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'snap_tolerance': 5,
            'join_tolerance': 3,
            'edge_min_length': 5,
            'min_words_vertical': 3,
            'min_words_horizontal': 1,
            'text_tolerance': 3,
            'text_x_tolerance': 3,
            'text_y_tolerance': 3,
            'intersection_tolerance': 3
        },
        '微信支付': {
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'snap_tolerance': 3,
            'join_tolerance': 2,
            'edge_min_length': 3,
            'min_words_vertical': 2,
            'min_words_horizontal': 1,
            'text_tolerance': 2,
            'text_x_tolerance': 2,
            'text_y_tolerance': 2,
            'intersection_tolerance': 2
        },
        '支付宝': {
            'vertical_strategy': 'text',
            'horizontal_strategy': 'text',
            'snap_tolerance': 3,
            'join_tolerance': 2,
            'edge_min_length': 3,
            'min_words_vertical': 2,
            'min_words_horizontal': 1,
            'text_tolerance': 2,
            'text_x_tolerance': 2,
            'text_y_tolerance': 2,
            'intersection_tolerance': 2
        }
    }
    
    # 如果没有对应银行的设置，返回默认设置
    return bank_settings.get(bank_type, {
        'vertical_strategy': 'text',
        'horizontal_strategy': 'text',
        'snap_tolerance': 5,
        'join_tolerance': 3,
        'edge_min_length': 5,
        'min_words_vertical': 3,
        'min_words_horizontal': 1,
        'text_tolerance': 3,
        'text_x_tolerance': 3,
        'text_y_tolerance': 3,
        'intersection_tolerance': 3
    })

def extract_data_from_text(text: str, bank_type: str = '未知银行') -> List[Dict[str, str]]:
    """
    从文本中提取银行流水数据
    
    Args:
        text: PDF页面文本内容
        bank_type: 银行类型
        
    Returns:
        提取的数据列表
    """
    data = []
    
    # 分割文本行
    lines = text.strip().split('\n')
    
    # 查找表头行
    header_line = None
    header_indices = {}
    
    # 常见的银行流水表头关键词
    header_keywords = ['交易时间', '流水号', '摘要', '交易金额', '账户余额', '对方户名', '对方账号', '对手行名', '交易渠道', '交易机构']
    
    # 查找包含表头的行
    for i, line in enumerate(lines):
        if any(keyword in line for keyword in header_keywords):
            header_line = line
            # 简单的表头解析（根据关键词位置）
            for keyword in header_keywords:
                if keyword in line:
                    header_indices[keyword] = line.index(keyword)
            break
    
    if header_line:
        # 处理数据行
        for line in lines:
            # 跳过表头行和空行
            if line == header_line or not line.strip():
                continue
            
            # 尝试解析数据行
            try:
                # 基于关键词位置提取数据
                row_data = {}
                
                # 提取交易时间
                if '交易时间' in header_indices:
                    start = header_indices['交易时间']
                    # 找到下一个关键词的位置作为结束
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '交易时间' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    transaction_time = line[start:end].strip()
                    if transaction_time:
                        row_data['交易时间'] = transaction_time
                
                # 提取流水号
                if '流水号' in header_indices:
                    start = header_indices['流水号']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '流水号' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    serial_no = line[start:end].strip()
                    if serial_no:
                        row_data['流水号'] = serial_no
                
                # 提取摘要
                if '摘要' in header_indices:
                    start = header_indices['摘要']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '摘要' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    summary = line[start:end].strip()
                    if summary:
                        row_data['摘要'] = summary
                
                # 提取交易金额
                if '交易金额' in header_indices:
                    start = header_indices['交易金额']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '交易金额' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    amount = line[start:end].strip()
                    if amount:
                        row_data['交易金额'] = amount
                
                # 提取账户余额
                if '账户余额' in header_indices:
                    start = header_indices['账户余额']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '账户余额' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    balance = line[start:end].strip()
                    if balance:
                        row_data['账户余额'] = balance
                
                # 提取对方户名
                if '对方户名' in header_indices:
                    start = header_indices['对方户名']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '对方户名' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    counterparty_name = line[start:end].strip()
                    if counterparty_name:
                        row_data['对方户名'] = counterparty_name
                
                # 提取对方账号
                if '对方账号' in header_indices:
                    start = header_indices['对方账号']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '对方账号' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    counterparty_account = line[start:end].strip()
                    if counterparty_account:
                        row_data['对方账号'] = counterparty_account
                
                # 提取对手行名
                if '对手行名' in header_indices:
                    start = header_indices['对手行名']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '对手行名' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    counterparty_bank = line[start:end].strip()
                    if counterparty_bank:
                        row_data['对手行名'] = counterparty_bank
                
                # 提取交易渠道
                if '交易渠道' in header_indices:
                    start = header_indices['交易渠道']
                    end = len(line)
                    for keyword in header_keywords:
                        if keyword != '交易渠道' and keyword in header_indices:
                            keyword_pos = header_indices[keyword]
                            if keyword_pos > start:
                                end = keyword_pos
                                break
                    channel = line[start:end].strip()
                    if channel:
                        row_data['交易渠道'] = channel
                
                # 提取交易机构
                if '交易机构' in header_indices:
                    start = header_indices['交易机构']
                    end = len(line)
                    institution = line[start:end].strip()
                    if institution:
                        row_data['交易机构'] = institution
                
                # 只有当提取到足够的数据时才添加
                if len(row_data) >= 3:
                    data.append(row_data)
            except Exception as e:
                logger.debug(f"【PDF处理】解析文本行失败：{str(e)}")
                continue
    
    # 如果没有通过表头提取到数据，尝试使用正则表达式提取
    if not data:
        # 针对不同银行使用不同的正则模式
        if bank_type == '微信支付':
            # 微信支付的特殊格式
            pattern = r'(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})\s+(.+?)\s+([+-]?\d+\.\d{2})'
            matches = re.findall(pattern, text)
            for match in matches:
                if len(match) >= 4:
                    row_data = {
                        '交易时间': f"{match[0]} {match[1]}",
                        '摘要': match[2].strip(),
                        '交易金额': match[3].strip()
                    }
                    data.append(row_data)
        elif bank_type == '支付宝':
            # 支付宝的特殊格式
            pattern = r'(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2}:\d{2})\s+(.+?)\s+([+-]?\d+\.\d{2})'
            matches = re.findall(pattern, text)
            for match in matches:
                if len(match) >= 4:
                    row_data = {
                        '交易时间': f"{match[0]} {match[1]}",
                        '摘要': match[2].strip(),
                        '交易金额': match[3].strip()
                    }
                    data.append(row_data)
        else:
            # 通用银行流水格式
            # 模式1：序号 交易时间 流水号 摘要 交易金额 账户余额 ...
            pattern1 = r'\s*\d+\s+([\d/\-\s:]+)\s+([\d]+)\s+([^\d]+)\s+([\d\.]+)\s+([\d\.]+)'
            matches = re.findall(pattern1, text)
            
            for match in matches:
                if len(match) >= 5:
                    row_data = {
                        '交易时间': match[0].strip(),
                        '流水号': match[1].strip(),
                        '摘要': match[2].strip(),
                        '交易金额': match[3].strip(),
                        '账户余额': match[4].strip()
                    }
                    data.append(row_data)
            
            # 模式2：针对特定格式的银行流水
            if not data:
                # 尝试匹配更具体的格式，如示例中的格式
                lines = text.strip().split('\n')
                in_data_section = False
                for line in lines:
                    # 查找数据行的开始
                    if '序号' in line and '交易时间' in line:
                        in_data_section = True
                        continue
                    
                    if in_data_section:
                        # 跳过空行和表头行
                        if not line.strip() or '序号' in line:
                            continue
                        
                        # 尝试按空格分割数据
                        parts = line.strip().split()
                        if len(parts) >= 6:
                            # 序号 交易时间 流水号 摘要 交易金额 账户余额 ...
                            try:
                                # 处理交易时间（可能包含空格）
                                transaction_time = ' '.join(parts[1:3])  # 例如：2023/11/04 10:05:27
                                serial_no = parts[3]
                                summary = parts[4]
                                amount = parts[5]
                                balance = parts[6]
                                
                                row_data = {
                                    '交易时间': transaction_time,
                                    '流水号': serial_no,
                                    '摘要': summary,
                                    '交易金额': amount,
                                    '账户余额': balance
                                }
                                
                                # 提取更多字段
                                if len(parts) >= 7:
                                    # 对方户名
                                    if len(parts) >= 8:
                                        counterparty_name = parts[7]
                                        row_data['对方户名'] = counterparty_name
                                    # 对方账号
                                    if len(parts) >= 9:
                                        counterparty_account = parts[8]
                                        row_data['对方账号'] = counterparty_account
                                    # 对手行名
                                    if len(parts) >= 10:
                                        counterparty_bank = parts[9]
                                        row_data['对手行名'] = counterparty_bank
                                    # 交易渠道
                                    if len(parts) >= 11:
                                        channel = parts[10]
                                        row_data['交易渠道'] = channel
                                    # 交易机构
                                    if len(parts) >= 12:
                                        institution = ' '.join(parts[11:])
                                        row_data['交易机构'] = institution
                                
                                data.append(row_data)
                            except Exception as e:
                                logger.debug(f"【PDF处理】解析数据行失败：{str(e)}")
                                continue
    
    # 提取账户信息
    account_info = extract_account_info(text)
    if account_info:
        # 将账户信息添加到每一行数据中
        for row in data:
            for key, value in account_info.items():
                if key not in row:
                    row[key] = value
    
    return data

def extract_account_info(pdf_text: str) -> Dict[str, str]:
    """
    从PDF文本中提取账户信息
    
    Args:
        pdf_text: PDF文本内容
        
    Returns:
        账户信息字典
    """
    account_info = {}
    
    # 提取账号
    account_match = re.search(r'账号[:：]\s*(\d+)', pdf_text)
    if account_match:
        account_info['本方账号'] = account_match.group(1)
    
    # 提取户名
    name_match = re.search(r'户名[:：]\s*(.+?)[\n\r]', pdf_text)
    if name_match:
        account_info['户主名称'] = name_match.group(1).strip()
    
    # 提取开户行
    bank_match = re.search(r'开户行[:：]\s*(.+?)[\n\r]', pdf_text)
    if bank_match:
        account_info['本方账户开户行'] = bank_match.group(1).strip()
    
    # 提取币种
    currency_match = re.search(r'币种[:：]\s*(.+?)[\n\r]', pdf_text)
    if currency_match:
        account_info['币种'] = currency_match.group(1).strip()
    
    return account_info