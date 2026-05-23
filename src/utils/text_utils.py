import re
import os
import chardet
from typing import Dict, List


def detect_file_encoding(file_path: str, config: Dict) -> str:
    """
    增强型文件编码检测（支持CSV/TXT/HTML）
    """
    import logging
    logger = logging.getLogger(__name__)
    
    try:
        with open(file_path, 'rb') as f:
            # 读取更多数据以提高检测准确性
            raw_data = f.read(2 * 1024 * 1024)  # 读取前2MB数据检测
            
            # 检查BOM（字节顺序标记）
            bom_encodings = {
                b'\xef\xbb\xbf': 'utf-8-sig',
                b'\xff\xfe': 'utf-16le',
                b'\xfe\xff': 'utf-16be',
                b'\xff\xfe\x00\x00': 'utf-32le',
                b'\x00\x00\xfe\xff': 'utf-32be'
            }
            
            for bom, enc in bom_encodings.items():
                if raw_data.startswith(bom):
                    logger.debug(f"编码检测：{os.path.basename(file_path)} 检测到BOM，使用编码 {enc}")
                    return enc
            
            # 使用chardet检测编码
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            confidence = result['confidence']
            
            logger.debug(f"编码检测：{os.path.basename(file_path)} chardet检测结果：{encoding} (置信度: {confidence})")

            # 如果置信度较高（>0.8），直接使用检测结果
            if encoding and confidence > 0.8:
                return encoding
        
        # 构建更全面的编码尝试列表
        preset_encodings = config["text_file_config"]["encodings"]
        # 添加更多常用编码，尤其是中文相关编码
        additional_encodings = [
            'utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'gb18030',
            'utf-16', 'utf-16le', 'utf-16be',
            'latin-1', 'iso-8859-1', 'cp1252',
            'big5', 'big5hkscs', 'shift_jis', 'euc-jp',
            'cp936', 'cp950', 'utf-32', 'utf-32le', 'utf-32be'
        ]
        
        # 合并预设编码和额外编码，并去重
        all_encodings = list(dict.fromkeys(preset_encodings + additional_encodings))
        
        # 优先尝试最常用的中文编码和UTF系列
        priority_encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030', 'gb2312']
        for enc in priority_encodings:
            if enc in all_encodings:
                all_encodings.remove(enc)
                all_encodings.insert(0, enc)
        
        # 尝试所有编码，使用更多数据验证
        for encoding in all_encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    # 尝试读取更多数据以确保编码正确
                    f.read(4096)  # 读取4KB数据验证
                logger.debug(f"编码检测：{os.path.basename(file_path)} 尝试编码 {encoding} 成功")
                return encoding
            except UnicodeDecodeError:
                continue
            except Exception as e:
                logger.debug(f"编码检测：{os.path.basename(file_path)} 尝试编码 {encoding} 失败：{str(e)}")
                continue
        
        logger.warning(f"检测文件编码失败，尝试多种编码均未成功，使用默认编码utf-8")
        return "utf-8"
    except Exception as e:
        logger.warning(f"检测文件编码失败，使用默认编码utf-8：{str(e)}")
        return "utf-8"


def detect_file_delimiter(file_path: str, encoding: str, config: Dict) -> str:
    """
    增强型文件分隔符检测（支持CSV/TXT）
    """
    try:
        # 读取更多行进行分隔符检测，提高准确性
        lines = []
        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            for i, line in enumerate(f):
                # 直接使用原始行，保留制表符和空格
                raw_line = line.rstrip('\n\r')  # 只去除换行符
                if raw_line:
                    lines.append(raw_line)
                    if len(lines) >= 30:  # 读取前30行，提高检测准确性
                        break
        
        if not lines:
            return config["text_file_config"]["delimiters"][0]  # 默认逗号
        
        # 首先检查文件中是否包含制表符
        tab_char = '\t'
        # 检查所有可能的制表符表示
        contains_tab = any(tab_char in line for line in lines)
        if not contains_tab:
            # 检查其他可能的制表符表示
            for line in lines:
                if '\x09' in line:
                    contains_tab = True
                    tab_char = '\x09'
                    break
        
        # 如果包含制表符，优先使用制表符作为分隔符
        if contains_tab:
            # 验证制表符是否能产生合理的列数
            sample_line = lines[0]
            tab_columns = sample_line.count(tab_char) + 1
            # 检查多行的一致性
            valid_lines = 0
            for line in lines[:10]:  # 检查前10行
                line_columns = line.count(tab_char) + 1
                if line_columns >= 3:
                    valid_lines += 1
            
            # 如果至少50%的行都能产生合理的列数（>=3列），则使用制表符
            if valid_lines >= len(lines[:10]) * 0.5 and tab_columns >= 3:
                return tab_char
            # 否则继续使用其他分隔符检测逻辑
        
        # 支持更多分隔符
        delimiters = config["text_file_config"]["delimiters"] + [',', '\t', ';', '|', r'\s+', '\x09', ' ', '\u0020', '\u00A0']
        # 去重
        delimiters = list(dict.fromkeys(delimiters))
        
        # 计算每个分隔符的得分
        delimiter_scores = {}
        
        for delimiter in delimiters:
            # 跳过空分隔符
            if not delimiter:
                continue
            
            # 计算该分隔符在所有行中的表现
            total_columns = 0
            valid_lines = 0
            consistent_lines = 0
            previous_columns = -1
            
            for line in lines:
                # 处理正则分隔符
                if delimiter in [r'\s+', ' ', '\u0020', '\u00A0']:
                    # 使用原始字符串避免无效转义序列
                    if delimiter == r'\s+':
                        delimiter = r'\s+'
                    columns = len(re.split(delimiter, line))
                else:
                    columns = line.count(delimiter) + 1
                
                # 只考虑有一定列数的行
                if columns >= 3:  # 至少3列才认为是有效数据行
                    valid_lines += 1
                    total_columns += columns
                    
                    # 检查与前一行的一致性
                    if previous_columns == -1:
                        previous_columns = columns
                        consistent_lines = 1
                    elif columns == previous_columns:
                        consistent_lines += 1
                    elif abs(columns - previous_columns) <= 1:
                        # 允许±1的差异
                        consistent_lines += 0.5
                    
            if valid_lines > 0:
                # 计算平均列数
                avg_columns = total_columns / valid_lines
                
                # 计算一致性得分（0-1）
                consistency_score = consistent_lines / valid_lines
                
                # 综合得分：平均列数 * 一致性得分
                # 优先选择列数较多且一致性较好的分隔符
                score = avg_columns * consistency_score
                
                delimiter_scores[delimiter] = score
            else:
                delimiter_scores[delimiter] = 0
        
        # 选择得分最高的分隔符
        best_delimiter = max(delimiter_scores, key=delimiter_scores.get)
        
        # 如果最佳分隔符的得分较低，使用默认分隔符
        if delimiter_scores[best_delimiter] < 2:  # 得分低于2，使用默认分隔符
            return config["text_file_config"]["delimiters"][0]
        
        return best_delimiter
    except Exception as e:
        return config["text_file_config"]["delimiters"][0]


def preprocess_txt_file(file_path: str, encoding: str, config: Dict) -> str:
    """
    TXT文件预处理，生成临时规整文件
    """
    import tempfile
    import os
    
    file_name = os.path.basename(file_path)
    temp_fd, temp_path = tempfile.mkstemp(suffix='.txt', prefix='temp_bank_txt_')
    os.close(temp_fd)

    try:
        with open(file_path, 'rb') as f_in:
            raw_data = f_in.read()
        
        # 统一换行符为\n
        # 处理不同的换行符格式
        raw_data = raw_data.replace(b'\r\n', b'\n')  # Windows格式
        raw_data = raw_data.replace(b'\r', b'\n')    # 旧Mac格式
        
        # 解码文件内容
        try:
            content = raw_data.decode(encoding, errors='ignore')
        except Exception:
            # 如果解码失败，尝试其他编码
            content = raw_data.decode('utf-8', errors='ignore')
        
        # 写入临时文件
        with open(temp_path, 'w', encoding='utf-8') as f_out:
            lines = content.split('\n')
            for line_num, line in enumerate(lines):
                # 跳过空白行
                if config["text_file_config"]["skip_blank_lines"] and not line.strip():
                    continue
                # 跳过指定行数（配置项）
                if line_num < config["text_file_config"]["skip_rows"]:
                    continue
                # 只去除不可见字符，保留制表符、空格和换行符
                # 注意：制表符(\t)的ASCII码是9，属于\u0000-\u001F范围，需要特别保留
                clean_line = re.sub(r'[\u0000-\u0008\u000a-\u001F\u007F-\u009F\u200b]', '', line)
                if clean_line:
                    f_out.write(clean_line + '\n')
        return temp_path
    except Exception as e:
        os.unlink(temp_path)
        return file_path
