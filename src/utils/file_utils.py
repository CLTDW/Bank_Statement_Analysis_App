import contextlib
import tempfile
import os
import fnmatch
import re
from typing import List, Dict
import msoffcrypto

# 尝试导入python-magic
magic_available = False
try:
    import magic
    magic_available = True
except ImportError:
    pass


@contextlib.contextmanager
def temp_excel_file(file_path: str = '', mode: str = 'wb'):
    """
    安全的临时文件上下文管理器，自动清理
    """
    temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx', prefix='temp_bank_')
    os.close(temp_fd)
    try:
        yield temp_path
    finally:
        if os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except Exception as e:
                pass





def is_temp_or_invalid_file(file_name: str, config: Dict) -> bool:
    """
    检测是否为临时文件/无效文件
    """
    # 匹配排除模式
    for pattern in config["excluded_patterns"]:
        if fnmatch.fnmatch(file_name, pattern):
            return True
    # 检测隐藏文件
    if file_name.startswith('.'):
        return True
    # 检测大小异常的空文件
    return False


def collect_all_supported_files(base_dir: str, config: Dict) -> List[str]:
    """
    遍历base_dir及其所有子文件夹，收集所有支持格式的文件
    """
    supported_files = []
    filtered_files = {
        "skipped_folders": 0,
        "excluded_files": 0,
        "temp_invalid_files": 0,
        "summary_result_files": 0,
        "unsupported_formats": 0
    }
    
    total_files = 0
    
    for root, dirs, files in os.walk(base_dir):
        # 跳过"原始文件备份"和"待复核文件"文件夹
        if "原始文件备份" in root or "待复核文件" in root:
            filtered_files["skipped_folders"] += len(files)
            continue
        
        # 从dirs中移除"原始文件备份"和"待复核文件"，避免进入该文件夹
        dirs[:] = [d for d in dirs if d != "原始文件备份" and d != "待复核文件"]
        
        total_files += len(files)
        
        for file_name in files:
            full_file_path = os.path.join(root, file_name)
            
            # 过滤排除文件
            if file_name in config["excluded_files"]:
                filtered_files["excluded_files"] += 1
                continue
            # 过滤临时/无效文件
            if is_temp_or_invalid_file(file_name, config):
                filtered_files["temp_invalid_files"] += 1
                continue
            # 过滤汇总结果文件和报告文件，避免重复处理
            if file_name.startswith("银行流水文件汇总结果") or "银行流水汇总结果" in file_name or "银行流水汇总情况报告" in file_name:
                filtered_files["summary_result_files"] += 1
                continue
            
            # 使用文件类型检测替代扩展名检查
            file_type = detect_file_type(full_file_path)
            ext = get_file_extension_from_mime(file_type)
            
            # 检查是否为支持的文件格式
            if ext in config["supported_formats"] or file_type in config["supported_formats"]:
                supported_files.append(full_file_path)
            else:
                filtered_files["unsupported_formats"] += 1
    
    # 记录详细过滤信息
    print(f"\n【文件收集统计】")
    print(f"总文件数: {total_files}")
    print(f"跳过文件夹中的文件数: {filtered_files['skipped_folders']}")
    print(f"排除文件数: {filtered_files['excluded_files']}")
    print(f"临时/无效文件数: {filtered_files['temp_invalid_files']}")
    print(f"汇总结果文件数: {filtered_files['summary_result_files']}")
    print(f"不支持的文件格式数: {filtered_files['unsupported_formats']}")
    print(f"最终收集文件数: {len(supported_files)}")
    
    # 对文件列表进行排序，确保处理顺序一致
    supported_files.sort()
    return supported_files


def detect_file_type(file_path: str) -> str:
    """
    使用python-magic检测文件类型，返回文件的MIME类型或扩展名
    
    Args:
        file_path: 文件路径
    
    Returns:
        文件类型，如'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'或'csv'
    """
    if magic_available:
        try:
            mime = magic.Magic(mime=True)
            file_mime = mime.from_file(file_path)
            return file_mime
        except Exception:
            pass
    
    # 降级方案：使用文件扩展名
    return os.path.splitext(file_path)[1].lower()


def get_file_extension_from_mime(mime_type: str) -> str:
    """
    根据MIME类型获取对应的文件扩展名
    
    Args:
        mime_type: MIME类型
    
    Returns:
        文件扩展名，如'.xlsx', '.csv'等
    """
    mime_to_ext = {
        # Excel文件
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx',
        'application/vnd.ms-excel': '.xls',
        'application/vnd.ms-excel.sheet.macroEnabled.12': '.xlsm',
        # CSV文件
        'text/csv': '.csv',
        'application/csv': '.csv',
        # TXT文件
        'text/plain': '.txt',
        # HTML文件
        'text/html': '.html',
        # PDF文件
        'application/pdf': '.pdf'
    }
    return mime_to_ext.get(mime_type, os.path.splitext(mime_type)[1].lower())


def is_excel_file(file_path: str) -> bool:
    """
    检测是否为Excel文件
    """
    file_type = detect_file_type(file_path)
    if magic_available:
        return file_type in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                           'application/vnd.ms-excel',
                           'application/vnd.ms-excel.sheet.macroEnabled.12']
    else:
        return file_path.lower().endswith(('.xlsx', '.xls', '.et'))


def is_text_file(file_path: str) -> bool:
    """
    检测是否为文本文件（CSV/TXT/HTML/HTM）
    """
    file_type = detect_file_type(file_path)
    if magic_available:
        return file_type in ['text/csv', 'application/csv', 'text/plain', 'text/html']
    else:
        return file_path.lower().endswith(('.csv', '.txt', '.html', '.htm'))


def is_pdf_file(file_path: str) -> bool:
    """
    检测是否为PDF文件
    """
    file_type = detect_file_type(file_path)
    if magic_available:
        return file_type == 'application/pdf'
    else:
        return file_path.lower().endswith('.pdf')


def is_excel_encrypted(file_path: str) -> bool:
    """
    检测Excel文件是否加密
    """
    if not is_excel_file(file_path):
        return False
    try:
        with open(file_path, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            return office_file.is_encrypted()
    except Exception:
        return False
