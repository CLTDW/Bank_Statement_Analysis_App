#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多账户银行流水处理模块
用于处理同一Sheet中存在多个账户流水分段展示的情况
"""
from typing import List, Dict, Tuple, Optional
import pandas as pd
import logging
import time

from src.utils import standardize_header


class MultiAccountHandler:
    """
    多账户流水处理器
    """
    
    def __init__(self, header_mapping: Dict[str, str], keyword_mapping: Dict[str, List[str]], object_names: List[str] = None, surnames: Dict[str, List[str]] = None):
        """
        初始化多账户处理器
        
        Args:
            header_mapping: 表头映射字典
            keyword_mapping: 关键词映射字典
            object_names: 用户输入的对象名称列表
            surnames: 姓氏库字典，包含单姓和复姓
        """
        self.header_mapping = header_mapping
        self.keyword_mapping = keyword_mapping
        self.object_names = object_names or []
        self.surnames = surnames or {'single': [], 'compound': []}
        self.logger = logging.getLogger('bank_statement_aggregator')
        
        # 定义账户信息相关的统一字段名
        self.account_info_fields = ['户主名称', '本方账号', '本方卡号', '本方账户开户行']
        
        # 定义交易数据相关的统一字段名
        self.transaction_fields = ['交易金额', '交易时间', '交易日期', '借方金额', '贷方金额', '账户余额']
        
        # 缓存机制
        self._separator_score_cache = {}  # 缓存分隔行得分
        self._account_info_cache = {}     # 缓存账户信息提取结果
        self._header_score_cache = {}      # 缓存表头得分
        
        # 缓存配置
        self.base_cache_size_limit = 1000  # 基础缓存大小限制
        self.cache_size_limit = self.base_cache_size_limit
        
        # 缓存优先级（数字越小优先级越高）
        self.cache_priorities = {
            'header_score': 1,      # 表头得分缓存优先级最高
            'separator_score': 2,   # 分隔行得分缓存优先级次之
            'account_info': 3       # 账户信息缓存优先级最低
        }
        
        # 缓存访问时间，用于LRU缓存清理
        self.cache_access_times = {
            'header_score': {},
            'separator_score': {},
            'account_info': {}
        }
        
        # 缓存过期时间（秒）
        self.cache_expiration_times = {
            'header_score': 3600,    # 1小时
            'separator_score': 1800,  # 30分钟
            'account_info': 900       # 15分钟
        }
        
        # 缓存创建时间，用于过期清理
        self.cache_creation_times = {
            'header_score': {},
            'separator_score': {},
            'account_info': {}
        }
        
        # 内存使用监控
        self.memory_threshold = 80  # 内存使用阈值（百分比）
        self.last_cache_cleanup = time.time()
        self.cache_cleanup_interval = 5  # 缓存清理间隔（秒）
        self.last_memory_check = time.time()
        self.memory_check_interval = 10  # 内存检查间隔（秒）
    
    def _get_available_memory(self) -> float:
        """
        获取当前可用内存百分比
        
        Returns:
            可用内存百分比
        """
        try:
            import psutil
            memory = psutil.virtual_memory()
            return 100 - memory.percent
        except:
            return 50  # 默认返回50%
    
    def _check_memory_usage(self):
        """
        检查内存使用情况，当超过阈值时进行内存清理
        """
        try:
            import psutil
            memory = psutil.virtual_memory()
            memory_usage = memory.percent
            

            
            if memory_usage > self.memory_threshold:
                self.logger.warning(f"【内存监控】内存使用超过阈值，进行紧急内存清理")
                
                # 紧急清理所有缓存
                for cache_type in self.cache_priorities.keys():
                    self._cleanup_cache(cache_type)
                
                # 强制垃圾回收
                import gc
                gc.collect()
                
                # 再次检查内存使用情况
                memory = psutil.virtual_memory()
                new_memory_usage = memory.percent
                self.logger.info(f"【内存监控】紧急清理后内存使用：{new_memory_usage:.2f}%")
        except:
            pass  # 如果无法获取内存信息，忽略
    
    def _adjust_cache_size(self):
        """
        根据系统内存情况动态调整缓存大小
        """
        available_memory = self._get_available_memory()
        
        # 根据可用内存调整缓存大小
        if available_memory < 20:
            # 内存紧张，减少缓存大小
            self.cache_size_limit = max(100, int(self.base_cache_size_limit * 0.3))
            self.logger.warning(f"【缓存管理】内存紧张，调整缓存大小为：{self.cache_size_limit}")
        elif available_memory < 50:
            # 内存适中，保持中等缓存大小
            self.cache_size_limit = int(self.base_cache_size_limit * 0.7)
        else:
            # 内存充足，使用基础缓存大小
            self.cache_size_limit = self.base_cache_size_limit
    
    def _cleanup_cache(self, cache_type: str):
        """
        清理指定类型的缓存
        
        Args:
            cache_type: 缓存类型
        """
        cache_mapping = {
            'header_score': self._header_score_cache,
            'separator_score': self._separator_score_cache,
            'account_info': self._account_info_cache
        }
        
        cache = cache_mapping.get(cache_type)
        if not cache:
            return
        
        current_time = time.time()
        expiration_time = self.cache_expiration_times.get(cache_type, 3600)
        
        # 清理过期缓存
        valid_items = []
        for key, value in cache.items():
            # 检查缓存是否过期
            creation_time = self.cache_creation_times.get(cache_type, {}).get(str(key), 0)
            if current_time - creation_time < expiration_time:
                valid_items.append((key, value))
        
        # 按访问时间排序，保留最近使用的缓存
        access_times = self.cache_access_times.get(cache_type, {})
        sorted_items = sorted(valid_items, 
                            key=lambda x: access_times.get(str(x[0]), 0), 
                            reverse=True)
        
        # 控制缓存大小
        keep_count = min(int(self.cache_size_limit * 0.8), len(sorted_items))
        new_cache = dict(sorted_items[:keep_count])
        
        # 更新缓存
        if cache_type == 'header_score':
            self._header_score_cache = new_cache
        elif cache_type == 'separator_score':
            self._separator_score_cache = new_cache
        elif cache_type == 'account_info':
            self._account_info_cache = new_cache
        
        # 清理访问时间和创建时间记录
        valid_keys = [str(key) for key in new_cache.keys()]
        self.cache_access_times[cache_type] = {k: v for k, v in access_times.items() if k in valid_keys}
        creation_times = self.cache_creation_times.get(cache_type, {})
        self.cache_creation_times[cache_type] = {k: v for k, v in creation_times.items() if k in valid_keys}
        
        cleaned_count = len(cache) - len(new_cache)
        if cleaned_count > 0:
            self.logger.info(f"【缓存清理】{cache_type}缓存清理完成，清理了{cleaned_count}项，当前大小：{len(new_cache)}")
    
    def _update_cache_access(self, cache_type: str, key: str):
        """
        更新缓存访问时间
        
        Args:
            cache_type: 缓存类型
            key: 缓存键
        """
        current_time = time.time()
        key_str = str(key)
        
        # 更新访问时间
        if cache_type not in self.cache_access_times:
            self.cache_access_times[cache_type] = {}
        self.cache_access_times[cache_type][key_str] = current_time
        
        # 记录缓存创建时间（如果是新缓存）
        if cache_type not in self.cache_creation_times:
            self.cache_creation_times[cache_type] = {}
        if key_str not in self.cache_creation_times[cache_type]:
            self.cache_creation_times[cache_type][key_str] = current_time
        
        # 定期清理缓存
        if current_time - self.last_cache_cleanup > self.cache_cleanup_interval:
            self._adjust_cache_size()
            # 按优先级清理缓存
            for cache_type in sorted(self.cache_priorities.keys(), key=lambda x: self.cache_priorities[x], reverse=True):
                self._cleanup_cache(cache_type)
            self.last_cache_cleanup = current_time
        
        # 定期检查内存使用情况
        if current_time - self.last_memory_check > self.memory_check_interval:
            self._check_memory_usage()
            self.last_memory_check = current_time
    
    def _is_account_info_row(self, row_data: pd.Series) -> bool:
        """
        判断一行是否为账户信息行（分隔行）
        
        Args:
            row_data: 行数据
            
        Returns:
            是否为账户信息行
        """
        # 使用得分计算方法判断
        score = self._calculate_separator_score(row_data)
        # 得分大于0的行为账户信息行
        return score > 0
    
    def _extract_account_info(self, row_data: pd.Series, df_raw: pd.DataFrame = None, row_index: int = None, existing_account_info: Dict[str, str] = None) -> Dict[str, str]:
        """
        从账户信息行中提取账户信息
        
        Args:
            row_data: 账户信息行数据
            df_raw: 原始DataFrame（可选，用于从下侧单元格查找值）
            row_index: 当前行索引（可选，用于从下侧单元格查找值）
            existing_account_info: 已存在的账户信息，用于保留原有数据
            
        Returns:
            提取到的账户信息字典
        """
        # 生成缓存键
        row_text = ' '.join([str(cell).strip() for cell in row_data.values if pd.notna(cell)])
        cache_key = row_text
        
        # 检查缓存
        if cache_key in self._account_info_cache:
            # 更新缓存访问时间
            self._update_cache_access('account_info', cache_key)
            cached_info = self._account_info_cache[cache_key].copy()
            # 保留原有数据
            if existing_account_info:
                for key in ['本方账号', '本方卡号']:
                    if key in existing_account_info:
                        cached_info[key] = existing_account_info[key]
            return cached_info
        
        account_info = {}
        # 保留原有数据
        if existing_account_info:
            for key in ['本方账号', '本方卡号']:
                if key in existing_account_info:
                    account_info[key] = existing_account_info[key]
        
        row_cells = [str(cell).strip() for cell in row_data.values if pd.notna(cell)]
        
        # 优先提取用户输入的对象名称作为户主名称
        if self.object_names:
            matched_names = []
            for name in self.object_names:
                if name in row_text:
                    matched_names.append(name)
            
            # 如果只匹配到一个名称，使用它作为户主名称
            if len(matched_names) == 1:
                account_info['户主名称'] = matched_names[0]
            elif len(matched_names) > 1:
                # 如果匹配到多个名称，跳过户主名称提取
                pass
        
        # 尝试从单元格中提取键值对
        for i, cell in enumerate(row_cells):
            cell = cell.strip()
            if not cell:
                continue
            
            # 检查是否包含中文逗号分隔的多个信息
            if '，' in cell:
                # 按中文逗号分割
                info_parts = cell.split('，')
                for i, part in enumerate(info_parts):
                    part = part.strip()
                    if not part:
                        continue
                    
                    # 检查是否包含分隔符
                    has_separator = False
                    for separator in [':', '：', '=', '→', '→', ' ', '\t', '、', '-', '—', '–']:
                        if separator in part:
                            has_separator = True
                            parts = part.split(separator, 1)
                            if len(parts) == 2:
                                key_part = parts[0].strip()
                                value_part = parts[1].strip()
                                
                                # 匹配键到统一字段
                                matched_field = self._match_key_to_field(key_part)
                                if matched_field and value_part:
                                    # 如果是户主名称且已经从用户输入中提取，则跳过
                                    if matched_field == '户主名称' and '户主名称' in account_info:
                                        continue
                                    # 只添加未存在的字段，保留原有数据
                                    if matched_field not in account_info:
                                        account_info[matched_field] = value_part
                                    break
                    
                    # 如果是第一个部分且没有分隔符，尝试作为户主名称
                    if i == 0 and not has_separator and '户主名称' not in account_info:
                        # 验证名称，优先匹配人名，其次匹配公司名
                        validated_name = self._validate_name(part)
                        if validated_name:
                            account_info['户主名称'] = validated_name
            else:
                # 尝试多种分隔符
                for separator in [':', '：', '=', '→', '→', ' ', '\t', '、', '-', '—', '–']:
                    if separator in cell:
                        parts = cell.split(separator, 1)
                        if len(parts) == 2:
                            key_part = parts[0].strip()
                            value_part = parts[1].strip()
                            
                            # 匹配键到统一字段
                            matched_field = self._match_key_to_field(key_part)
                            if matched_field and value_part:
                                # 如果是户主名称且已经从用户输入中提取，则跳过
                                if matched_field == '户主名称' and '户主名称' in account_info:
                                    continue
                                # 验证户主名称，优先匹配人名，其次匹配公司名
                                if matched_field == '户主名称':
                                    validated_name = self._validate_name(value_part)
                                    if validated_name and '户主名称' not in account_info:
                                        account_info[matched_field] = validated_name
                                elif matched_field not in account_info:
                                    # 只添加未存在的字段，保留原有数据
                                    account_info[matched_field] = value_part
                                break
            
            # 如果键值对没有匹配成功，且该字段还未被添加，尝试在右侧单元格查找值
            if i + 1 < len(row_cells):
                matched_field = self._match_key_to_field(cell)
                if matched_field and matched_field not in account_info:
                    next_cell = row_cells[i + 1].strip()
                    if next_cell and not self._match_key_to_field(next_cell):
                        # 验证户主名称，优先匹配人名，其次匹配公司名
                        if matched_field == '户主名称':
                            validated_name = self._validate_name(next_cell)
                            if validated_name:
                                account_info[matched_field] = validated_name
                        else:
                            account_info[matched_field] = next_cell
        
        # 如果键值对没有匹配成功，且该字段还未被添加，尝试在下侧单元格查找值
        if df_raw is not None and row_index is not None and row_index + 1 < len(df_raw):
            # 遍历当前行的每个单元格
            for i, (col, cell) in enumerate(row_data.items()):
                cell = str(cell).strip()
                if not cell:
                    continue
                
                matched_field = self._match_key_to_field(cell)
                if matched_field and matched_field not in account_info:
                    # 获取下一行对应列的值
                    next_row_data = df_raw.iloc[row_index + 1]
                    next_cell = str(next_row_data.iloc[i]).strip() if i < len(next_row_data) else ''
                    if next_cell and not self._match_key_to_field(next_cell):
                        # 验证户主名称，优先匹配人名，其次匹配公司名
                        if matched_field == '户主名称':
                            validated_name = self._validate_name(next_cell)
                            if validated_name:
                                account_info[matched_field] = validated_name
                        else:
                            account_info[matched_field] = next_cell
        
        # 如果没有提取到户主名称，在账户分隔行上下三行内搜索
        if '户主名称' not in account_info and df_raw is not None and row_index is not None:
            # 搜索上下三行
            start_search = max(0, row_index - 3)
            end_search = min(len(df_raw), row_index + 4)  # +4 because end is exclusive
            
            for i in range(start_search, end_search):
                if i == row_index:
                    continue
                
                search_row = df_raw.iloc[i]
                search_score = self._calculate_separator_score(search_row)
                if search_score > 0:
                    search_info = self._extract_account_info(search_row, df_raw, i)
                    if '户主名称' in search_info:
                        account_info['户主名称'] = search_info['户主名称']
                        break
        
        # 如果仍然没有提取到户主名称，搜索表格页前50行
        if '户主名称' not in account_info and df_raw is not None:
            search_range = min(50, len(df_raw))
            for i in range(search_range):
                if i == row_index:
                    continue
                
                search_row = df_raw.iloc[i]
                search_score = self._calculate_separator_score(search_row)
                if search_score > 0:
                    search_info = self._extract_account_info(search_row, df_raw, i)
                    if '户主名称' in search_info:
                        account_info['户主名称'] = search_info['户主名称']
                        break
        
        # 如果没有提取到足够信息，尝试整行文本提取
        if len(account_info) == 0:
            account_info = self._extract_from_full_text(row_text)
        
        # 缓存结果（缓存原始提取结果，不包含原有数据）
        self._account_info_cache[cache_key] = account_info.copy()
        # 更新缓存访问时间
        self._update_cache_access('account_info', cache_key)
        
        return account_info
    
    def _match_key_to_field(self, key: str) -> Optional[str]:
        """
        将键匹配到统一字段名
        
        Args:
            key: 键文本
            
        Returns:
            匹配到的统一字段名，或None
        """
        key = key.strip()
        
        # 检查精确映射
        if key in self.header_mapping:
            unified_field = self.header_mapping[key]
            if unified_field in self.account_info_fields:
                return unified_field
        
        # 检查关键词映射
        for field in self.account_info_fields:
            if field in self.keyword_mapping:
                for keyword in self.keyword_mapping[field]:
                    if keyword in key:
                        return field
        
        return None
    
    def _has_valid_surname(self, name: str) -> bool:
        """
        检查名称是否包含有效的姓氏
        
        Args:
            name: 要检查的名称
            
        Returns:
            是否包含有效的姓氏
        """
        if not name:
            return False
        
        # 检查复姓
        for surname in self.surnames.get('compound', []):
            if name.startswith(surname):
                return True
        
        # 检查单姓（前1-2个字符）
        for i in range(1, min(3, len(name) + 1)):
            surname_candidate = name[:i]
            if surname_candidate in self.surnames.get('single', []):
                return True
        
        return False
    
    def _is_company_name(self, name: str) -> bool:
        """
        检查名称是否为公司名
        
        Args:
            name: 要检查的名称
            
        Returns:
            是否为公司名
        """
        if not name:
            return False
        
        # 公司名关键词
        company_keywords = [
            '公司', '有限责任公司', '有限公司', '股份有限公司', '集团', '集团公司',
            '企业', '企业集团', '合作社', '协会', '商会', '联合会',
            '事务所', '工作室', '中心', '研究院', '研究所', '学院',
            '大学', '学校', '医院', '医疗机构', '银行', '金融机构',
            '保险公司', '证券公司', '基金公司', '科技公司', '网络公司',
            '软件公司', '贸易公司', '建筑公司', '房地产公司', '制药公司',
            '食品公司', '服装公司', '物流公司', '运输公司', '快递公司',
            '餐饮公司', '酒店', '旅行社', '传媒公司', '广告公司',
            '咨询公司', '管理公司', '投资公司', '资产管理公司', '财务公司'
        ]
        
        # 检查是否包含公司名关键词
        for keyword in company_keywords:
            if keyword in name:
                return True
        
        return False
    
    def _validate_name(self, name: str) -> str:
        """
        验证名称，优先返回人名，其次返回公司名
        
        Args:
            name: 要验证的名称
            
        Returns:
            验证后的名称，如果都不满足返回空字符串
        """
        if not name:
            return ""
        
        # 优先检查是否为人名
        if self._has_valid_surname(name):
            return name
        
        # 其次检查是否为公司名
        if self._is_company_name(name):
            return name
        
        # 都不满足返回空字符串
        return ""
    
    def _extract_from_full_text(self, text: str) -> Dict[str, str]:
        """
        从完整文本中尝试提取账户信息
        
        Args:
            text: 完整文本
            
        Returns:
            提取到的账户信息
        """
        account_info = {}
        
        # 尝试匹配账号/卡号（纯数字或数字+分隔符）
        import re
        account_patterns = [
            r'(\d{4,30})',  # 4-30位数字
            r'(\d[\d\s\-]{4,30}\d)',  # 带空格或连字符的账号
        ]
        
        for pattern in account_patterns:
            matches = re.findall(pattern, text)
            for match in matches:
                cleaned = re.sub(r'[\s\-]', '', match)
                if len(cleaned) >= 4:
                    # 优先匹配到本方账号
                    if '本方账号' not in account_info:
                        account_info['本方账号'] = cleaned
                    elif '本方卡号' not in account_info:
                        account_info['本方卡号'] = cleaned
                    break
        
        return account_info
    
    def _calculate_separator_score(self, row_data: pd.Series) -> float:
        """
        计算分隔行的匹配度得分
        
        Args:
            row_data: 行数据
            
        Returns:
            分隔行得分，越高越好
        """
        # 生成缓存键
        row_text = ' '.join([str(cell).strip() for cell in row_data.values if pd.notna(cell)])
        cache_key = row_text
        
        # 检查缓存
        if cache_key in self._separator_score_cache:
            # 更新缓存访问时间
            self._update_cache_access('separator_score', cache_key)
            return self._separator_score_cache[cache_key]
        
        score = 0
        
        # 关键账户信息字段，权重更高
        key_account_fields = ['户主名称', '本方账号', '本方卡号']
        
        # 计算关键账户信息字段的得分
        for field in key_account_fields:
            # 检查精确映射
            for raw_header, unified_header in self.header_mapping.items():
                if unified_header == field and raw_header in row_text:
                    score += 3  # 关键字段权重更高
                    break
            
            # 检查关键词映射
            if field in self.keyword_mapping:
                for keyword in self.keyword_mapping[field]:
                    if keyword in row_text:
                        score += 2  # 关键词匹配也有较高权重
                        break
        
        # 确保不包含交易数据关键词
        for field in self.transaction_fields:
            # 检查精确映射
            for raw_header, unified_header in self.header_mapping.items():
                if unified_header == field and raw_header in row_text:
                    score -= 5  # 包含交易数据关键词会大幅降低得分
                    break
            
            # 检查关键词映射
            if field in self.keyword_mapping:
                for keyword in self.keyword_mapping[field]:
                    if keyword in row_text:
                        score -= 3  # 包含交易数据关键词会降低得分
                        break
        
        # 确保得分不为负
        score = max(0, score)
        
        # 缓存结果
        self._separator_score_cache[cache_key] = score
        # 更新缓存访问时间
        self._update_cache_access('separator_score', cache_key)
        
        return score
    
    def _calculate_header_score(self, standardized_headers: List[str]) -> float:
        """
        计算表头得分（与原有逻辑一致）
        
        Args:
            standardized_headers: 标准化后的表头列表
            
        Returns:
            表头得分
        """
        # 生成缓存键
        cache_key = tuple(standardized_headers)
        
        # 检查缓存
        if cache_key in self._header_score_cache:
            # 更新缓存访问时间
            self._update_cache_access('header_score', cache_key)
            return self._header_score_cache[cache_key]
        
        exact_match_score = sum(1 for h in standardized_headers if h in self.header_mapping)
        keyword_match_score = 0
        for h in standardized_headers:
            for unified_col, keywords in self.keyword_mapping.items():
                if any(keyword in h for keyword in keywords):
                    keyword_match_score += 1
                    break
        
        score = exact_match_score * 2 + keyword_match_score * 1
        
        # 缓存结果
        self._header_score_cache[cache_key] = score
        # 更新缓存访问时间（同时会记录创建时间）
        self._update_cache_access('header_score', cache_key)
        
        return score
    
    def _find_best_header_row(self, df_segment: pd.DataFrame) -> Tuple[int, float]:
        """
        在数据块中寻找最佳表头行（与原有逻辑一致）
        
        Args:
            df_segment: 数据块
            
        Returns:
            (最佳表头行索引, 表头得分)
        """
        best_row = 0
        best_score = 0
        
        # 最多检查前50行
        max_check_rows = min(50, len(df_segment))
        
        for i in range(max_check_rows):
            # 尝试将第i行作为表头
            header_candidate = df_segment.iloc[i]
            
            # 标准化表头候选
            standardized_headers = [standardize_header(str(h)) for h in header_candidate]
            
            # 计算表头得分（与原有逻辑一致）
            total_score = self._calculate_header_score(standardized_headers)
            
            # 更新最佳表头：只在得分更高时更新，得分相同时保留更前面的行
            if total_score > best_score:
                best_score = total_score
                best_row = i
        
        return best_row, best_score
    
    def identify_separator_rows(self, df_raw: pd.DataFrame) -> List[int]:
        """
        识别账户分隔行（第一阶段：快速扫描）
        
        Args:
            df_raw: 原始Sheet数据
            
        Returns:
            有效的账户分隔行行号列表
        """
        # 识别所有账户分隔行并计算得分
        separator_candidates = []
        for i in range(len(df_raw)):
            score = self._calculate_separator_score(df_raw.iloc[i])
            if score > 0:
                separator_candidates.append((i, score))
        
        # 如果没有识别到账户分隔行，返回空列表
        if len(separator_candidates) == 0:
            return []
        
        # 选择最佳分隔行（选择得分最高的行）
        max_score = max(score for _, score in separator_candidates)
        
        # 选择所有得分等于最高分的行
        best_separators = [row for row, score in separator_candidates if score == max_score]
        
        # 按行号排序
        best_separators.sort()
        
        return best_separators
    
    def process_multi_account_sheet(self, df_raw: pd.DataFrame, file_identifier: str, best_header_score: float = 0) -> Optional[pd.DataFrame]:
        """
        处理多账户共存的Sheet数据
        
        Args:
            df_raw: 原始Sheet数据
            file_identifier: 文件标识符
            best_header_score: 最优表头行评分
            
        Returns:
            处理后的合并数据，或None
        """
        self.logger.info(f"【多账户处理】开始检测和处理多账户数据：{file_identifier}")
        self.logger.info(f"【多账户处理】最优表头行评分：{best_header_score}")
        
        # 检查内存使用情况
        self._check_memory_usage()
        
        # 第一阶段：识别账户分隔行
        best_separators = self.identify_separator_rows(df_raw)
        
        # 检查每个分隔行上下5行内是否存在评分大于最优表头行的表头行
        valid_separators = []
        for separator_row in best_separators:
            # 定义检查范围：上下5行
            start_check = max(0, separator_row - 5)
            end_check = min(len(df_raw), separator_row + 6)  # +6 because end is exclusive
            
            # 检查该范围内是否存在评分大于等于最优表头行的表头行
            has_better_header = False
            for i in range(start_check, end_check):
                # 跳过分隔行本身
                if i == separator_row:
                    continue
                
                # 检查当前行是否为评分大于等于最优表头行的表头行
                header_candidate = df_raw.iloc[i]
                standardized_headers = [standardize_header(str(h)) for h in header_candidate]
                header_score = self._calculate_header_score(standardized_headers)
                
                if header_score >= best_header_score:
                    has_better_header = True
                    break
            
            if has_better_header:
                valid_separators.append(separator_row)
        
        account_separator_rows = valid_separators
        
        self.logger.info(f"【多账户处理】识别到 {len(account_separator_rows)} 个最佳账户分隔行")
        
        # 如果没有有效的账户分隔行，返回None
        if len(account_separator_rows) == 0:
            self.logger.warning(f"【多账户处理】没有有效的账户分隔行，使用原有处理逻辑：{file_identifier}")
            return None
        
        # 第二阶段：分块处理数据
        all_processed_blocks = []
        
        # 处理第一个分隔行之前的数据（如果有的话）
        if account_separator_rows[0] > 0:
            # 检查是否有表头
            header_row, header_score = self._find_best_header_row(df_raw.iloc[:account_separator_rows[0]])
            if header_score > 0:
                # 处理第一个数据块
                block_df = df_raw.iloc[0:account_separator_rows[0]].copy()
                processed_block = self._process_data_block(block_df, {})
                if processed_block is not None:
                    all_processed_blocks.append(processed_block)
                # 及时释放内存
                del block_df
                import gc
                gc.collect()
        
        # 处理各个分隔行之间的数据块
        for i in range(len(account_separator_rows)):
            # 检查内存使用情况
            self._check_memory_usage()
            
            separator_row = account_separator_rows[i]
            
            if i + 1 < len(account_separator_rows):
                end_row = account_separator_rows[i + 1]
            else:
                end_row = len(df_raw)
            
            # 提取账户信息
            account_info = self._extract_account_info(df_raw.iloc[separator_row], df_raw, separator_row)
            self.logger.info(f"【多账户处理】第{i+1}个账户提取到信息：{account_info}")
            
            # 处理数据块
            block_df = df_raw.iloc[separator_row:end_row].copy()
            processed_block = self._process_data_block(block_df, account_info)
            if processed_block is not None:
                all_processed_blocks.append(processed_block)
            # 及时释放内存
            del block_df
            import gc
            gc.collect()
        
        # 第四步：合并所有账户数据
        if all_processed_blocks:
            # 分批次合并数据，避免一次性合并过多数据
            combined_df = None
            batch_size = 5  # 每批合并5个数据块
            
            for i in range(0, len(all_processed_blocks), batch_size):
                # 检查内存使用情况
                self._check_memory_usage()
                
                batch_data = all_processed_blocks[i:i+batch_size]
                if combined_df is None:
                    combined_df = pd.concat(batch_data, ignore_index=True)
                else:
                    combined_df = pd.concat([combined_df] + batch_data, ignore_index=True)
                # 及时释放内存
                del batch_data
                import gc
                gc.collect()
            
            self.logger.info(f"【多账户处理】多账户数据处理完成，共{len(combined_df)}行数据")
            # 最后检查内存使用情况
            self._check_memory_usage()
            return combined_df
        else:
            self.logger.warning(f"【多账户处理】没有成功处理任何账户数据块")
            return None
    
    def _process_data_block(self, block_df: pd.DataFrame, account_info: Dict[str, str]) -> Optional[pd.DataFrame]:
        """
        处理单个数据块
        
        Args:
            block_df: 数据块
            account_info: 账户信息
            
        Returns:
            处理后的数据块，或None
        """
        # 寻找最佳表头行
        header_row, header_score = self._find_best_header_row(block_df)
        
        if header_score == 0:
            return None
        
        # 设置表头
        headers = block_df.iloc[header_row].apply(lambda x: str(x).strip() if pd.notna(x) else '')
        block_df = block_df.iloc[header_row + 1:].reset_index(drop=True)
        block_df.columns = headers
        
        # 移除全NA列
        block_df = block_df.dropna(axis=1, how='all')
        
        # 检查内存使用情况
        self._check_memory_usage()
        
        # 标准化列名，用于匹配
        standardized_columns = {col: standardize_header(col) for col in block_df.columns}
        
        # 检查是否存在本方账号、本方卡号的匹配列
        existing_account_info = {}
        for col, standardized_col in standardized_columns.items():
            if standardized_col in self.header_mapping:
                unified_col = self.header_mapping[standardized_col]
                if unified_col in ['本方账号', '本方卡号']:
                    # 尝试获取该列的非空值作为原有数据
                    non_empty_values = block_df[col].dropna().astype(str).unique()
                    if len(non_empty_values) > 0:
                        existing_account_info[unified_col] = non_empty_values[0]
            
        # 处理账户信息，保留原有数据
        processed_account_info = {}
        if existing_account_info:
            # 保留原有数据
            processed_account_info.update(existing_account_info)
            # 添加其他账户信息
            for key, value in account_info.items():
                if key not in processed_account_info:
                    processed_account_info[key] = value
        else:
            processed_account_info = account_info
        
        # 添加账户信息列
        for key, value in processed_account_info.items():
            block_df[key] = value
        
        if not block_df.empty:
            # 确保列名唯一
            block_df.columns = [col.strip() for col in block_df.columns]
            # 检查并处理重复列名
            seen_cols = set()
            new_columns = []
            for col in block_df.columns:
                if col in seen_cols:
                    # 为重复列名添加后缀
                    suffix = 1
                    new_col = f"{col}_{suffix}"
                    while new_col in seen_cols:
                        suffix += 1
                        new_col = f"{col}_{suffix}"
                    new_columns.append(new_col)
                    seen_cols.add(new_col)
                else:
                    new_columns.append(col)
                    seen_cols.add(col)
            block_df.columns = new_columns
            
            self.logger.info(f"【多账户处理】数据块处理完成，共{len(block_df)}行")
            return block_df
        else:
            return None
    
    def _search_for_account_info(self, df_raw: pd.DataFrame, start_row: int = 0, end_row: int = 50) -> Optional[Dict[str, str]]:
        """
        在指定范围内搜索账户信息
        
        Args:
            df_raw: 原始数据
            start_row: 开始行
            end_row: 结束行
            
        Returns:
            提取到的账户信息，或None
        """
        search_range = min(end_row, len(df_raw))
        for i in range(start_row, search_range):
            row_data = df_raw.iloc[i]
            score = self._calculate_separator_score(row_data)
            if score > 0:
                account_info = self._extract_account_info(row_data, df_raw, i)
                if account_info:
                    return account_info
        return None
    
    def _process_single_account_no_name(self, df_raw: pd.DataFrame) -> Optional[pd.DataFrame]:
        """
        处理无户名单账户流水
        
        Args:
            df_raw: 原始数据
            
        Returns:
            处理后的数据，或None
        """
        # 寻找最佳表头
        best_header_row, best_header_score = self._find_best_header_row(df_raw)
        if best_header_score == 0:
            return None
        
        # 搜索前50行寻找账户信息
        account_info = self._search_for_account_info(df_raw, 0, 50)
        
        # 处理数据块
        block_df = df_raw.copy()
        processed_block = self._process_data_block(block_df, account_info or {})
        
        return processed_block
