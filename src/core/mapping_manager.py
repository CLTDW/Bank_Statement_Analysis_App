from typing import Dict, List, Optional
from pathlib import Path
import pandas as pd
import logging
import yaml

class MappingManager:
    """
    映射管理器类，负责处理表头映射和关键词映射
    """
    def __init__(self, config: Dict):
        """
        初始化映射管理器
        
        Args:
            config: 配置字典
        """
        self.logger = logging.getLogger(__name__)
        self.config = config
        
        # 加载映射配置文件
        self.mappings_config = self._load_mappings_config()
        
        # 从配置文件读取预期输出列
        self.expected_columns = self._load_expected_columns()
        
        # 初始化映射关系
        self.header_mapping: Dict[str, str] = self._load_header_mapping()
        self.keyword_mapping: Dict[str, List[str]] = self._load_keyword_mapping()
        
        # 加载排除规则
        self.exclusion_rules: Dict[str, List[str]] = self._load_exclusion_rules()
        
        # 加载姓氏库
        self.surnames: Dict[str, List[str]] = self._load_surnames()
        
        # 扩展关键词映射
        self._expand_keyword_mapping()
    
    def _load_mappings_config(self) -> Dict:
        """
        加载映射配置文件
        
        Returns:
            映射配置字典
        """
        try:
            # 配置文件路径
            mappings_file = Path(__file__).parent.parent / 'config' / 'mappings.yaml'
            with open(mappings_file, 'r', encoding='utf-8') as f:
                mappings_config = yaml.safe_load(f)
            self.logger.info(f"成功加载映射配置文件：{mappings_file}")
            return mappings_config
        except Exception as e:
            self.logger.error(f"加载映射配置文件失败：{str(e)}")
            # 返回空字典，后续方法会处理默认值
            return {}
    
    def _load_expected_columns(self) -> List[str]:
        """
        从模板文件加载预期输出列
        
        Returns:
            预期输出列列表
        """
        try:
            # 使用项目根目录作为模板文件的基础目录
            project_root = Path(__file__).parent.parent.parent
            template_file = project_root / self.config.get('template_file', '流水汇总模板.xlsx')
            df_template = pd.read_excel(template_file, header=0)
            expected_columns = list(df_template.columns)
            self.logger.info(f"从模板文件读取预期列：共{len(expected_columns)}列")
        except Exception as e:
            # 如果模板文件读取失败，从映射配置文件获取
            fallback_columns = self.mappings_config.get('fallback_expected_columns', [])
            if fallback_columns:
                expected_columns = fallback_columns
                self.logger.warning(f"从模板文件读取预期列失败，使用配置文件中的预期列：{str(e)}")
            else:
                # 配置文件也没有，则使用硬编码的预期列
                expected_columns = ['序号', '户主名称', '本方账户开户行', '银行卡类别', '本方账号', '本方卡号', '交易时间', '交易地点', 
                               '交易对方账号', '交易对方卡号', '交易对方名称', '交易对方开户行', '借贷标志', '币种', '交易金额', 
                               '账户余额', '交易摘要', '交易备注', '现金标识', '交易方式标识', '现金标志', '交易流水号', '柜员号', 
                               'IP地址', 'MAC地址', '一级流水分类', '二级流水分类', '三级流水分类']
                self.logger.warning(f"从模板文件和配置文件读取预期列均失败，使用硬编码预期列：{str(e)}")
        
        self.logger.info(f"预期输出列：{', '.join(expected_columns)}")
        return expected_columns
    
    def _load_header_mapping(self) -> Dict[str, str]:
        """
        加载表头映射关系，并排查重复映射
        
        Returns:
            表头映射字典
        """
        # 从配置文件加载表头映射
        header_mapping = self.mappings_config.get('header_mapping', {})
        if not header_mapping:
            self.logger.warning("配置文件中未找到表头映射，使用空映射")
            return header_mapping
        
        # 排查重复映射，保留所有唯一的原始表头到统一表头的映射
        # 注意：我们需要保留所有可能的原始表头映射，只要它们是唯一的
        unique_header_mapping = {}
        duplicate_count = 0
        
        for raw_header, unified_header in header_mapping.items():
            if raw_header not in unique_header_mapping:
                # 新的映射，直接添加
                unique_header_mapping[raw_header] = unified_header
            else:
                # 相同的原始表头映射到不同的统一列，保留第一个
                duplicate_count += 1
                self.logger.info(f"【表头映射去重】忽略重复原始表头: {raw_header} (已映射到 {unique_header_mapping[raw_header]})")
        
        if duplicate_count > 0:
            self.logger.info(f"【表头映射去重】完成，共处理 {duplicate_count} 个重复映射")
        
        return unique_header_mapping
    
    def _load_keyword_mapping(self) -> Dict[str, List[str]]:
        """
        加载关键词映射关系，并排查重复字段
        
        Returns:
            关键词映射字典
        """
        # 从配置文件加载关键词映射
        keyword_mapping = self.mappings_config.get('keyword_mapping', {})
        if not keyword_mapping:
            self.logger.warning("配置文件中未找到关键词映射，使用空映射")
            return keyword_mapping
        
        # 排查重复字段，保留合适唯一项
        unique_keyword_mapping = {}
        duplicate_count = 0
        
        for unified_col, keywords in keyword_mapping.items():
            if unified_col not in unique_keyword_mapping:
                # 去重关键词列表
                unique_keywords = []
                keyword_set = set()
                
                for keyword in keywords:
                    if keyword not in keyword_set:
                        keyword_set.add(keyword)
                        unique_keywords.append(keyword)
                    else:
                        duplicate_count += 1
                        self.logger.info(f"【关键词去重】忽略重复关键词: {keyword} (在 {unified_col} 中)")
                
                unique_keyword_mapping[unified_col] = unique_keywords
            else:
                # 存在重复的统一列，保留长度较长的关键词列表
                existing_keywords = unique_keyword_mapping[unified_col]
                if len(keywords) > len(existing_keywords):
                    # 去重新的关键词列表
                    unique_new_keywords = list(set(keywords))
                    unique_keyword_mapping[unified_col] = unique_new_keywords
                    duplicate_count += 1
                    self.logger.info(f"【关键词去重】替换重复统一列: {unified_col} (关键词数量: {len(existing_keywords)} → {len(unique_new_keywords)}")
                else:
                    duplicate_count += 1
                    self.logger.info(f"【关键词去重】忽略重复统一列: {unified_col} (已存在 {len(existing_keywords)} 个关键词)")
        
        if duplicate_count > 0:
            self.logger.info(f"【关键词去重】完成，共处理 {duplicate_count} 个重复项")
        
        return unique_keyword_mapping
    
    def _expand_keyword_mapping(self) -> None:
        """
        扩展关键词映射功能已取消
        现在只使用keyword_mapping中明确定义的关键词进行匹配
        """
        # 取消扩展关键词映射，不再将header_mapping中的键值对添加到keyword_mapping中
        # 这样可以避免将通用词（如"金额"）作为关键词匹配到不应匹配的字段
        pass
    
    def get_expected_columns(self) -> List[str]:
        """
        获取预期输出列
        
        Returns:
            预期输出列列表
        """
        return self.expected_columns
    
    def get_header_mapping(self) -> Dict[str, str]:
        """
        获取表头映射
        
        Returns:
            表头映射字典
        """
        return self.header_mapping
    
    def get_keyword_mapping(self) -> Dict[str, List[str]]:
        """
        获取关键词映射
        
        Returns:
            关键词映射字典
        """
        return self.keyword_mapping
    
    def get_exclusion_rules(self) -> Dict[str, List[str]]:
        """
        获取排除规则
        
        Returns:
            排除规则字典
        """
        return self.exclusion_rules
    
    def _load_exclusion_rules(self) -> Dict[str, List[str]]:
        """
        加载排除规则
        
        Returns:
            排除规则字典
        """
        try:
            # 从配置文件加载排除规则
            exclusion_rules = self.mappings_config.get('matching_rules', {}).get('exclusion_rules', {})
            if exclusion_rules:
                self.logger.info(f"成功加载排除规则：共{len(exclusion_rules)}个规则组")
            else:
                self.logger.warning("配置文件中未找到排除规则，使用空规则")
            return exclusion_rules
        except Exception as e:
            self.logger.error(f"加载排除规则失败：{str(e)}")
            return {}
    
    def _load_surnames(self) -> Dict[str, List[str]]:
        """
        加载姓氏库
        
        Returns:
            姓氏库字典，包含单姓和复姓
        """
        try:
            # 从配置文件加载姓氏库
            surnames = self.mappings_config.get('matching_rules', {}).get('surnames', {})
            if surnames:
                single_surnames = surnames.get('single', [])
                compound_surnames = surnames.get('compound', [])
                self.logger.info(f"成功加载姓氏库：单姓{len(single_surnames)}个，复姓{len(compound_surnames)}个")
            else:
                self.logger.warning("配置文件中未找到姓氏库，使用空库")
                surnames = {'single': [], 'compound': []}
            return surnames
        except Exception as e:
            self.logger.error(f"加载姓氏库失败：{str(e)}")
            return {'single': [], 'compound': []}
    
    def get_surnames(self) -> Dict[str, List[str]]:
        """
        获取姓氏库
        
        Returns:
            姓氏库字典，包含单姓和复姓
        """
        return self.surnames
