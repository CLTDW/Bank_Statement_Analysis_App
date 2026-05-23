import yaml
from pathlib import Path
from typing import Dict, Any

class ConfigManager:
    """
    配置管理类，负责加载和管理配置文件
    """
    def __init__(self, config_file: str = "config.yaml"):
        """
        初始化配置管理器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = Path(config_file)
        self.config: Dict[str, Any] = {}
        self.load_config()
    
    def load_config(self) -> None:
        """
        加载配置文件
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                self.config = yaml.safe_load(f)
        except FileNotFoundError:
            raise FileNotFoundError(f"配置文件 {self.config_file} 不存在")
        except yaml.YAMLError as e:
            raise ValueError(f"配置文件格式错误：{str(e)}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        获取配置值
        
        Args:
            key: 配置键
            default: 默认值
            
        Returns:
            配置值
        """
        return self.config.get(key, default)
    
    def update(self, key: str, value: Any) -> None:
        """
        更新配置值
        
        Args:
            key: 配置键
            value: 配置值
        """
        self.config[key] = value
    
    def save(self) -> None:
        """
        保存配置到文件
        """
        with open(self.config_file, 'w', encoding='utf-8') as f:
            yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
    
    def get_base_dir(self) -> Path:
        """
        获取基础目录
        
        Returns:
            基础目录路径
        """
        return Path(self.config.get('base_dir', '.'))
    
    def get_template_file(self) -> Path:
        """
        获取模板文件路径
        
        Returns:
            模板文件路径
        """
        import os
        # 模板文件存放在程序目录下
        program_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        template_file = self.config.get('template_file', '流水汇总模板.xlsx')
        return Path(program_dir) / template_file
    
    def get_output_file(self) -> Path:
        """
        获取输出文件路径
        
        Returns:
            输出文件路径
        """
        return self.get_base_dir() / self.config.get('output_file', '银行流水汇总结果.xlsx')
    
    def get_log_file(self) -> Path:
        """
        获取日志文件路径
        
        Returns:
            日志文件路径
        """
        return self.get_base_dir() / self.config.get('log_file', '银行流水汇总日志.log')
    
    def get_report_file(self) -> Path:
        """
        获取报告文件路径
        
        Returns:
            报告文件路径
        """
        return self.get_base_dir() / self.config.get('report_file', '银行流水汇总情况报告.txt')
    
    def get_log_level(self) -> str:
        """
        获取日志级别
        
        Returns:
            日志级别
        """
        return self.config.get('log_level', 'INFO')
    
    def get_supported_formats(self) -> list:
        """
        获取支持的文件格式
        
        Returns:
            支持的文件格式列表
        """
        return self.config.get('supported_formats', ['.xlsx', '.xls', '.csv', '.txt', '.pdf', '.html', '.htm', '.et'])
    
    def get_excluded_files(self) -> list:
        """
        获取排除的文件列表
        
        Returns:
            排除的文件列表
        """
        return self.config.get('excluded_files', [])
    
    def get_excluded_patterns(self) -> list:
        """
        获取排除的文件模式列表
        
        Returns:
            排除的文件模式列表
        """
        return self.config.get('excluded_patterns', [])
