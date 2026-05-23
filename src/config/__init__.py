import yaml
from typing import Dict


def load_config(config_path: str) -> Dict:
    """
    从YAML文件加载配置
    """
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        raise Exception(f"加载配置文件失败：{str(e)}")


def init_bank_alias_mapping(config: Dict) -> None:
    """
    从standard_banks中自动提取别名，生成alias_to_standard映射
    无需手动维护，新增银行仅需补充standard_banks配置
    """
    bank_standard_config = config.get("bank_standardization_config", {})
    standard_banks = bank_standard_config.get("standard_banks", {})
    alias_to_standard = {}

    for standard_name, bank_info in standard_banks.items():
        aliases = bank_info.get("alias", [])
        # 1. 所有别名映射到标准名称
        for alias in aliases:
            if alias and alias not in alias_to_standard:
                alias_to_standard[alias] = standard_name
        # 2. 标准名称自身映射到自己（避免遗漏精准匹配）
        if standard_name not in alias_to_standard:
            alias_to_standard[standard_name] = standard_name

    # 更新到配置中，供后续使用
    bank_standard_config["alias_to_standard"] = alias_to_standard
