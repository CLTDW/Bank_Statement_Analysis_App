import yaml
import os

# 读取 mappings.yaml 文件
mappings_path = os.path.join('src', 'config', 'mappings.yaml')

with open(mappings_path, 'r', encoding='utf-8') as f:
    mappings = yaml.safe_load(f)

# 获取 header_mapping 和 keyword_mapping
header_mapping = mappings.get('header_mapping', {})
keyword_mapping = mappings.get('keyword_mapping', {})

# 遍历 header_mapping，将原始表头添加到 keyword_mapping 中对应统一表头的关键词列表
for original_header, unified_header in header_mapping.items():
    # 确保 unified_header 在 keyword_mapping 中存在
    if unified_header not in keyword_mapping:
        keyword_mapping[unified_header] = []
    
    # 如果原始表头不在关键词列表中，添加它
    if original_header not in keyword_mapping[unified_header]:
        keyword_mapping[unified_header].append(original_header)

# 保存更新后的 mappings
with open(mappings_path, 'w', encoding='utf-8') as f:
    yaml.dump(mappings, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

print("Mappings updated successfully!")