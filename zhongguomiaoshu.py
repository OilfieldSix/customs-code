import pandas as pd

# 读取两个Excel文件
mapping_df = pd.read_excel('中美海关编码映射表_第三组_匹配后.xlsx')
china_code_df = pd.read_excel('中国海关商品编码 - 精选.xlsx')

# 将编号转换为字符串类型（防止前导零丢失）
china_code_df['编号'] = china_code_df['编号'].astype(str)
# 创建编码到名称的映射字典
code_to_name = dict(zip(china_code_df['编号'], china_code_df['名称']))

# 添加新列并进行编码匹配
mapping_df['HS_Code_China'] = mapping_df['HS_Code_China'].astype(str)
mapping_df['商品名称'] = mapping_df['HS_Code_China'].map(code_to_name)

# 保存结果到新文件
mapping_df.to_excel('中美海关编码映射表_第三组_匹配后_带商品名称.xlsx', index=False)