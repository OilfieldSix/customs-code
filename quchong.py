import pandas as pd

# 读取Excel文件
file_path = "中国新加坡海关编码映射表_第一组.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# 删除完全重复的行（基于所有列）
df_cleaned = df.drop_duplicates()

# 检查是否有重复行被删除
print(f"原始行数: {len(df)}")
print(f"去重后行数: {len(df_cleaned)}")

# 保存到新文件（或覆盖原文件）
output_path = "中国新加坡海关编码映射表_第一组_去重后.xlsx"
df_cleaned.to_excel(output_path, index=False, sheet_name="Sheet1")

print(f"去重完成，结果已保存到: {output_path}")