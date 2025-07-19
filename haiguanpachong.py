import pandas as pd

# 读取人工对应海关编码文件
china_df = pd.read_excel("体系-中国海关编码-第一队 带编号-处理后.xlsx", header=None)
china_df.columns = ["Product", "HS_Code_China1", "Desc1"]
# 合并所有海关编码列
china_hscodes = china_df.melt(id_vars=["Product"],
                            value_vars=["HS_Code_China1"],
                            value_name="HS_Code_China")[["Product", "HS_Code_China"]]

# 提取中国HS Code前6位
china_hscodes["HS6"] = china_hscodes["HS_Code_China"].astype(str).str[:6]

# 读取外国海关编码文件
us_df = pd.read_excel("vietnam.xlsx", header=None)
us_df.columns = ["Country", "ID", "Root", "Children", "HS_Code_US", "Desc"]
# 清洗外国编码（去除点号）
us_df["HS6"] = us_df["HS_Code_US"].str.replace(r"[^\d]", "", regex=True).str[:6]

# 合并映射关系
mapping_df = pd.merge(china_hscodes, us_df, on="HS6", how="inner")
mapping_df = mapping_df[["Product", "HS_Code_China", "HS6", "HS_Code_US", "Desc"]]
mapping_df.to_excel("中国越南海关编码映射表_第一组.xlsx", index=False)