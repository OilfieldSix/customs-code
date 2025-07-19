import pandas as pd
import re
from openai import OpenAI
import time
import os

# 初始化OpenAI客户端（连接本地Ollama）
client = OpenAI(
    base_url="http://localhost:11434/v1",
    api_key="ollama"
)


def extract_last_number(text):
    """从文本中提取最后一个数字"""
    # 找到文本中所有连续数字序列
    all_numbers = re.findall(r'\d+', str(text))

    # 如果有数字存在，返回最后一个
    if all_numbers:
        return int(all_numbers[-1])

    return None  # 没有找到任何数字


def select_best_match(product_name, options):
    """
    使用大模型选择最匹配的选项
    返回选择的选项索引（从1开始）
    """
    # 构建选项字符串
    options_str = "\n".join([f"{idx}. [{code}] {desc}" for idx, (code, desc) in enumerate(options, start=1)])

    # 构建提示词
    prompt = f"""
        你是一位海关编码专家，需要为以下产品选择最合适的HS编码：

        产品名称: {product_name}

        可选编码:
        {options_str}

        请严格遵循以下规则：
        1. 只返回选择的选项数字（1, 2, 3...），不要任何解释
        2. 优先选择描述与产品名称最匹配的编码
        3. 当匹配度相近时，优先选择HS编码更具体的（位数更多的）
        4. 如果无法确定，选择第二个选项
        5. 尽快做出选择

        你的选择（只返回数字）:
        """

    # 调用模型
    try:
        response = client.chat.completions.create(
            model="deepseek-r1:1.5b",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,
            max_tokens=500
        )
        answer = response.choices[0].message.content.strip()
        print(f"产品: {product_name[:20]}... 模型返回: {answer}")
        return extract_last_number(answer)
    except Exception as e:
        print(f"处理产品 {product_name} 时出错：{str(e)}")
        return None


# 主处理函数
def process_hs_mapping(input_file, output_file):
    """
    处理海关编码映射表（简化兜底规则版）
    兜底规则仅使用HS编码位数判断
    """
    # 读取数据
    df = pd.read_excel(input_file)
    print(f"原始数据量: {len(df)}行")

    result_rows = []          # 存储结果
    fallback_products = []    # 记录触发兜底规则的产品

    # 按产品分组处理
    grouped = df.groupby('Product')
    for product, group in grouped:
        # 单选项直接采用
        if len(group) == 1:
            row = group.iloc[0].to_dict()
            row["选择方式"] = "唯一选项"
            result_rows.append(row)
            continue

        # 多选项处理流程
        options = list(zip(group['HS_Code_US'], group['Desc']))
        choice_idx = select_best_match(product, options)

        # 简化后的兜底规则（仅用HS编码位数）
        if not choice_idx or not 1 <= choice_idx <= len(options):

            # 计算整个HS编码中的数字字符数量 选择六位的
            group['编码位数'] = group['HS_Code_US'].str.count(r'\d')
            # 按编码位数升序排序
            group = group.sort_values(by='编码位数', ascending=True)
            choice_idx = 1
            select_method = "兜底规则(编码位数)"
        else:
            select_method = "模型选择"

        # 记录最终选择
        selected_row = group.iloc[choice_idx - 1].to_dict()
        selected_row["选择方式"] = select_method
        result_rows.append(selected_row)

        if select_method.startswith("兜底规则"):
            fallback_products.append(product)

    # 输出结果
    result_df = pd.DataFrame(result_rows)
    print(f"处理后数据量: {len(result_df)}行")
    print(f"触发兜底规则的产品数: {len(fallback_products)}")
    result_df.to_excel(output_file, index=False)



if __name__ == "__main__":
    # 文件路径
    input_file = "中国俄罗斯海关编码映射表_第一组.xlsx"
    output_file = "中国俄罗斯海关编码映射表_第一组_匹配后.xlsx"

    # 执行处理
    process_hs_mapping(input_file, output_file)