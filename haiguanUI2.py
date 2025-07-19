import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import os

# 国家列表和对应的文件名
COUNTRIES = {
    "美国": {"filename": "中美海关编码映射表_第一二三组_匹配后_带商品名称.xlsx", "hs_code_col": "HS_Code_US"},
    "俄罗斯": {"filename": "中国俄罗斯海关编码映射表_第一二三组_匹配后_带商品名称.xlsx", "hs_code_col": "HS_Code_RU"},
    "日本": {"filename": "中国日本海关编码映射表_第一二三组_去重后_匹配后_带商品名称.xlsx", "hs_code_col": "HS_Code_JP"},
    "泰国": {"filename": "中国泰国海关编码映射表_第一二三组_去重后_匹配后_带商品名称.xlsx", "hs_code_col": "HS_Code_TH"},
    "新加坡": {"filename": "中国新加坡海关编码映射表_第一二三组_去重后_匹配后_带商品名称.xlsx", "hs_code_col": "HS_Code_SG"},
    "越南": {"filename": "中国越南海关编码映射表_第一二三组_去重后_匹配后_带商品名称.xlsx", "hs_code_col": "HS_Code_VN"}
}


class HS_Code_Query_App:
    def __init__(self, root):
        self.root = root
        self.root.title("海关编码查询系统")
        self.root.geometry("1000x700")  # 增加窗口大小以容纳更多结果

        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 查询输入部分
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=10)

        ttk.Label(input_frame, text="产品名称:").pack(side=tk.LEFT, padx=5)
        self.product_entry = ttk.Entry(input_frame, width=50)
        self.product_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        search_btn = ttk.Button(input_frame, text="查询", command=self.search)
        search_btn.pack(side=tk.LEFT, padx=5)

        # 结果显示部分
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 中国海关信息 - 使用表格显示多个结果
        china_frame = ttk.LabelFrame(result_frame, text="中国海关信息")
        china_frame.pack(fill=tk.X, padx=5, pady=5)

        # 创建中国海关信息表格
        china_columns = ("海关编码", "商品描述")
        self.china_tree = ttk.Treeview(china_frame, columns=china_columns, show="headings", height=3)

        # 设置列宽
        self.china_tree.column("海关编码", width=120, anchor=tk.W)
        self.china_tree.column("商品描述", width=700, anchor=tk.W)

        # 设置表头
        for col in china_columns:
            self.china_tree.heading(col, text=col)

        # 添加滚动条
        china_scrollbar = ttk.Scrollbar(china_frame, orient=tk.VERTICAL, command=self.china_tree.yview)
        self.china_tree.configure(yscroll=china_scrollbar.set)

        self.china_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        china_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 其他国家信息表格
        country_frame = ttk.LabelFrame(result_frame, text="其他国家海关信息")
        country_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        columns = ("国家", "海关编码", "商品描述", "中文译名")
        self.tree = ttk.Treeview(country_frame, columns=columns, show="headings", height=15)

        # 设置列宽
        self.tree.column("国家", width=80, anchor=tk.W)
        self.tree.column("海关编码", width=120, anchor=tk.W)
        self.tree.column("商品描述", width=350, anchor=tk.W)
        self.tree.column("中文译名", width=350, anchor=tk.W)

        # 设置表头
        for col in columns:
            self.tree.heading(col, text=col)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(country_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 状态栏
        self.status = ttk.Label(root, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        # 加载数据
        self.data = {}
        self.load_data()

        # 设置回车键绑定
        self.product_entry.bind("<Return>", lambda event: self.search())

    def load_data(self):
        """加载所有国家的海关数据"""
        self.status.config(text="正在加载数据...")
        self.root.update()

        for country, info in COUNTRIES.items():
            filename = info["filename"]
            try:
                if os.path.exists(filename):
                    # 读取Excel文件
                    df = pd.read_excel(filename)

                    # 检查必要的列是否存在
                    required_columns = ["Product", "HS_Code_China", "商品名称", info["hs_code_col"], "Desc", "描述"]
                    missing_cols = [col for col in required_columns if col not in df.columns]

                    if missing_cols:
                        messagebox.showwarning("列缺失",
                                               f"{country}文件中缺少必要的列: {', '.join(missing_cols)}")
                        continue

                    # 存储数据
                    self.data[country] = {
                        "df": df,
                        "hs_code_col": info["hs_code_col"]
                    }
                    self.status.config(text=f"已加载: {country} 数据")
                    self.root.update()
                else:
                    messagebox.showwarning("文件缺失", f"找不到文件: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"加载 {country} 数据失败: {str(e)}")

        self.status.config(text="数据加载完成")

    def search(self):
        """执行查询操作，支持部分关键词匹配并显示多个结果"""
        keyword = self.product_entry.get().strip()
        if not keyword:
            messagebox.showinfo("提示", "请输入查询关键词")
            return

        self.status.config(text=f"正在查询: {keyword}")
        self.root.update()

        # 清空结果
        for item in self.china_tree.get_children():
            self.china_tree.delete(item)
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 用于存储所有匹配结果
        all_results = []
        china_results = []

        # 在所有国家数据中查找
        for country, info in self.data.items():
            df = info["df"]
            hs_code_col = info["hs_code_col"]

            # 使用模糊匹配（包含关键词即可）
            mask = df["Product"].str.contains(keyword, case=False, na=False)
            result = df[mask]

            if not result.empty:
                # 处理每个匹配结果
                for _, row in result.iterrows():
                    # 收集中国海关信息（去重）
                    china_info = (row["HS_Code_China"], row["商品名称"])
                    if china_info not in china_results:
                        china_results.append(china_info)

                    # 添加目标国家信息到结果列表
                    all_results.append({
                        "国家": country,
                        "海关编码": row[hs_code_col],
                        "商品描述": row["Desc"],
                        "中文译名": row["描述"]
                    })

        # 显示中国海关信息
        if china_results:
            for code, desc in china_results:
                self.china_tree.insert("", tk.END, values=(code, desc))

        # 显示其他国家信息
        if all_results:
            # 按照国家名称排序
            all_results.sort(key=lambda x: x["国家"])

            for result in all_results:
                self.tree.insert("", tk.END, values=(
                    result["国家"],
                    result["海关编码"],
                    result["商品描述"],
                    result["中文译名"]
                ))

        # 更新状态栏
        china_count = len(china_results)
        country_count = len(all_results)

        if china_count > 0 or country_count > 0:
            status_text = f"找到 {china_count} 个中国海关编码和 {country_count} 条其他国家记录"
            self.status.config(text=status_text)
        else:
            self.status.config(text="未找到匹配的产品信息")
            messagebox.showinfo("查询结果", f"未找到包含关键词 '{keyword}' 的产品")


def main():
    root = tk.Tk()
    app = HS_Code_Query_App(root)
    root.mainloop()


if __name__ == "__main__":
    main()