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
        self.root.geometry("800x600")

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

        # 中国海关信息
        china_frame = ttk.LabelFrame(result_frame, text="中国海关信息")
        china_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(china_frame, text="海关编码:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.cn_code = ttk.Label(china_frame, text="", font=("Arial", 10, "bold"))
        self.cn_code.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        ttk.Label(china_frame, text="商品描述:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.cn_desc = ttk.Label(china_frame, text="", wraplength=600, font=("Arial", 10))
        self.cn_desc.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)

        # 其他国家信息表格
        country_frame = ttk.LabelFrame(result_frame, text="其他国家海关信息")
        country_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        columns = ("国家", "海关编码", "商品描述", "中文译名")
        self.tree = ttk.Treeview(country_frame, columns=columns, show="headings", height=10)

        # 设置列宽
        self.tree.column("国家", width=80, anchor=tk.W)
        self.tree.column("海关编码", width=120, anchor=tk.W)
        self.tree.column("商品描述", width=250, anchor=tk.W)
        self.tree.column("中文译名", width=250, anchor=tk.W)

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
        """执行查询操作"""
        product_name = self.product_entry.get().strip()
        if not product_name:
            messagebox.showinfo("提示", "请输入产品名称")
            return

        self.status.config(text=f"正在查询: {product_name}")
        self.root.update()

        # 清空结果
        self.cn_code.config(text="")
        self.cn_desc.config(text="")
        for item in self.tree.get_children():
            self.tree.delete(item)

        found_in_china = False
        found_count = 0

        # 在所有国家数据中查找
        for country, info in self.data.items():
            df = info["df"]
            hs_code_col = info["hs_code_col"]

            # 查找匹配的产品
            result = df[df["Product"] == product_name]

            if not result.empty:
                row = result.iloc[0]
                found_count += 1

                # 获取中国海关信息
                if not found_in_china:
                    self.cn_code.config(text=row["HS_Code_China"])
                    self.cn_desc.config(text=row["商品名称"])
                    found_in_china = True

                # 获取目标国家信息
                country_info = {
                    "国家": country,
                    "海关编码": row[hs_code_col],
                    "商品描述": row["Desc"],
                    "中文译名": row["描述"]
                }

                # 添加到表格
                self.tree.insert("", tk.END, values=(
                    country_info["国家"],
                    country_info["海关编码"],
                    country_info["商品描述"],
                    country_info["中文译名"]
                ))

        if found_count > 0:
            self.status.config(text=f"找到 {found_count} 个国家的匹配结果")
        else:
            self.status.config(text="未找到匹配的产品信息")
            messagebox.showinfo("查询结果", f"未找到产品: {product_name}")


def main():
    root = tk.Tk()
    app = HS_Code_Query_App(root)
    root.mainloop()


if __name__ == "__main__":
    main()