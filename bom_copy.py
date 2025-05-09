import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def process_excel():
    # 弹出文件选择对话框
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="选择Excel文件",
        filetypes=[("Excel文件", "*.xlsx *.xls")]
    )
    if not file_path:
        print("未选择文件。")
        return

    try:
        # 读取Excel文件（假设无标题行）
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"读取文件失败: {e}")
        return

    # 处理B列（第2列，索引1）
    try:
        b_col = df.iloc[:, 1]
    except IndexError:
        print("Excel文件中没有B列。")
        return

    # 转换为小写并统计
    lower_b = b_col.astype(str).str.lower()
    counts = lower_b.value_counts()
    duplicates = counts[counts >= 2].index.tolist()

    if not duplicates:
        print("B列没有发现重复内容。")
        return

    # 记录首次出现位置
    first_occurrence = {}
    for idx, value in enumerate(lower_b):
        if value not in first_occurrence:
            first_occurrence[value] = idx

    # 按首次出现顺序排序重复项
    sorted_duplicates = sorted(
        [value for value in duplicates],
        key=lambda x: first_occurrence[x]
    )

    # 创建编号字典
    group_dict = {value: i+1 for i, value in enumerate(sorted_duplicates)}

    # 准备L列数据
    if df.shape[1] > 11:
        l_col = df.iloc[:, 11].tolist()
    else:
        l_col = [''] * len(df)

    # 更新重复行的L列
    for idx, value in enumerate(lower_b):
        if value in group_dict:
            l_col[idx] = group_dict[value]

    # 写入DataFrame
    if df.shape[1] > 11:
        df.iloc[:, 11] = l_col
    else:
        df[11] = l_col

    # 生成输出文件路径
    base_name, ext = os.path.splitext(file_path)
    output_path = f"{base_name}_处理结果{ext}"

    # 保存处理后的文件
    try:
        if ext == ".xlsx":
            df.to_excel(output_path, index=False, header=False, engine="openpyxl")
        elif ext == ".xls":
            df.to_excel(output_path, index=False, header=False, engine="xlwt")
        else:
            print("不支持的文件格式。")
            return
        print(f"文件已保存至: {output_path}")
    except Exception as e:
        print(f"保存文件失败: {e}")

if __name__ == "__main__":
    process_excel()