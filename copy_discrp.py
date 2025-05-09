import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def process_files(excel2_path):
    # 选择第一个Excel文件（处理后的文件）
    root = tk.Tk()
    root.withdraw()
    excel1_path = filedialog.askopenfilename(
        title="选择第一个Excel文件（已处理文件）",
        filetypes=[("Excel文件", "*.xlsx *.xls")]
    )
    if not excel1_path:
        print("未选择第一个Excel文件。")
        return

    try:
        # 读取第一个Excel文件（假设无标题行）
        df1 = pd.read_excel(excel1_path, header=None)
    except Exception as e:
        print(f"读取第一个文件失败: {e}")
        return

    try:
        # 读取第二个Excel文件（假设无标题行）
        df2 = pd.read_excel(excel2_path, header=None)
    except Exception as e:
        print(f"读取第二个文件失败: {e}")
        return

    # 检查列是否存在
    if df1.shape[1] < 12:
        print("第一个文件缺少L列（第12列）")
        return
    if df2.shape[1] < 4:
        print("第二个文件缺少C/D列（第3/4列）")
        return

    # 构建第二个文件的C列到D列的映射字典（保留第一个匹配项）
    c_to_d = {}
    for idx, row in df2.iterrows():
        c_value = str(row[2]) if pd.notnull(row[2]) else ""  # C列是第3列（索引2）
        if c_value not in c_to_d:
            d_value = row[3] if pd.notnull(row[3]) else ""    # D列是第4列（索引3）
            c_to_d[c_value] = d_value

    # 处理第一个文件的M列（第13列，索引12）
    m_col = []
    for idx, row in df1.iterrows():
        l_value = str(row[11]) if pd.notnull(row[11]) else ""  # L列是第12列（索引11）
        # 查找匹配项
        matched_d = c_to_d.get(l_value, "")
        m_col.append(matched_d)

    # 更新第一个文件的M列
    if df1.shape[1] > 12:
        df1.iloc[:, 12] = m_col
    else:
        df1[12] = m_col

    # 保存第一个文件（覆盖原文件）
    try:
        if excel1_path.endswith(".xlsx"):
            df1.to_excel(excel1_path, index=False, header=False, engine="openpyxl")
        elif excel1_path.endswith(".xls"):
            df1.to_excel(excel1_path, index=False, header=False, engine="xlwt")
        print(f"文件已更新保存至: {excel1_path}")
    except Exception as e:
        print(f"保存文件失败: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("请通过命令行参数指定第二个Excel文件路径")
        print("示例: python copy_discrp.py 'path/to/second_file.xlsx'")
        sys.exit(1)
    
    excel2_path = sys.argv[1]
    process_files(excel2_path)