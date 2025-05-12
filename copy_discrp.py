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
    if df2.shape[1] < 8:
        print("第二个文件缺少H列（第8列）")
        return

    # 构建C列的完整映射字典（记录所有匹配项）
    c_mapping = {}
    for idx, row in df2.iterrows():
        c_value = str(row[2]).strip() if pd.notnull(row[2]) else ""
        if c_value not in c_mapping:
            c_mapping[c_value] = []
        # 存储D列、H列和原始行号
        c_mapping[c_value].append({
            'd_value': row[3] if pd.notnull(row[3]) else "",
            'h_value': row[7] if pd.notnull(row[7]) else 0,
            'row_num': idx
        })

    # 处理第一个文件的M列和N列
    m_col, n_col = [], []
    for idx, row in df1.iterrows():
        l_value = str(row[11]).strip() if pd.notnull(row[11]) else ""
        matched_d = ""
        matched_h = ""
        
        if l_value in c_mapping:
            candidates = c_mapping[l_value]
            
            # 优先查找H列不为0的项
            non_zero = [x for x in candidates if x['h_value'] != 0]
            if len(non_zero) > 0:
                # 取第一个H列不为0的项
                target = non_zero[0]
            else:
                # 全部为0时取最后一项
                target = candidates[-1]
            
            matched_d = target['d_value']
            matched_h = target['h_value']

        m_col.append(matched_d)
        n_col.append(matched_h)

    # 更新列数据
    df1[12] = m_col  # M列
    df1[13] = n_col  # N列

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