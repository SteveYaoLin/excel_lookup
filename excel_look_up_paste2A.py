import pandas as pd
import os
import re

def clean_path(path):
    """清除路径中的非法字符"""
    cleaned = re.sub(r'[\x00-\x1F\x7F]', '', path).strip()
    return os.path.normpath(cleaned)

def main():
    try:
        # 获取文件路径
        excel_a_path = clean_path(input("请输入ExcelA文件路径："))
        excel_b_path = clean_path(input("请输入ExcelB文件路径："))

        # 安全警告
        print("\n[!] 警告：该操作将直接修改ExcelA文件")
        confirm = input("[!] 确认要继续吗？(y/n): ").lower()
        if confirm != 'y':
            print("操作已取消")
            return

        # 验证文件存在性
        for path in [excel_a_path, excel_b_path]:
            if not os.path.isfile(path):
                raise FileNotFoundError(f"文件不存在：{path}")

        # 读取Excel（D-K列保留原始数据类型）
        df_a = pd.read_excel(excel_a_path, header=None, 
                            dtype={2: str},  # 仅强制C列为字符串
                            engine='openpyxl')
        df_b = pd.read_excel(excel_b_path, header=None,
                            dtype={2: str},  # 仅强制C列为字符串
                            engine='openpyxl')

        # 数据清洗
        df_a[2] = df_a[2].astype(str).str.strip().replace(['nan', ''], pd.NA)
        df_b[2] = df_b[2].astype(str).str.strip().replace(['nan', ''], pd.NA)

        missing_pins = []
        updated_count = 0

        # 构建ExcelB的查找字典（加速查询）
        b_dict = {}
        for index_b, row_b in df_b.iterrows():
            pin = row_b[2]
            if pd.notna(pin):
                b_dict.setdefault(pin, []).append(index_b)

        # 遍历ExcelA进行更新
        for index_a, row_a in df_a.iterrows():
            pin = row_a[2]
            if pd.isna(pin):
                continue

            # 查找匹配项
            matched_indices = b_dict.get(str(pin), [])
            
            if matched_indices:
                # 取第一个匹配项（如需取全部可遍历matched_indices）
                index_b = matched_indices[0]
                
                # 复制D-K列（列索引3到10）
                df_a.iloc[index_a, 3:11] = df_b.iloc[index_b, 3:11].values
                updated_count += 1
            else:
                missing_pins.append(pin)

        # 保存修改到ExcelA
        df_a.to_excel(
            excel_a_path,
            index=False,
            header=False,
            engine='openpyxl'
        )

        # 输出报告
        print("\n操作完成！以下为处理结果：")
        print(f"已更新文件：{excel_a_path}")
        print(f"成功更新行数：{updated_count}")
        if missing_pins:
            print("\n未找到的PIN：")
            print('\n'.join(f"• {pin}" for pin in set(missing_pins)))
        else:
            print("\n所有PIN均已成功匹配并更新")

    except Exception as e:
        print(f"\n[错误] 处理失败：{str(e)}")
        print("建议排查步骤：")
        print("1. 确认文件未被其他程序打开")
        print("2. 检查Excel列数是否满足要求（需至少包含K列）")
        print("3. 验证C列值是否包含前导/尾随空格")

if __name__ == "__main__":
    main()