import pandas as pd
import os
import re

def clean_path(path):
    """清除路径中的非法字符"""
    cleaned = re.sub(r'[\x00-\x1F\x7F]', '', path).strip()
    return os.path.normpath(cleaned)

def main():
    try:
        excel_a_path = clean_path(input("请输入ExcelA文件路径："))
        excel_b_path = clean_path(input("请输入ExcelB文件路径："))

        print("\n[!] 警告：该操作将直接修改ExcelB文件")
        confirm = input("[!] 确认要继续吗？(y/n): ").lower()
        if confirm != 'y':
            print("操作已取消")
            return

        if not all(map(os.path.isfile, [excel_a_path, excel_b_path])):
            missing = [p for p in [excel_a_path, excel_b_path] if not os.path.exists(p)]
            raise FileNotFoundError(f"文件不存在：{', '.join(missing)}")

        # 修正1：指定所有处理列为字符串类型
        df_a = pd.read_excel(excel_a_path, header=None, 
                            dtype={0: str, 1: str, 2: str},
                            engine='openpyxl')
        df_b = pd.read_excel(excel_b_path, header=None,
                            dtype={0: str, 1: str, 2: str},
                            engine='openpyxl')

        df_a[2] = df_a[2].astype(str).str.strip().replace(['nan', ''], pd.NA)
        df_b[2] = df_b[2].astype(str).str.strip().replace(['nan', ''], pd.NA)

        missing_pins = []

        for index_a, row_a in df_a.iterrows():
            pin = row_a[2]
            if pd.isna(pin):
                continue

            mask = df_b[2].astype(str) == str(pin)
            if mask.any():
                # 确保赋值类型一致性
                df_b.loc[mask, 0] = str(row_a[0])
                df_b.loc[mask, 1] = str(row_a[1])
            else:
                missing_pins.append(pin)

        # 修正2：移除encoding参数
        df_b.to_excel(
            excel_b_path,
            index=False,
            header=False,
            engine='openpyxl'
        )

        print("\n操作完成！以下为处理结果：")
        print(f"已更新文件：{excel_b_path}")
        if missing_pins:
            print("\n未找到的PIN：")
            print('\n'.join(f"• {pin}" for pin in set(missing_pins)))
        else:
            print("\n所有PIN均已成功匹配并更新")

    except Exception as e:
        print(f"\n[错误] 处理失败：{str(e)}")
        print("建议：1. 检查文件是否被其他程序打开 2. 确认文件路径正确")

if __name__ == "__main__":
    main()