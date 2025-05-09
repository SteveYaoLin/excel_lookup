import sys
import pandas as pd
import logging

def setup_logging():
    logging.basicConfig(
        filename="lookup.log",
        level=logging.INFO,
        format="%(asctime)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

def search_in_excel(file_path, keywords):
    try:
        df = pd.read_excel(file_path)
        print("实际列名:", df.columns.tolist())  # 调试信息
    except Exception as e:
        logging.error(f"读取文件失败: {e}")
        return

    # 修改为实际列名（示例）
    required_columns = ["物料名称", "物料描述"]
    if not all(col in df.columns for col in required_columns):
        logging.error(f"文件缺少必要列，实际列名: {df.columns.tolist()}")
        return

    keyword_list = keywords.split("_")[1:]
    logging.info(f"搜索关键字: {keyword_list}")

    matched_rows = df[
        df["物料描述"].apply(
            lambda x: all(
                keyword.lower() in str(x).lower() for keyword in keyword_list
            )
        )
    ]

    if not matched_rows.empty:
        for _, row in matched_rows.iterrows():
            b_value = row["物料名称"]
            logging.info(f"找到匹配项: {b_value}")
    else:
        logging.info("未找到匹配项")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: python lookup.py <Excel路径> <关键字>")
        sys.exit(1)

    setup_logging()
    excel_path = sys.argv[1]
    keywords = sys.argv[2]
    search_in_excel(excel_path, keywords)