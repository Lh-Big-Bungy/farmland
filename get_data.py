import os
import glob
import pandas as pd
import re
def get_data():
    # 获取当前目录
    current_dir = os.getcwd()

    # 查找所有 .xlsx 文件（排除临时文件，如 `~$` 开头的文件）
    xlsx_files = [f for f in glob.glob(os.path.join(current_dir, "*.xlsx")) if not os.path.basename(f).startswith("~$")]

    # 检查文件数量
    if len(xlsx_files) == 0:
        print("❌ 未找到任何 .xlsx 文件")
    elif len(xlsx_files) > 1:
        print("❌ 目录中有多个 .xlsx 文件，请只保留一个")
    else:
        # 读取唯一的 Excel 文件
        file_path = xlsx_files[0]
        print(f"✅ 读取文件: {file_path}")
        file_name = os.path.basename(file_path)
        village_name = file_name.split('.')[0].split('（')[-1].replace('）', '')
        print('村庄名称：', village_name)
        df = pd.read_excel(file_path, header=None)  # 默认读取第一个 sheet
        # 选择 B、C、E 列（索引分别是 1, 2, 4），从第 3 行（索引 2）开始
        selected_data = df.iloc[2:, [1, 2, 3, 4]]
        # 替换 NaN 为 None
        selected_data = selected_data.where(pd.notna(selected_data), None)
        # 转换为列表，每行一个子列表
        result = selected_data.values.tolist()
        date = result[-1][-1].split('：')[1]
        first_cell = df.iloc[0, 0].split('面积')[0]
        # 去除换行符（适用于多行合并的情况）
        first_cell = first_cell.replace("\n", "").replace("\r", "")
        # 或者使用 strip() 仅去除行首/行尾的换行符
        first_cell = first_cell.strip() + "补偿兑付表"
        first_cell = re.sub(r'（[^（）]*）$', '', first_cell)
        print(first_cell)
        print(date)
        print(result)
        print(result[-3])
        return village_name, result, date, first_cell



if __name__ == '__main__':
    village, data, date, excel_header = get_data()
    for i in data:
        if i[0] == "胡方明":
            print(i)
