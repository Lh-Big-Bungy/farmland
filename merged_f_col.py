from openpyxl import Workbook, load_workbook
from openpyxl import load_workbook
import glob
def get_megerd_data():
    # 查找当前目录下所有符合模式的文件
    files = glob.glob("*（*）.xlsx")

    # 如果找到了至少一个文件
    if files:
        wb = load_workbook(files[0])  # 加载第一个匹配的文件
        print(f"加载的文件为: {files[0]}")
    else:
        print("未找到匹配的文件")
    # 加载 Excel 文件
    ws = wb.active

    # 获取合并单元格范围
    merged_ranges = ws.merged_cells.ranges

    # 存储被合并覆盖到的行号
    merged_rows_in_col_d = set()

    # 遍历合并区域
    for merged_range in merged_ranges:
        # 获取起始和结束行列
        min_col, min_row, max_col, max_row = merged_range.bounds

        # 如果合并区域在 B 列（第 2 列）
        if min_col == 4 and max_col == 4:
            # 将这个范围内的所有行都记录下来（通常用于标记）
            for row in range(min_row, max_row + 1):
                merged_rows_in_col_d.add(row)

    # 打印被合并的行
    print("D列中被合并的行有：", sorted(merged_rows_in_col_d))
    merged_dict = {}
    for row in sorted(merged_rows_in_col_d):
        name = ws.cell(row=row, column=2).value  # 名字是第2列
        area = ws.cell(row=row, column=5).value
        if name in merged_dict:
            merged_dict[name].append(area)
        else:
            merged_dict[name] = [area]
    print(merged_dict)
    return merged_dict

def merged_f_col(merged_dict):
# 读取 Excel 文件
    file_path = "output_file.xlsx"  # 你的 Excel 文件
    wb = load_workbook(file_path)
    # 获取所有 sheet 名称，并排除最后一个
    sheets = wb.sheetnames

    for sheet in sheets:
        if sheet in merged_dict:
            col_area_dict = {}
            ws = wb[sheet]
            c_col = ws["C"]
            # print(f"\n📄 Sheet: {sheet} - C列所有值:")
            start_end_list = []
            merged_list = []
            for i in merged_dict[sheet]:
                start_index = False
                end_index = False
                s_e_list = []
                for cell in c_col[3:-1]:
                    row = cell.row
            #        print(cell.coordinate, "->", cell.value)
                    if cell.value == i and not start_index:  #  获取开始的行数
                        start_index = row
                    elif cell.value != i and start_index and not end_index:
                        if cell.value == None:   # 若数据在末尾，需要把无数据的那一行加入再减去1即得到该行数
                            end_index = row - 1
                        else:
                            end_index = row  # 若数据不是在最后一行，直接获取当前行数
                s_e_list.append(start_index)
                s_e_list.append(end_index)
            #    print(s_e_list)
                start_end_list.append(s_e_list)
            # print(start_end_list)
            for x in range(len(start_end_list) - 1):
                if start_end_list[x][1] == start_end_list[x+1][0]:  # 行号尾首相连表明是挨着的，是同一块土地
                    if start_end_list[x][1] not in merged_list:
                        merged_list.append(start_end_list[x][0])
                        merged_list.append(start_end_list[x+1][1])
            # print(99999999, merged_list)
            ws.merge_cells(f"F{merged_list[0]}:F{merged_list[-1]}")
            wb.save(file_path)

def f_col_merged_run():
    merged_dict = get_megerd_data()

    merged_f_col(merged_dict)
if __name__ == '__main__':
    f_col_merged_run()