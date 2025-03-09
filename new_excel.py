from datetime import datetime
from openpyxl.styles import Alignment
import os
from openpyxl import load_workbook, Workbook

def header_into_excel(name, village_name, date, excel_header):
    # **Step 1: 定义 Excel 文件名**
    file_name = 'output_file.xlsx'

    # **Step 2: 获取当前应该使用的 sheet 编号**
    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        existing_sheets = wb.sheetnames
        flag = True
        max_num = 1  # 如果找不到，则默认从 1 开始
        for sheet in existing_sheets:
            if sheet.startswith('Sheet'):
                try:
                    sheet_num = int(sheet.replace('Sheet', ''))
                    max_num = max(max_num, sheet_num)
                except ValueError:
                    continue  # 如果 Sheet 名不是数字结尾，就跳过

        num = max_num + 1  # 下一个 Sheet 号
        sheet_name = f'Sheet{num}'
    else:
        flag = False
        num = 1  # 如果文件不存在，则从 1 开始
        sheet_name = 'Sheet'
        wb = Workbook()  # 创建新的 Excel
        wb.save(file_name)  # 先保存文件

    # **Step 3: 打开 Excel，准备操作**
    wb = load_workbook(file_name)

    # **Step 4: 创建新的 Sheet**

    if flag:
        ws = wb.create_sheet(title=sheet_name)
    else:
        ws = wb.active
    # **Step 5: 合并 A1 到 D1，并填入表头**
    merge_range = "A1:F1"
    ws.merge_cells(merge_range)
    ws["A1"].value = excel_header
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # **Step 6: 合并 "户号" 这一列的单元格（如 A2:B2）**
    ws.merge_cells("A2:B2")  # 让第一行数据的"户号"占两格
    ws["A2"].value = f"户号0000{num}"
    ws["A3"].value = f"项目"
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C2"].value = f"户主：{name}"
    ws["B3"].value = f"单位"
    ws["D2"].value = f"住址：{village_name}"
    ws["C3"].value = f"数量"
    ws["F2"].value = f"{date}"
    ws["D3"].value = f"单价"
    ws["E3"].value = f"小计"
    ws["F3"].value = f"备注"


    wb.save(file_name)
    return sheet_name
def data_into_excel(sheet_name, type, area, anzhi=None, buchang=None, qingmiao=None, lingxing=None, shuichi=None, shuijing=None,
                    shuiguan=None, beifen=None, pufen=None, yutang=None, yumiao=None, shaichang=None, tree_type=None):
    wb = load_workbook('output_file.xlsx')
    ws = wb[sheet_name]
    # 3. 找到最后一行
    # 方法 1: 使用 max_row 属性找到最后一行
    last_row = ws.max_row  # 获取最后一行号
    if type == "旱地":
        new_data = [
            ['土地补偿费（户）', '亩', area, 1, buchang],
            ['土地安置补助费', '亩', area, 1, anzhi],
            ['耕地青苗费', '亩', area, 1, qingmiao],
            ['耕地零星林木', '亩', area, 1, lingxing],
        ]
    elif type in "林地、建设用地、道路、沟渠":
        new_data = [
            ['土地补偿费（户）', '亩', area, 1, buchang],
            ['土地安置补助费', '亩', area, 1, anzhi],
        ]
    elif type == "有主碑坟":
        new_data = [
            ['有主碑坟', '座', area, 5000, beifen],
        ]
    elif type == "有主普坟":
        new_data = [
            ['有主普坟', '座', 1, 1, 1],
        ]
    elif type == "晒场硬化":
        new_data = [
            ['土地补偿费（户）', '亩', 1, 1, 1],
            ['土地安置补助费', '亩', 1, 1, 1],
            ['晒场硬化', 'm2', 1, 1, 1],
        ]
    elif type == "水井":
        new_data = [
            ['水井', '眼', 1, 1, 1],
        ]
    elif type == "给水管":
        new_data = [
            ['给水管', 'm', 1, 1, 1],
        ]
    elif type == "地窖":
        new_data = [
            ['地窖', '座', 1, 1, 1],
        ]
    elif type == "浆砌水池":
        new_data = [
            ['土地补偿费（户）', '亩', 1, 1, 1],
            ['土地安置补助费', '亩', 1, 1, 1],
            ['浆砌水池', 'm3', 1, 1, 1],
        ]
    elif type == "土鱼塘":
        new_data = [
            ['土地补偿费（户）', '亩', 1, 1, 1],
            ['土地安置补助费', '亩', 1, 1, 1],
            ['土鱼塘', 'm3', 1, 1, 1],
            ['鱼损', '亩', 1, 1, 1],
        ]
    else:
        new_data = [
            ['土地补偿费（户）', '亩', 1, 1, 1],
            ['土地安置补助费', '亩', 1, 1, 1],
            [tree_type, '亩', 1, 1, 1],
        ]

    # 将新数据写入工作表
    for row in new_data:
        last_row += 1  # 移动到下一行
        for col, value in enumerate(row, start=1):  # 从第 1 列开始
            ws.cell(row=last_row, column=col, value=value)

    wb.save('output_file.xlsx')

if __name__ == '__main__':

    header_into_excel(name='qqqqqqqqq', village_name='sssss', date=datetime.now().strftime('%Y-%m-%d'), excel_header='sdawdsdads')