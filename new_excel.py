from datetime import datetime
from openpyxl.styles import Alignment
import os
from openpyxl import load_workbook, Workbook
from cn2an import an2cn  # 用于转换数字为中文大写
from hanziconv import HanziConv  # 简体转繁体

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

def handle_handi(sheet_name, area, anzhi, buchang, qingmiao, lingxing, anzhidanjia,
                 buchangdanjia):
    """处理旱地的补偿数据"""
    new_data = [
        ['土地补偿费（户）', '亩', area, buchangdanjia, buchang],
        ['土地安置补助费', '亩', area, anzhidanjia, anzhi],
        ['耕地青苗费', '亩', area, 2200.00, qingmiao],
        ['耕地零星林木', '亩', area, 2000.00, lingxing],
    ]
    data_into_excel(sheet_name, new_data)

def handle_lindi(sheet_name, area, anzhi, buchang, anzhidanjia, buchangdanjia):
    """处理林地、建设用地、道路、沟渠的补偿数据"""
    new_data = [
        ['土地补偿费（户）', '亩', area, buchangdanjia, buchang],
        ['土地安置补助费', '亩', area, anzhidanjia, anzhi],
    ]
    data_into_excel(sheet_name, new_data)

def handle_beifen(sheet_name, number, beifen):
    """处理有主碑坟的补偿数据"""
    new_data =  [
        ['有主碑坟', '座', number, 5000.00, beifen],
    ]
    data_into_excel(sheet_name, new_data)

def handle_pufen(sheet_name, number, pufen):
    """处理有主普坟的补偿数据"""
    new_data = [
        ['有主普坟', '座', number, 3000.00, pufen],
    ]
    data_into_excel(sheet_name, new_data)

def handle_shaichang(sheet_name, area, anzhi, buchang, anzhidanjia, buchangdanjia, shaichang):
    """处理晒场硬化的补偿数据"""
    new_data = [
        ['土地补偿费（户）', 'm2', area, buchangdanjia, buchang],
        ['土地安置补助费', 'm2', area, anzhidanjia, anzhi],
        ['晒场硬化', 'm2', area, 40.00, shaichang],
    ]
    data_into_excel(sheet_name, new_data)

def handle_shuijing(sheet_name, number, shuijing):
    """处理水井的补偿数据"""
    new_data = [
        ['水井', '眼', number, 500.00, shuijing],
    ]
    data_into_excel(sheet_name, new_data)

def handle_shuiguan(sheet_name, number, shuiguan):
    """处理给水管的补偿数据"""
    new_data = [
        ['给水管', 'm', number, 7.00, shuiguan],
    ]
    data_into_excel(sheet_name, new_data)

def handle_dijiao(sheet_name, number, dijiao):
    """处理地窖的补偿数据"""
    new_data = [
        ['地窖', '座', number, 800.00, dijiao],
    ]
    data_into_excel(sheet_name, new_data)

def handle_shuichi(sheet_name, area, volume, anzhi, buchang, anzhidanjia, buchangdanjia, shuichi):
    """处理浆砌水池的补偿数据"""
    new_data = [
        ['土地补偿费（户）', '亩', area, buchangdanjia, buchang],
        ['土地安置补助费', '亩', area, anzhidanjia, anzhi],
        ['浆砌水池', 'm3', volume, 440.00, shuichi],
    ]
    data_into_excel(sheet_name, new_data)

def handle_yutang(sheet_name, area, volume, anzhi, buchang,anzhidanjia, buchangdanjia, yutang, yumiao):
    """处理土鱼塘的补偿数据"""
    new_data = [
        ['土地补偿费（户）', '亩', area, buchangdanjia, buchang],
        ['土地安置补助费', '亩', area, anzhidanjia, anzhi],
        ['土鱼塘', 'm3', volume, 7.40, yutang],
        ['鱼损', '亩', area, 1000.00, yumiao],
    ]
    data_into_excel(sheet_name, new_data)

def handle_default(sheet_name, tree_type, area, anzhi, buchang, anzhidanjia, buchangdanjia, tree, treedanjia):
    """默认情况（其他类型）"""
    new_data = [
        ['土地补偿费（户）', '亩', area, buchangdanjia, buchang],
        ['土地安置补助费', '亩', area, anzhidanjia, anzhi],
        [tree_type, '亩', area, treedanjia, tree],
    ]
    data_into_excel(sheet_name, new_data)

def village_jiti_into_excel(sheet_name, area, cunjiti, cunjitidanjia):
    new_data = [
        ['村集体土地补偿费', '亩', area, cunjitidanjia, cunjiti],
    ]
    data_into_excel(sheet_name, new_data)

def data_into_excel(sheet_name, new_data):
    """根据类型选择对应的函数，写入 Excel"""
    # 读取 Excel 文件
    wb = load_workbook('output_file.xlsx')
    ws = wb[sheet_name]

    # 找到最后一行
    last_row = ws.max_row
    # 记录数据起始行
    start_row = last_row + 1
    # 将新数据写入工作表
    for row in new_data:
        last_row += 1
        for col, value in enumerate(row, start=1):
            # 处理小数格式
            if col == 3:  # C列 (第3列) 保留3位小数
                value = round(float(value), 3) if isinstance(value, (int, float)) else value
                cell_format = '0.000'  # Excel 显示 3 位小数
            elif col in [4, 5]:  # D、E列 (第4、5列) 保留2位小数
                value = round(float(value), 2) if isinstance(value, (int, float)) else value
                cell_format = '0.00'  # Excel 显示 2 位小数
            else:
                cell_format = None  # 其他列不做特殊格式化

            cell = ws.cell(row=last_row, column=col, value=value)

            # 设置 Excel 显示格式
            if cell_format:
                cell.number_format = cell_format  # 确保 Excel 显示小数

    # 保存 Excel
    wb.save('output_file.xlsx')
def summary_into_excel(sheet_name):
    wb = load_workbook('output_file.xlsx')
    ws = wb[sheet_name]

    # 记录数据起始行
    start_row = 4
    # 记录最后一行
    last_row = ws.max_row
    # 计算 E 列总和
    sum_formula = round(sum(ws.cell(row=row, column=5).value for row in range(start_row, last_row + 1)),2)
    # 转换为中文大写
    total_chinese = an2cn(sum_formula) + "元"
    # 转换为 **繁体中文**
    total_chinese_traditional = HanziConv.toTraditional(total_chinese)
    ws.cell(row=last_row + 1, column=5, value=sum_formula).number_format = '0.00'  # E列写公式，保留2位小数
    # 在 A 列最后一行填入 "合计"
    ws.cell(row=last_row + 1, column=1, value="合计")
    ws.cell(row=last_row + 2, column=1, value="大写")
    ws.cell(row=last_row + 2, column=2, value="人民币")
    merge_range = f"C{last_row + 2}:F{last_row + 2}"
    ws.merge_cells(merge_range)
    ws.cell(row=last_row + 2, column=3, value=total_chinese)
    # 获取单元格并设置居中
    cell = ws.cell(row=last_row + 2, column=3)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    # 保存 Excel
    wb.save('output_file.xlsx')

if __name__ == '__main__':

    header_into_excel(name='qqqqqqqqq', village_name='sssss', date=datetime.now().strftime('%Y-%m-%d'), excel_header='sdawdsdads')