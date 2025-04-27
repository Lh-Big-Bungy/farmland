import os
import win32com.client

def excel_to_pdf(input_excel, output_pdf):
    """将 Excel 转换为 PDF，并保留表格格式"""
    # 获取当前脚本的目录
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # 合并脚本目录和文件名，确保文件路径正确
    input_excel_path = os.path.join(script_dir, input_excel)
    output_pdf_path = os.path.join(script_dir, output_pdf)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False  # 不显示 Excel 窗口

    # 打开 Excel 文件
    wb = excel.Workbooks.Open(input_excel_path)

    # 获取所有有数据的 sheet
    for sheet in wb.Sheets:
        sheet.UsedRange.Borders.LineStyle = 1  # 设置边框
        # 去掉第一行和第二行的边框
        sheet.Rows(1).Borders.LineStyle = 0  # 去掉第一行的边框
        sheet.Rows(2).Borders.LineStyle = 0  # 去掉第二行的边框
        # 设置第一行字体为黑体，并放大加粗
        sheet.Rows(1).Font.Name = "黑体"  # 设置字体为黑体
        sheet.Rows(1).Font.Size = 16  # 设置字体大小为 16
        sheet.Rows(1).Font.Bold = True  # 设置字体加粗
        # 在第一行和第二行之间插入两行空白行
        sheet.Rows("2:3").Insert()  # 在第二行插入两行空白行
        row_height = 25  # 每行高度
        # 设置每行高度
        for row in sheet.UsedRange.Rows:
            row.RowHeight = row_height  # 设置每行高度为 25 磅

        sheet.PageSetup.LeftMargin = 80  # 左边距设置为 80 磅
        sheet.PageSetup.RightMargin = 65  # 右边距设置为 65 磅

        total_rows = sheet.UsedRange.Rows.Count
        total_height = total_rows * row_height  # 计算表格的总高度
        # **让 Excel 计算分页行**

        page_height = 1123
        top_margin = sheet.PageSetup.TopMargin
        bottom_margin = sheet.PageSetup.BottomMargin
        available_height = page_height - top_margin - bottom_margin
        # **计算是否分页**
        if (available_height - total_height) > 354:
            sheet.PageSetup.FooterMargin = 150  # 表格小，页脚固定 150
            sheet.PageSetup.LeftFooter = ' ' * 40 + "户主签字（盖章）："

        else:
            # 获取最后一行的行数
            last_row = sheet.UsedRange.Rows.Count
            print("last row is ", last_row)
            # 等于28行时，不用设置页脚的方法，签名会去到下一分页
            if last_row == 28:
                sheet.PageSetup.FooterMargin = 50
                # 设置右下页脚
                sheet.PageSetup.LeftFooter = ' ' * 40 + "户主签字（盖章）"
            else:# 计算目标行号（最后一行的后两行）
                target_row = last_row + 2
                # 获取中间列的列号
                last_column = sheet.Cells(1, sheet.Columns.Count).End(-4159).Column  # -4159 表示 xlToLeft
                # 在目标行的中间列设置值
                sheet.Cells(target_row, 4).Value = '户主签字（盖章）：'
    # 转换为 PDF（0 代表整个工作簿导出）
    wb.ExportAsFixedFormat(0, output_pdf_path)

    # 关闭 Excel
    wb.Close(SaveChanges=False)
    excel.Quit()

if __name__ == '__main__':
    excel_to_pdf('output_file.xlsx', 'output.pdf')