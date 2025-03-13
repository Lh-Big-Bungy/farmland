import os
import win32com.client

def excel_to_pdf(input_excel, output_pdf):
    """将 Excel 转换为 PDF，并保留表格格式"""
    # 获取当前脚本的目录
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # 合并脚本目录和文件名，确保文件路径正确
    input_excel_path = os.path.join(script_dir, input_excel)
    output_pdf_path = os.path.join(script_dir, output_pdf)

    excel = win32com.client.Dispatch("Excel.Application")
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
        # 设置每行高度
        for row in sheet.UsedRange.Rows:
            row.RowHeight = 20  # 设置每行高度为 20 磅
        # 在页眉加入户主签名
        sheet.PageSetup.RightFooter = ("户主签字（盖章）：                  ")
        sheet.PageSetup.FooterMargin = 120  # 调整页脚边距（单位：磅，默认值为 15）
        sheet.PageSetup.LeftMargin = 70  # 左边距设置为 70 磅
        sheet.PageSetup.RightMargin = 70  # 右边距设置为 70 磅
    # 转换为 PDF（0 代表整个工作簿导出）
    wb.ExportAsFixedFormat(0, output_pdf_path)

    # 关闭 Excel
    wb.Close(SaveChanges=False)
    excel.Quit()

if __name__ == '__main__':
    excel_to_pdf('output_file.xlsx', 'output.pdf')