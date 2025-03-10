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

    # 获取所有 sheet
    for sheet in wb.Sheets:
        # 设置页脚内容：每页底部加“户主签名”
        sheet.PageSetup.CenterFooter = "户主签名："
    # 转换为 PDF（0 代表整个工作簿导出）
    wb.ExportAsFixedFormat(0, output_pdf_path)

    # 关闭 Excel
    wb.Close(SaveChanges=False)
    excel.Quit()

