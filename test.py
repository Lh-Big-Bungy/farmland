import os
import win32com.client

def excel_to_pdf(input_excel, output_pdf):
    """将 Excel 转换为 PDF，仅在最后一页添加页脚"""
    # 获取当前脚本的目录
    script_dir = os.path.dirname(os.path.realpath(__file__))

    # 合并脚本目录和文件名，确保文件路径正确
    input_excel_path = os.path.join(script_dir, input_excel)
    output_pdf_path = os.path.join(script_dir, output_pdf)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False  # 不显示 Excel 窗口

    try:
        # 打开 Excel 文件
        wb = excel.Workbooks.Open(input_excel_path)
        sheet = wb.Sheets(1)  # 获取第一个工作表

        # 获取分页符的位置
        page_breaks = [pb.Location.Row for pb in sheet.HPageBreaks] if sheet.HPageBreaks.Count > 0 else []
        total_pages = len(page_breaks) + 1  # 总页数 = 分页符数量 + 1

        # 如果只有一页，直接设置页脚并导出
        if total_pages == 1:
            sheet.PageSetup.CenterFooter = "页脚"
            wb.ExportAsFixedFormat(0, output_pdf_path)
            print("成功导出为 PDF，总页数: 1")
            return

        # 如果有多页，拆分内容到多个工作表
        for page in range(total_pages):
            # 复制原始工作表
            new_sheet = sheet.Copy()
            new_sheet.Name = f"Page{page + 1}"

            # 删除不需要的行
            if page == 0:
                # 第一页：删除分页符之后的内容
                start_row = 1
                end_row = page_breaks[page] - 1
                new_sheet.Rows(f"{end_row + 1}:{new_sheet.Rows.Count}").Delete()
            elif page == total_pages - 1:
                # 最后一页：删除分页符之前的内容
                start_row = page_breaks[page - 1]
                new_sheet.Rows(f"1:{start_row - 1}").Delete()
                # 设置最后一页的页脚
                new_sheet.PageSetup.CenterFooter = "最后一页页脚"
            else:
                # 中间页：删除分页符之前和之后的内容
                start_row = page_breaks[page - 1]
                end_row = page_breaks[page] - 1
                new_sheet.Rows(f"1:{start_row - 1}").Delete()
                new_sheet.Rows(f"{end_row + 1}:{new_sheet.Rows.Count}").Delete()

        # 删除原始工作表
        sheet.Delete()

        # 导出为 PDF
        wb.ExportAsFixedFormat(0, output_pdf_path)

        print(f"成功导出为 PDF，总页数: {total_pages}")
    except Exception as e:
        print(f"转换过程中出现错误: {e}")
    finally:
        # 关闭 Excel 文件
        wb.Close(SaveChanges=False)
        excel.Quit()

# 示例调用
if __name__ == '__main__':
    excel_to_pdf('output_file.xlsx', 'output_pdf.pdf')