import sys
import openpyxl
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side, numbers
import tkinter as tk
from tkinter import filedialog

class honmeDialog:

    def __init__(self):
        self.root = tk.Tk()
        self.root.resizable(0, 0)
        self.root.attributes('-topmost', True)
        self.root.title("选择文件")
        self.root.geometry("500x350+500+100")

        self.tip = tk.Label(self.root, wraplength="400", text="", fg="red", font=("宋体", 15))
        self.tip.place(x=50, y=10)
        self.msg = tk.Label(self.root, wraplength="400", text="", fg="green", font=("宋体", 15))
        self.msg.place(x=90, y=30)
        # 选择证书统计表
        tk.Label(self.root, text="选择证书统计表：", font=("宋体", 15)).place(x=90, y=88)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15), command=self.on_select_1)
        self.select_button.place(x=290, y=87)

        # 生成文件路径
        tk.Label(self.root, text="生成文件路径：", font=("宋体", 15)).place(x=90, y=160)
        self.select_button = tk.Button(self.root, text="选择", width="10", font=("宋体", 15),
                                       command=self.on_select_dir)
        self.select_button.place(x=290, y=160)


        # 确认按钮
        self.ok_button = tk.Button(self.root, text="生成", width="10", font=("宋体", 15), command=self.on_ok)
        self.ok_button.place(x=290, y=250)

        # 结果
        self.resultFlag = False
        self.filePath_1 = ""
        self.fileDir = ""

        self.tempResultMap = {}
        self.result = []
        self.titleIndexMap = {9: "经济金融类", 10: "国际贸易融资类", 11: "财务审计类",
                              12: "风险管理类", 13: "法律类", 14: "计算机类",
                              15: "人力资源类", 16: "其他", 17: "职称类证书"}



    def on_ok(self):
        if(self.filePath_1 == ""):
            self.tip.config(text=f"请选择选择证书统计表")
            return
        if (self.fileDir == ""):
            self.tip.config(text=f"请选择生成文件目录")
            return
        self.tip.config(text=f"")
        self.msg.config(text=f"开始执行,请稍后...")

        self.fill_data()
        # self.do_excel(self.filePath_1, self.filePath_2, self.filePath_3, self.filePath_4, self.user_entry_now.get(), self.user_entry_now.get())
        self.resultFlag = True
        self.on_close()

    def on_cancel(self):
        self.resultFlag = False
        self.on_close()

    def on_close(self):
        self.root.destroy()

    def show(self):
        self.root.mainloop()



    def on_select_1(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.tip.config(text=f"路径:{file_path}")
            self.filePath_1 = file_path
            self.do_excel(self.filePath_1)

    def on_select_dir(self):
        selected_path = filedialog.askdirectory()

        # 如果选择了路径，则打印路径
        if selected_path:
            self.tip.config(text=f"选择的路径:{selected_path}")
            self.fileDir = selected_path


    # -------------------------------------------------------------

    def fill_data(self):
        # 输出文件路径（新的Excel表格）
        output_file_path = self.fileDir + "/人员持证汇总表.xlsx"
        print(self.tempResultMap)
        itemResult = []
        for key, value in self.tempResultMap.items():
            itemResult.append(value[0])
            itemResult.append(value[1])
            itemResult.append(value[3])
            itemResult.append(value[6])
            itemResult.append(value[7])
            itemResult.append(value[8])




    def do_excel(self, input_file_path):
        # 打开Excel文件
        wb_now = openpyxl.load_workbook(input_file_path, data_only=True)
        # 选择倒数第一个工作表
        sheet_now = wb_now.worksheets[-1]

        self.read_excel(sheet_now)
        # 关闭Excel文件
        wb_now.close()

    def read_project_row_num(self, sheet, target_project):
        target_num = 0
        column_number = 1
        for cell in sheet.iter_cols(min_col=column_number, max_col=column_number, values_only=True):
            for row_number, value in enumerate(cell, start=1):
                if target_project in str(value):
                    target_num = row_number
                    return target_num
        return target_num




    # 读取行数
    def read_excel(self, sheet):
        for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            # 遍历每一行
            if row_number > 1 and row[0] is not None:
                key = row[8]
                if key is not None:
                    if key not in self.tempResultMap:
                        self.tempResultMap[key] = row
                    else:
                        preRow = self.tempResultMap.get(key, "default")
                        if preRow[1] < row[1]:
                            self.tempResultMap[key] = row









    # 设置列宽
    def set_excel_col_width(self, worksheet):
        for i in range(1, 80):
            if i == 39 or i == 44 or i == 53:
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 36
            elif i == 79:
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 54
            else:
                worksheet.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 18

    # 设置表头格式
    def set_excel_head_1_style(self, worksheet):
        # 表头第一行合并范围
        merge_row_1 = [(1, 1, 79, 1)]
        start_col = 1
        start_row = 1
        end_col = 79
        end_row = 1
        merge_range_str = f'{openpyxl.utils.get_column_letter(start_col)}{start_row}:{openpyxl.utils.get_column_letter(end_col)}{end_row}'
        worksheet.merge_cells(merge_range_str)
        # 设置第一行样式
        font_row_1 = Font(name="宋体", size=20)
        alignment_row_1 = Alignment(vertical="center", horizontal="center", wrap_text=True)
        worksheet.cell(1, 1).font = font_row_1
        worksheet.cell(1, 1).alignment = alignment_row_1
        worksheet.cell(1, 1, '成本支出情况表')

    # 设置表头第二行
    def set_excel_head_2_style(self, worksheet):
        font_row_1 = Font(name="宋体", size=12)
        alignment_row_1 = Alignment(vertical="center", horizontal="left", wrap_text=True)
        worksheet.cell(2, 78).font = font_row_1
        worksheet.cell(2, 78).alignment = alignment_row_1
        worksheet.column_dimensions[openpyxl.utils.get_column_letter(78)].width = 18
        worksheet.cell(2, 78, '单位：万元')

    # 设置表头第四行格式
    def set_excel_head_4_style(self, worksheet):
        font_row = Font(name="宋体", size=12)
        alignment_row = Alignment(vertical="center", horizontal="center", wrap_text=True)
        # fill = PatternFill(fill_type="solid", fgColor="808080")
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"),
                        bottom=Side(style="thin"))
        # 表头合并范围
        # (start_col, start_row, end_col, end_row)

        for i in range(1, 80):
            worksheet.cell(4, i).border = border

        value_dict = {2: "业务及管理费", 6: "职工工资", 12: "正式工人数", 15: "劳务工人数", 18: "劳务性支出",
                      22: "社保、住房公积金", 26: "工会经费+教育经费", 30: "福利费", 35: "宣传费", 40: "房租相关费用",
                      45: "物业费", 50: "水电取暖费", 55: "钞币运送费", 60: "代理公司存款营销费", 65: "折旧费及摊销",
                      70: "外包费", 75: "其他项目"}
        merge_row_list = [(2, 4, 5, 4), (6, 4, 11, 4), (12, 4, 14, 4), (15, 4, 17, 4), (18, 4, 21, 4), (22, 4, 25, 4),
                          (26, 4, 29, 4), (30, 4, 34, 4), (35, 4, 39, 4), (40, 4, 44, 4), (45, 4, 49, 4),
                          (50, 4, 54, 4), (55, 4, 59, 4), (60, 4, 64, 4), (65, 4, 69, 4), (70, 4, 74, 4),
                          (75, 4, 78, 4)]
        for merge_range in merge_row_list:
            start_col, start_row, end_col, end_row = merge_range
            merge_range_str = f'{openpyxl.utils.get_column_letter(start_col)}{start_row}:{openpyxl.utils.get_column_letter(end_col)}{end_row}'
            worksheet.merge_cells(merge_range_str)
            worksheet.cell(4, start_col).font = font_row
            worksheet.cell(4, start_col).alignment = alignment_row
            worksheet.cell(4, start_col, value_dict.get(start_col))

        worksheet.cell(4, 1).font = font_row
        worksheet.cell(4, 1).alignment = alignment_row
        worksheet.cell(4, 1, "项目")
        worksheet.cell(4, 79).font = font_row
        worksheet.cell(4, 79).alignment = alignment_row
        worksheet.cell(4, 79, "备注其他情况")

        worksheet.row_dimensions[4].height = 30

    #  设置第五行格式
    def set_excel_head_5_style(self, worksheet, date_now, date_pre):
        font_row = Font(name="宋体", size=12)
        alignment_row = Alignment(vertical="center", horizontal="center", wrap_text=True)
        # fill = PatternFill(fill_type="solid", fgColor="808080")
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"),
                        bottom=Side(style="thin"))
        col_1 = date_now
        col_2 = date_pre
        col_3 = "同比增加"
        col_4 = "同比增幅"
        data_to_insert = ['单位', col_1, col_2, col_3, col_4, col_1, col_2, col_3, col_4, '绩效考核结果', '备注原因',
                          col_1,
                          col_2, col_3, col_1, col_2, col_3,
                          col_1, col_2, col_3, col_4, col_1, col_2, col_3, col_4, col_1, col_2, col_3, col_4, col_1,
                          col_2,
                          col_3, col_4, '备注',
                          col_1, col_2, col_3, col_4, '备注', col_1, col_2, col_3, col_4, '备注', col_1, col_2, col_3,
                          col_4, '备注',
                          col_1, col_2, col_3, col_4, '备注', col_1, col_2, col_3, col_4, '备注', col_1, col_2, col_3,
                          col_4, '备注',
                          col_1, col_2, col_3, col_4, '备注', col_1, col_2, col_3, col_4, '备注', col_1, col_2, col_3,
                          col_4, ]
        worksheet.append(data_to_insert)
        # 表头合并范围
        # (start_col, start_row, end_col, end_row)

        for i in range(1, 80):
            worksheet.cell(5, i).font = font_row
            worksheet.cell(5, i).alignment = alignment_row
            worksheet.cell(5, i).border = border

        worksheet.row_dimensions[5].height = 30
        merge_range_str = f'{openpyxl.utils.get_column_letter(79)}{4}:{openpyxl.utils.get_column_letter(79)}{5}'
        worksheet.merge_cells(merge_range_str)

    def set_excel_data(self, worksheet, data):
        dataLength = len(data)
        print('我的长度', dataLength)
        font_row = Font(name="宋体", size=12)
        alignment_row = Alignment(vertical="center", horizontal="center", wrap_text=True)
        # fill = PatternFill(fill_type="solid", fgColor="808080")
        border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"),
                        bottom=Side(style="thin"))
        for item in data:
            worksheet.append(item)

        number_format_style = NamedStyle(name='number_format_style',
                                         number_format=numbers.FORMAT_NUMBER_COMMA_SEPARATED1)
        percent_format_style = NamedStyle(name='percent_format_style', number_format=numbers.FORMAT_PERCENTAGE_00)
        percent_list = [5, 9, 21, 25, 29, 33, 38, 43, 48, 53, 58, 63, 68, 73, 78]
        number_list = [12, 13, 14, 15, 16, 17]
        for i in range(6, dataLength + 6):
            worksheet.row_dimensions[i].height = 90
            for j in range(1, 80):
                if j in percent_list:
                    worksheet.cell(i, j).style = percent_format_style
                elif j not in number_list:
                    worksheet.cell(i, j).style = number_format_style
                worksheet.cell(i, j).font = font_row
                worksheet.cell(i, j).alignment = alignment_row
                worksheet.cell(i, j).border = border

    # 设置excel表格式
    def set_excel_style(self, output_file_path, data, date_now, date_pre):
        workbook = openpyxl.Workbook()
        worksheet = workbook.create_sheet("支出明细表", index=0)

        self.set_excel_col_width(worksheet)
        self.set_excel_head_1_style(worksheet)
        self.set_excel_head_2_style(worksheet)
        self.set_excel_head_4_style(worksheet)
        self.set_excel_head_5_style(worksheet, date_now, date_pre)
        self.set_excel_data(worksheet, data)

        workbook.save(output_file_path)





dialog = honmeDialog()
dialog.show()
# revar1 = dialog.resultFlag
# revar2 = dialog.filePath