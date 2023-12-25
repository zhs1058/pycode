import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# 创建工作簿和新增Sheet页
filePath = "C:/Users/89704/OneDrive/Documents/我的文档/工作/RPA测试文档/test.xlsx"
sheetName = "新增Sheet5"
rows = 5
workbook = openpyxl.load_workbook(filePath)
worksheet = workbook[sheetName]


# 设置列头样式
font = Font(name="微软雅黑", size=10, italic=True, color="808080")
alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

# 设置列头单元格样式
for index in range(rows):
    curRow = index + 2
    worksheet.row_dimensions[curRow].height = 76
    for cell in worksheet[curRow]:
        cell.font = font
        cell.alignment = alignment
        cell.border = border

# 保存Excel文件
workbook.save(filePath)
