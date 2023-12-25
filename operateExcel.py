import openpyxl
import pandas as pd
from decimal import Decimal, getcontext
# 原始表和最终表银行名称对照字典
bank_name_map = {"东城区支行": "东城", "西城区支行": "西城", "金融大街支行": "金融街", "朝阳区支行": "朝阳",
                 "望京支行": "望京", "海淀区支行": "海淀", "中关村支行": "中关村", "丰台区支行": "丰台",
                 "石景山区支行": "石景山", "大兴区支行": "大兴", "通州区支行": "通州", "房山区支行": "房山",
                 "顺义区支行": "顺义", "门头沟区支行": "门头沟", "密云区支行": "密云", "延庆区支行": "延庆",
                 "平谷区支行": "平谷", "昌平区支行": "昌平", "怀柔区支行": "怀柔", "亦庄支行": "亦庄"}

total_now = 0.0
total_pre = 0.0
other_now = 0.0
other_pre = 0.0
# 读取项目所在的行号
def read_project_row_num(sheet, target_project):
    target_num = 0
    column_number = 1
    for cell in sheet.iter_cols(min_col=column_number, max_col=column_number, values_only=True):
        for row_number, value in enumerate(cell, start=1):
            if target_project in str(value):
                target_num = row_number
                return target_num
    return target_num


# 读取人员表所在的行号和列号并获取值
def read_person_row_col(sheet, row_name, col_name):
    row_num = 0
    col_num = 0
    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        # 遍历当前行的每一列
        for col_number, cell_value in enumerate(row, start=1):
            # 检查单元格的值是否包含目标字符串
            if col_name in str(cell_value):
                col_num = col_number
                break
            if row_name in str(cell_value):
                row_num = row_number
                break
    cell_value = sheet.cell(row=row_num, column=col_num).value
    if cell_value is None:
        cell_value = 0
    return cell_value


# 读取支行所在列号和行号
def read_excel(sheet, target_bank_name):
    bank_row_num = 0
    bank_col_num = 0
    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        # 遍历当前行的每一列
        for col_number, cell_value in enumerate(row, start=1):
            # 检查单元格的值是否包含目标字符串
            if target_bank_name in str(cell_value):
                bank_row_num = row_number
                bank_col_num = col_number
                break
    return bank_row_num, bank_col_num + 1


# 获取某支行某一个项目的本年累计数
def get_amount_bank_item(sheet, row_num, col_num):
    cell_value = sheet.cell(row=row_num, column=col_num).value
    return cell_value


# 获取单元格值
def get_cell_value(sheet, item_name, bank_name):
    item_row_num = read_project_row_num(sheet, item_name)
    bank_row_num, bank_col_num = read_excel(sheet, bank_name)

    return get_amount_bank_item(sheet, item_row_num, bank_col_num)


# 获取sheet页
def get_sheet(wb, sheet_name_target):
    try:
        # 获取所有 sheet 页的名称
        sheet_names = wb.sheetnames

        # 遍历 sheet 页名称，查找包含指定字符串的 sheet 页
        for sheet_name in sheet_names:
            if sheet_name_target in sheet_name:
                return wb[sheet_name]
        return None
    finally:
        print("选择Sheet页结束")
        # 关闭Excel文件
        # wb.close()


# 计算业务及管理费
def get_item_1(sheet_now, sheet_pre, bank_key):
    global total_now
    global  total_pre
    # 获取业务及管理费
    business_expense_now = float(get_cell_value(sheet_now, "业务及管理费", bank_key))

    business_expense_pre = float(get_cell_value(sheet_pre, "业务及管理费", bank_key))

    # 获取储蓄存款代理费
    store_expense_now = float(get_cell_value(sheet_now, "储蓄存款代理费", bank_key))
    store_expense_pre = float(get_cell_value(sheet_pre, "储蓄存款代理费", bank_key))
    # 获取外币存款代理费
    foreign_store_expense_now = float(get_cell_value(sheet_now, "外币存款代理费", bank_key))
    foreign_store_expense_pre = float(get_cell_value(sheet_pre, "外币存款代理费", bank_key))
    # 当期
    business_expense_now = (business_expense_now - store_expense_now - foreign_store_expense_now) / 10000
    total_now = business_expense_now
    # 去年同期
    business_expense_pre = (business_expense_pre - store_expense_pre - foreign_store_expense_pre) / 10000
    total_pre = business_expense_pre
    # 同比增加
    increase = business_expense_now - business_expense_pre
    # 同比增幅
    increase_ratio = 0
    if business_expense_pre != 0:
        increase_ratio = (increase / business_expense_pre) * 100

    result = [round(business_expense_now, 2), round(business_expense_pre, 2), round(increase, 2), str(round(increase_ratio, 2)) + '%']

    return result

# 计算职工工资
def get_item_2(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 获取职工工资
    # 当期
    wage_now = float(get_cell_value(sheet_now, "职工工资", bank_key)) / 10000
    other_now += wage_now
    # 去年同期
    wage_pre = float(get_cell_value(sheet_pre, "职工工资", bank_key)) / 10000
    other_pre += wage_pre
    # 同比增加
    increase = wage_now - wage_pre
    # 同比增幅
    increase_ratio = 0
    if wage_pre != 0:
        increase_ratio = (increase / wage_pre) * 100

    result = [round(wage_now, 2), round(wage_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '', '']

    return result

# 正式工人数
def get_item_3(sheet_now, sheet_pre, bank_key):
    r1_now_1 = int(read_person_row_col(sheet_now, bank_key, "合同用工"))
    r1_now_2 = int(read_person_row_col(sheet_now, bank_key, "保留关系"))
    r1_pre = int(read_person_row_col(sheet_pre, bank_key, "合同用工"))
    r1_now = r1_now_1 + r1_now_2
    # 同比增加
    increase = r1_now - r1_pre

    result = [r1_now, r1_pre, increase]

    return result

# 劳务工人数
def get_item_4(sheet_now, sheet_pre, bank_key):
    r1_now = int(read_person_row_col(sheet_now, bank_key, "劳务用工"))
    r1_pre = int(read_person_row_col(sheet_pre, bank_key, "劳务用工"))

    # 同比增加
    increase = r1_now - r1_pre

    result = [r1_now, r1_pre, increase]

    return result

# 劳务性支出
def get_item_5(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    labour_now = float(get_cell_value(sheet_now, "劳务性支出", bank_key)) / 10000
    other_now += labour_now
    # 去年同期
    labour_pre = float(get_cell_value(sheet_pre, "劳务性支出", bank_key)) / 10000
    other_pre += labour_pre
    # 同比增加
    increase = labour_now - labour_pre
    # 同比增幅
    increase_ratio = 0
    if labour_pre != 0:
        increase_ratio = (increase / labour_pre) * 100

    result = [round(labour_now, 2), round(labour_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%']

    return result

# 社保、住房公积金
def get_item_6(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r1_now = float(get_cell_value(sheet_now, "社会保险费", bank_key)) / 10000
    r2_now = float(get_cell_value(sheet_now, "补充养老保险费", bank_key)) / 10000
    r3_now = float(get_cell_value(sheet_now, "补充医疗保险费", bank_key)) / 10000
    r4_now = float(get_cell_value(sheet_now, "住房公积金", bank_key)) / 10000
    # 去年同期
    r1_pre = float(get_cell_value(sheet_pre, "社会保险费", bank_key)) / 10000
    r2_pre = float(get_cell_value(sheet_pre, "补充养老保险费", bank_key)) / 10000
    r3_pre = float(get_cell_value(sheet_pre, "补充医疗保险费", bank_key)) / 10000
    r4_pre = float(get_cell_value(sheet_pre, "住房公积金", bank_key)) / 10000

    r_now = r1_now + r2_now + r3_now + r4_now
    r_pre = r1_pre + r2_pre + r3_pre + r4_pre
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%']

    return result


# 工会经费+教育经费
def get_item_7(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r1_now = float(get_cell_value(sheet_now, "工会经费", bank_key)) / 10000
    r2_now = float(get_cell_value(sheet_now, "职工教育经费", bank_key)) / 10000
    # 去年同期
    r1_pre = float(get_cell_value(sheet_pre, "工会经费", bank_key)) / 10000
    r2_pre = float(get_cell_value(sheet_pre, "职工教育经费", bank_key)) / 10000

    r_now = r1_now + r2_now
    r_pre = r1_pre + r2_pre
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%']

    return result


# 福利费
def get_item_8(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "职工福利费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "职工福利费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 宣传费
def get_item_9(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "业务宣传费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "业务宣传费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 房租相关费用
def get_item_10(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r1_now = float(get_cell_value(sheet_now, "租赁费", bank_key)) / 10000
    r2_now = float(get_cell_value(sheet_now, "租赁款利息费", bank_key)) / 10000
    r3_now = float(get_cell_value(sheet_now, "使用权资产折旧费", bank_key)) / 10000
    # 去年同期
    r1_pre = float(get_cell_value(sheet_pre, "租赁费", bank_key)) / 10000
    r2_pre = float(get_cell_value(sheet_pre, "租赁款利息费", bank_key)) / 10000
    r3_pre = float(get_cell_value(sheet_pre, "使用权资产折旧费", bank_key)) / 10000

    r_now = r1_now + r2_now + r3_now
    r_pre = r1_pre + r2_pre + r3_pre
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 物业费
def get_item_11(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "物业管理费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "物业管理费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 水电取暖费
def get_item_12(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "水电取暖费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "水电取暖费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 钞币运送费
def get_item_13(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "钞币运送费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "钞币运送费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 代理公司存款营销费
def get_item_14(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "公司存款营销费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "公司存款营销费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100


    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 折旧费及摊销
def get_item_15(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r1_now = float(get_cell_value(sheet_now, "固定资产折旧费", bank_key)) / 10000
    r2_now = float(get_cell_value(sheet_now, "无形资产摊销", bank_key)) / 10000
    r3_now = float(get_cell_value(sheet_now, "长期待摊费用摊销", bank_key)) / 10000
    # 去年同期
    r1_pre = float(get_cell_value(sheet_pre, "固定资产折旧费", bank_key)) / 10000
    r2_pre = float(get_cell_value(sheet_pre, "无形资产摊销", bank_key)) / 10000
    r3_pre = float(get_cell_value(sheet_pre, "长期待摊费用摊销", bank_key)) / 10000

    r_now = r1_now + r2_now + r3_now
    r_pre = r1_pre + r2_pre + r3_pre
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 外包费
def get_item_16(sheet_now, sheet_pre, bank_key):
    global other_now
    global other_pre
    # 当期
    r_now = float(get_cell_value(sheet_now, "外包费", bank_key)) / 10000
    # 去年同期
    r_pre = float(get_cell_value(sheet_pre, "外包费", bank_key)) / 10000
    other_now += r_now
    other_pre += r_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']

    return result


# 其他项目
def get_item_17(sheet_now, sheet_pre, bank_key):
    global other_now, other_pre, total_now, total_pre


    r_now = total_now - other_now
    r_pre = total_pre - other_pre
    # 同比增加
    increase = r_now - r_pre
    # 同比增幅
    increase_ratio = 0
    if r_pre != 0:
        increase_ratio = (increase / r_pre) * 100

    result = [round(r_now, 2), round(r_pre, 2), round(increase, 2),
              str(round(increase_ratio, 2)) + '%', '']
    return result


# 正式工人数合计
def get_item_3_total(sheet_now, sheet_pre, bank_key):
    r1_now_1 = int(read_person_row_col(sheet_now, "合计", "合同用工"))
    r1_now_2 = int(read_person_row_col(sheet_now, "合计", "保留关系"))

    r1_pre_1 = int(read_person_row_col(sheet_pre, "合计", "合同用工"))
    r1_pre_2 = int(read_person_row_col(sheet_pre, "合计", "保留关系"))
    r1_now = r1_now_1 + r1_now_2
    r1_pre = r1_pre_1 + r1_pre_2
    # 同比增加
    increase = r1_now - r1_pre

    result = [r1_now, r1_pre, increase]

    return result



# 劳务工人数合计
def get_item_4_total(sheet_now, sheet_pre, bank_key):
    r1_now = int(read_person_row_col(sheet_now, "合计", "劳务用工"))

    r1_pre = int(read_person_row_col(sheet_pre, "合计", "劳务用工"))

    # 同比增加
    increase = r1_now - r1_pre

    result = [r1_now, r1_pre, increase]

    return result


# 正式工人数
def get_item_3_branch(sheet_now, sheet_pre, bank_key):
    r1_now_1 = int(read_person_row_col(sheet_now, "分行机关", "合同用工"))
    r1_now_2 = int(read_person_row_col(sheet_now, "分行机关", "保留关系"))
    r2_now_1 = int(read_person_row_col(sheet_now, "运营中心", "合同用工"))
    r2_now_2 = int(read_person_row_col(sheet_now, "运营中心", "保留关系"))
    r3_now_1 = int(read_person_row_col(sheet_now, "分行营业部", "合同用工"))
    r3_now_2 = int(read_person_row_col(sheet_now, "分行营业部", "保留关系"))

    r1_pre_1 = int(read_person_row_col(sheet_pre, "分行机关", "合同用工"))
    r1_pre_2 = int(read_person_row_col(sheet_pre, "分行机关", "保留关系"))
    r2_pre_1 = int(read_person_row_col(sheet_pre, "运营中心", "合同用工"))
    r2_pre_2 = int(read_person_row_col(sheet_pre, "运营中心", "保留关系"))
    r3_pre_1 = int(read_person_row_col(sheet_pre, "分行营业部", "合同用工"))
    r3_pre_2 = int(read_person_row_col(sheet_pre, "分行营业部", "保留关系"))
    r1_now = r1_now_1 + r1_now_2 + r2_now_1 + r2_now_2 + r3_now_1 + r3_now_2
    r1_pre = r1_pre_1 + r1_pre_2 + r2_pre_1 + r2_pre_2 + r3_pre_1 + r3_pre_2
    # 同比增加
    increase = r1_now - r1_pre

    result = [r1_now, r1_pre, increase]

    return result

# 劳务工人数
def get_item_4_branch(sheet_now, sheet_pre, bank_key):
    r1_now_1 = int(read_person_row_col(sheet_now, "分行机关", "劳务用工"))
    r1_now_2 = int(read_person_row_col(sheet_now, "运营中心", "劳务用工"))
    r1_now_3 = int(read_person_row_col(sheet_now, "分行营业部", "劳务用工"))
    r1_now = r1_now_1 + r1_now_2 + r1_now_3

    r1_pre_1 = int(read_person_row_col(sheet_pre, "分行机关", "劳务用工"))
    r1_pre_2 = int(read_person_row_col(sheet_pre, "运营中心", "劳务用工"))
    r1_pre_3 = int(read_person_row_col(sheet_pre, "分行营业部", "劳务用工"))
    r1_pre = r1_pre_1 + r1_pre_2 + r1_pre_3

    # 同比增加
    increase = r1_now - r1_pre

    result = [r1_now, r1_pre, increase]

    return result

# 计算小计
def count_total_1(list1, list2):
    result_item = []
    result_item.append("小计")
    list_len_1 = len(list1)
    list_len_2 = len(list2)
    if(list_len_1 == list_len_2):
        for i in range(1, list_len_1):
            item1 = list1[i]
            item2 = list2[i]
            if type(item1) == float and type(item2) == float:
                result_item.append(round((item2 - item1), 2))
            elif type(item1) == int and type(item2) == int:
                result_item.append(item2 - item1)
            elif type(item1) == str and type(item2) == str and '%' in item1 and '%' in item2:
                increase = result_item[i-1]
                pre = result_item[i-2]
                increase_ratio = 0
                if pre != 0:
                    increase_ratio = (increase / pre) * 100
                result_item.append(str(round(increase_ratio, 2)) + '%')
            else:
                result_item.append('')



    return result_item





# 计算分行
def count_total_2(sheet_now, sheet_pre, sheet_person_now, sheet_person_pre):
    global other_now, other_pre, total_now, total_pre

    result_item = []
    total_now = 0.0
    total_pre = 0.0
    other_now = 0.0
    other_pre = 0.0
    key = "11000013"
    result_item.append("分行")
    result_item += get_item_1(sheet_now, sheet_pre, key)
    result_item += get_item_2(sheet_now, sheet_pre, key)
    result_item += get_item_3_branch(sheet_person_now, sheet_person_pre, key)
    result_item += get_item_4_branch(sheet_person_now, sheet_person_pre, key)
    result_item += get_item_5(sheet_now, sheet_pre, key)
    result_item += get_item_6(sheet_now, sheet_pre, key)
    result_item += get_item_7(sheet_now, sheet_pre, key)
    result_item += get_item_8(sheet_now, sheet_pre, key)
    result_item += get_item_9(sheet_now, sheet_pre, key)
    result_item += get_item_10(sheet_now, sheet_pre, key)
    result_item += get_item_11(sheet_now, sheet_pre, key)
    result_item += get_item_12(sheet_now, sheet_pre, key)
    result_item += get_item_13(sheet_now, sheet_pre, key)
    result_item += get_item_14(sheet_now, sheet_pre, key)
    result_item += get_item_15(sheet_now, sheet_pre, key)
    result_item += get_item_16(sheet_now, sheet_pre, key)
    result_item += get_item_17(sheet_now, sheet_pre, key)

    return result_item

# 计算合计
def count_total_3(sheet_now, sheet_pre, sheet_person_now, sheet_person_pre):
    global other_now, other_pre, total_now, total_pre

    result_item = []
    total_now = 0.0
    total_pre = 0.0
    other_now = 0.0
    other_pre = 0.0
    key = "北京分行（合并）"
    result_item.append("合计")
    result_item += get_item_1(sheet_now, sheet_pre, key)
    result_item += get_item_2(sheet_now, sheet_pre, key)
    result_item += get_item_3_total(sheet_person_now, sheet_person_pre, key)
    result_item += get_item_4_total(sheet_person_now, sheet_person_pre, key)
    result_item += get_item_5(sheet_now, sheet_pre, key)
    result_item += get_item_6(sheet_now, sheet_pre, key)
    result_item += get_item_7(sheet_now, sheet_pre, key)
    result_item += get_item_8(sheet_now, sheet_pre, key)
    result_item += get_item_9(sheet_now, sheet_pre, key)
    result_item += get_item_10(sheet_now, sheet_pre, key)
    result_item += get_item_11(sheet_now, sheet_pre, key)
    result_item += get_item_12(sheet_now, sheet_pre, key)
    result_item += get_item_13(sheet_now, sheet_pre, key)
    result_item += get_item_14(sheet_now, sheet_pre, key)
    result_item += get_item_15(sheet_now, sheet_pre, key)
    result_item += get_item_16(sheet_now, sheet_pre, key)
    result_item += get_item_17(sheet_now, sheet_pre, key)

    return result_item








if __name__ == "__main__":
    # 输入文件路径（支出明细表）
    input_file_path_now = "D:/RPATestDocumet/支出明细表-原表.xlsx"
    input_file_path_pre = "D:/RPATestDocumet/月报明细报表（202211CNY元）.xlsx"

    #人员表路径
    input_file_path_person_now = "D:/RPATestDocumet/person/2023年劳产率每月（11月）（人力原表）.xlsx"
    input_file_path_person_pre = "D:/RPATestDocumet/person/2022年劳产率每月（11月）.xlsx"

    # 打开Excel文件
    wb_now = openpyxl.load_workbook(input_file_path_now, data_only=True)
    wb_pre = openpyxl.load_workbook(input_file_path_pre, data_only=True)

    person_wb_now = openpyxl.load_workbook(input_file_path_person_now, data_only=True)
    person_wb_pre = openpyxl.load_workbook(input_file_path_person_pre, data_only=True)


    # 选择第一个工作表
    sheet_now = get_sheet(wb_now, "支出明细表")
    sheet_pre = get_sheet(wb_pre, "支出明细表")

    # 选择最后一个工作表
    sheet_person_now = person_wb_now.worksheets[-1]
    sheet_person_pre = person_wb_pre.worksheets[-1]


    # 输出文件路径（新的Excel表格）
    output_file_path = "D:/RPATestDocumet/file.xlsx"
    result = []
    for key, value in bank_name_map.items():
        result_item = []
        result_item.append(value)
        result_item += get_item_1(sheet_now, sheet_pre, key)
        result_item += get_item_2(sheet_now, sheet_pre, key)
        result_item += get_item_3(sheet_person_now, sheet_person_pre, key)
        result_item += get_item_4(sheet_person_now, sheet_person_pre, key)
        result_item += get_item_5(sheet_now, sheet_pre, key)
        result_item += get_item_6(sheet_now, sheet_pre, key)
        result_item += get_item_7(sheet_now, sheet_pre, key)
        result_item += get_item_8(sheet_now, sheet_pre, key)
        result_item += get_item_9(sheet_now, sheet_pre, key)
        result_item += get_item_10(sheet_now, sheet_pre, key)
        result_item += get_item_11(sheet_now, sheet_pre, key)
        result_item += get_item_12(sheet_now, sheet_pre, key)
        result_item += get_item_13(sheet_now, sheet_pre, key)
        result_item += get_item_14(sheet_now, sheet_pre, key)
        result_item += get_item_15(sheet_now, sheet_pre, key)
        result_item += get_item_16(sheet_now, sheet_pre, key)
        result_item += get_item_17(sheet_now, sheet_pre, key)
        total_now = 0.0
        total_pre = 0.0
        other_now = 0.0
        other_pre = 0.0

        result.append(result_item)
        print(result_item)
    list2 = count_total_2(sheet_now, sheet_pre, sheet_person_now, sheet_person_pre)
    list3 = count_total_3(sheet_now, sheet_pre, sheet_person_now, sheet_person_pre)
    list1 = count_total_1(list2, list3)
    print(list1)
    print(list2)
    print(list3)


    # 关闭Excel文件
    wb_now.close()
    wb_pre.close()
    person_wb_now.close()
    person_wb_pre.close()
