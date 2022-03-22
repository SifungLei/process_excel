#coding=utf-8

# 分析EXCEL表，赠送的数据
import sys
import openpyxl
import os
import user_input
import time
import merge_sheet

# 中文字符与数字
char_num = {
"一" : 1,
"二" : 2,
"三" : 3,
"四" : 4,
"五" : 5,
"六" : 6,
"七" : 7,
"八" : 8,
"九" : 9,
"十" : 10,
}

# 列，求和
def cal_column_sum(sheet, col_str, from_row):
    column_valid_cell = merge_sheet.cal_column_cell(sheet, col_str, from_row)
    ret = 0.0
    for cell in column_valid_cell:
        ret += float(cell.value)
    print("calling cal_column_sum, sheet = %r, 列号 = %s, 第%d行到末行, 求和 = %f" % (sheet, col_str, from_row, ret))
    return ret

# 判断文件是否存在，不存在就退出程序
def check_file_exist(file_name):
    if os.path.isfile(file_name) == False:
        print("%s不存在" % (file_name))
        sys.exit(0)

# 加载文件
def load_original_file():
    check_file_exist(user_input.file_name_original)
    print("原文件名 = %s, 工作表名 = %s" % (user_input.file_name_original, user_input.sheet_name_original))

    print("准备加载%s" % (user_input.file_name_original))
    wb = openpyxl.load_workbook(user_input.file_name_original, data_only=True)
    if wb == None:
        print("加载%s失败" % (user_input.file_name_original))
        sys.exit(0)
    print("加载%s完毕" % (user_input.file_name_original))

    all_sheet_name = wb.get_sheet_names()
    if user_input.sheet_name_original not in all_sheet_name:
        print("%s不存在" % (user_input.sheet_name_original))
        sys.exit(0)

    ws = wb[user_input.sheet_name_original]
    return wb, ws

# 处理赠送的
def process_present(wb, ws, all_express_no_cell, all_business_no_cell, all_price_cell, all_buy_cnt_cell, all_present_money_cell):
    # KEY是快递单号，V是数据
    express_no_data = {}
    index = -1
    for cell in all_express_no_cell:
        index += 1
        if index < user_input.valid_row_no - 1:
            continue

        row = cell.row # 行号，数
        # print("calling process_present, 正在遍历数据, row = %d" % row)

        # 快递单号
        express_no = cell.value
        if express_no == None:
            ouput_none_cell(row, "快递单号")
            continue
        express_no = str(express_no)

        if row <= 10:
            # print("calling process_present, row = %d, express_no = %s" % (row, express_no))
            None

        # 商家编码
        business_no = all_business_no_cell[index].value
        if business_no == None:
            ouput_none_cell(row, "商家编码")
            continue
        business_no = str(business_no)
        # print("calling process_present, row = %d, business_no = %s" % (row, business_no))

        # 价格
        price = all_price_cell[index].value
        if price == None:
            ouput_none_cell(row, "价格")
            continue
        price = float(price)

        # 购买数量
        buy_cnt = all_buy_cnt_cell[index].value
        if buy_cnt == None:
            ouput_none_cell(row, "购买数量")
            continue
        buy_cnt = int(buy_cnt)

        if business_no.count("送") == 0 and business_no.count("满") == 0 and business_no.count("组送") == 0:
            # print("calling process_present, row = %d, business_no = %s, 不满足赠送费的条件" % (row, business_no))
            continue

        print("row = %d, 快递单号%s, 商家编码%s, 价格%f, 购买数量%d" % (row, express_no, business_no, price, buy_cnt))

        if express_no not in express_no_data:
            express_no_data[express_no] = {}

        if business_no.count("组") > 0:
            group_cnt = cal_group_cnt_by_type(business_no)
            single_cnt = cal_single_cnt_by_type(business_no)
            full_info = "类型：%d组送%d" % (group_cnt, single_cnt)

            if full_info not in express_no_data[express_no]:
                express_no_data[express_no][full_info] = {"price" : price, "buy_cnt" : 0, "present_price" : float(price / 5.0)}
            express_no_data[express_no][full_info]["buy_cnt"] += buy_cnt
        else:
            full_cnt = cal_full_cnt_by_type(business_no)
            single_cnt = cal_single_cnt_by_type(business_no)
            full_info = "类型：满%d送%d" % (full_cnt, single_cnt)

            """
            if express_no == "No:3305671358105":
                print("full_info = %s" % full_info)
            """

            if full_info not in express_no_data[express_no]:
                express_no_data[express_no][full_info] = {"price" : price, "buy_cnt" : 0, "present_price" : price}

            # 特别的赠送价格
            if business_no.count("袜"):
                express_no_data[express_no][full_info]["present_price"] = user_input.present_price_socks

            express_no_data[express_no][full_info]["buy_cnt"] += buy_cnt

        # if index >= 1000:
        #     break

    print("\n\nexpress_no_data = %r\n\n" % express_no_data)
    if len(express_no_data) == 0:
        print("calling process_present, express_no_data 为空")
        return

    # begin
    # 计算所有赠送的钱
    # 所有赠送的钱
    all_present_money = {}

    for express_no, express_no_value in express_no_data.items():
        for full_info, business_no_value in express_no_value.items():
            full_cnt = 0
            if full_info.count("组送") > 0:
                group_cnt = cal_group_cnt_by_type(full_info)
                if group_cnt == 0:
                    continue
                if business_no_value["buy_cnt"] < group_cnt:
                    continue
                full_cnt = group_cnt
            else:
                full_cnt = cal_full_cnt_by_type(full_info)
                if full_cnt == 0:
                    continue
                if business_no_value["buy_cnt"] < full_cnt:
                    continue

                """
                if express_no == "No:3305671358105":
                    print("符合要求，express_no = %s, full_info = %s, full_cnt = %d, buy_cnt = %d" % (express_no, full_info, full_cnt, business_no_value["buy_cnt"]))
                """
            if express_no not in all_present_money:
                all_present_money[express_no] = {}
            if full_info not in all_present_money[express_no]:
                all_present_money[express_no][full_info] = business_no_value

            # 赠送数量
            present_cnt = (int(business_no_value["buy_cnt"] / full_cnt)) * (cal_single_cnt_by_type(full_info))
            # 赠送钱
            present_money = float(present_cnt) * float(business_no_value["present_price"])

            all_present_money[express_no][full_info]["present_cnt"] = present_cnt
            all_present_money[express_no][full_info]["present_money"] = present_money

    print("\n\nall_present_money = %r\n\n" % all_present_money)
    # end

    # begin
    # 输出数据到赠送费这一列
    index = -1
    has_output_express = []
    for cell in all_express_no_cell:
        index += 1
        if index < user_input.valid_row_no - 1:
            continue

        row = cell.row # 行号，数

        # 快递单号
        express_no = cell.value

        if express_no not in all_present_money:
            all_present_money_cell[index].value = 0
            continue

        if has_output_express.count(express_no) > 0:
            # 已经输出过
            print("row = %d, express_no = %s, 已经输出过" % (row, express_no))
            all_present_money_cell[index].value = 0
            continue

        has_output_express.append(express_no)
        all_full_info = all_present_money[express_no].keys()
        if len(all_full_info) >= 2:
            print("express_no = %s, all_full_info = %r" % (express_no, all_full_info))
        total_money = 0 # 该快递单号总的赠送钱
        for full_info in all_full_info:
            total_money += all_present_money[express_no][full_info]["present_money"]

        print("输出赠送费，row = %d, express_no = %s, total_money = %f" % (row, express_no, total_money))
        all_present_money_cell[index].value = total_money
    # end

    cal_column_sum(ws, "j", 2)

    # 保存
    wb.save(user_input.file_name_original)

# 根据类型计算X组
def cal_group_cnt_by_type(type_str):
    type_str_original = type_str
    # print("处理前type_str = %s" % (type_str_original))

    for char, num in char_num.items():
        type_str = type_str.replace("%s组" % (char), "%d组" % (num))

    # print("处理后type_str = %s" % (type_str))

    temp = []
    for i in range(1, 11):
        if type_str.count("%d组" % (i)) == 1:
            temp.append(i)

    if temp == []:
        print("error, type_str = %s, temp = %r" % (type_str_original, temp))

    # 取最大的
    ret = sorted(temp, reverse=True)[0]
    return ret

# 根据类型计算出满多少
def cal_full_cnt_by_type(type_str):
    type_str_original = type_str
    # print("处理前type_str = %s" % (type_str_original))

    for char, num in char_num.items():
        type_str = type_str.replace("满%s" % (char), "满%d" % (num))

    # print("处理后type_str = %s" % (type_str))

    temp = []
    for i in range(1, 11):
        if type_str.count("满%d" % (i)) == 1:
            temp.append(i)

    if temp == []:
        print("error, type_str = %s, temp = %r" % (type_str_original, temp))

    # 取最大的
    ret = sorted(temp, reverse=True)[0]
    return ret

# 判断列号是否存在
def check_column_exist(column_index, column_cnt):
    if column_index >= column_cnt:
        print("列号%s不存在" % (openpyxl.utils.get_column_letter(column_index + 1)))
        sys.exit(0)

# 通过列号，计算该列所有数据的单元格
def cal_all_cell_by_column_str(column_str, column_cnt):
    col_index = openpyxl.utils.column_index_from_string(column_str) - 1
    check_column_exist(col_index, column_cnt)
    all_cell = ws.columns[col_index]
    return all_cell

# 输出空单元格的坐标
def ouput_none_cell(row, name):
    print("行号数 = %d, %s为空单元格" % (row, name))

# 根据类型计算出赠送单项数量
def cal_single_cnt_by_type(type_str):
    for char, num in char_num.items():
        type_str = type_str.replace("送%s" % (char), "送%d" % (num))

    temp = []
    for i in range(1, 11):
        if type_str.count("送%d" % (i)) == 1:
            temp.append(i)

    if len(temp) == 0:
        # 不注明的，默认送1
        return 1
    # 取最大的
    ret = sorted(temp, reverse=True)[0]
    return ret

# 使用说明
def use_des():
    des = "使用python3、openpyxl来处理excel数据。\n用法：python3 %s\n" % (sys.argv[0])
    print(des)

# 处理结束后
def after_process():
    print("\n处理完毕, 按下回车键可以关闭窗口\n")
    a = input()
    # time.sleep(2)

if __name__ == "__main__":
    use_des()
    # 工作簿、工作表
    wb, ws = load_original_file()

    # 行数量
    row_cnt = len(ws.rows)
    print("行数量 = %d" % (row_cnt))

    # 列数量
    column_cnt = len(ws.columns)
    print("列数量 = %d" % (column_cnt))

    # 获取所有快递单号的单元格
    all_express_no_cell = cal_all_cell_by_column_str(user_input.express_no_column_str, column_cnt)

    # 获取所有商家编码的单元格
    all_business_no_cell = cal_all_cell_by_column_str(user_input.business_no_column_str, column_cnt)

    # 获取所有价格的单元格
    all_price_cell = cal_all_cell_by_column_str(user_input.price_column_str, column_cnt)

    # 获取所有购买数量的单元格
    all_buy_cnt_cell = cal_all_cell_by_column_str(user_input.buy_cnt_column_str, column_cnt)

    # 获取所有赠送费的单元格
    all_present_money_cell = cal_all_cell_by_column_str(user_input.present_money_column_str, column_cnt)

    print("开始处理数据")
    process_present(wb, ws, all_express_no_cell, all_business_no_cell, all_price_cell, all_buy_cnt_cell, all_present_money_cell)
    after_process()

