
#  鞋子集合为全局变量
# 10双一盒
rest_shoes_list = []
# 存放多余的型号
rest_model_sex = set()
# box_total = 0
rest_box_total = 0
# 将10只鞋装一箱子
'''
last 最后一次接收数量的下标，用于处理最后几双鞋，不满10双的  True为最后一箱
total_index 合计的行数
num 控制下一个表单
last 为是否最后一条数据
accept_index 数据录入的位置
j 是打印两份

装10箱
'''
def pac_boxes_five(j, accept_index, total_index, sheet_read, box_sheet, num, last, content_style):

    global rest_shoes_list

    global rest_model_sex

    global rest_box_total


    # 存放鞋的信息
    shoes_list = []

    # 表示有剩余的鞋
    box_total = 0
    if len(rest_shoes_list) > 0:
        model_sex = set(rest_model_sex)
        shoes_list.extend(rest_shoes_list)
        box_total = rest_box_total

        if j == 1:
            rest_shoes_list.clear()
            rest_model_sex.clear()
            rest_box_total = 0
    else:
        # 专门存放型号
        model_sex = set()

    index = 0
    for i in range((accept_index + 1), total_index, 1):

        # 存放1盒中的型号
        model_sex.add(sheet_read.cell(i, 3).value)

        shoes_dict = {}

        shoes_dict['model'] = str(sheet_read.cell(i, 3).value)
        shoes_dict['num'] = int(sheet_read.cell(i, 5).value)
        shoes_list.append(shoes_dict)

        box_total += int(sheet_read.cell(i, 5).value)
        # 如果装够10双鞋，结束本次循环，重新开始装
        if last == False:
            # 如果鞋大于10双
            if box_total >= 5:
                if box_total == 5:
                    box_total = 0

                else:
                    box_total -= 5
                    shoes_list[len(shoes_list) - 1]['num'] = box_total

                modelNum(model_sex, shoes_list, box_sheet, num, content_style)

                index = i

                # 如果还有数据则
                if box_total > 0 and j == 1:
                    rest_model_sex.clear()
                    rest_shoes_list.clear()
                    rest_box_total = 0

                    rest_box_total = box_total
                    shoes_dict['model'] = str(sheet_read.cell(i, 3).value)
                    shoes_dict['num'] = int(box_total)
                    rest_shoes_list.append(shoes_dict)

                    # 存储剩余的型号
                    rest_model_sex.add(sheet_read.cell(i, 3).value)
                break

    if last == True:
        modelNum(model_sex, shoes_list, box_sheet, num, content_style)
        if j == 1:
            rest_box_total = 0
            rest_model_sex.clear()
            rest_shoes_list.clear()

    # 加空白格的边框
    if len(model_sex) <= 5:
        model_length = num + len(model_sex)
        for i in range(model_length, (num + 10), 1):
            if i < (num + 5):
                box_sheet.write(i, 1, '', content_style)
                box_sheet.write(i, 2, '', content_style)
            else:
                box_sheet.write(i - 5, 5, '', content_style)
                box_sheet.write(i - 5, 6, '', content_style)
    else:
        model_length = len(model_sex) - 5 + num
        for b in range(model_length, (num + 5), 1):
            box_sheet.write(b, 5, '', content_style)
            box_sheet.write(b, 6, '', content_style)
    if j == 1:
        return index
    else:
        return accept_index

'''
    将型号数量写入列表中
'''

def modelNum(model_sex, shoes_list, box_sheet, num, content_style):
    # 将10只鞋装箱盒中
    # 计算循环的次数
    i = 1
    for model in model_sex:
        # 按型号取得数量
        model_num = 0
        for shoes in shoes_list:
            if model == shoes['model']:
                model_num += int(shoes['num'])

        # 分两列显示 如果型号超过5个，则在第二列显示
        model_col = 0
        num_col = 0
        if len(model_sex) <= 5:
            model_col = 1
            num_col = 2
        elif i <= 5:
            model_col = 1
            num_col = 2
        else:
            model_col = 5
            num_col = 6
            num -= 5
        box_sheet.write(num, model_col, model, content_style)
        box_sheet.write(num, num_col, model_num, content_style)
        num += 1
        i += 1