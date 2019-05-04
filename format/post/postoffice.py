import xlrd
from xlutils.copy import copy
from format.sheetStyle import *


def postoffice(index, fromPath, toPath, type):

    try:
        # 如果excel中有文件，则在后面添加
        to_rb = xlrd.open_workbook(toPath, formatting_info=True)
        from_rb = xlrd.open_workbook(fromPath)
        if to_rb != '':
            workbook = copy(to_rb)
        else:
            # 创建excel
            workbook = xlwt.Workbook()
            workbook.add_sheet('收货信息')
        # 读取第一个工作表中（索引顺序获取）
        from_sheet = from_rb.sheet_by_index(0)

        # 修改第一个sheet中的内容
        wb_sheet = workbook.get_sheet(0)

        # 合计的行数
        total_row = 0
        for rownum in range(from_sheet.nrows):
            if '合计' in from_sheet.row_values(rownum):
                total_row = rownum
                break

        # 获取合计的内容
        total_val = from_sheet.cell(total_row, 5).value

        # 收货单位
        recive_unit = str(from_sheet.cell(3, 1).value).split('：')[1]

        # 收货地址
        recive_address = str(from_sheet.cell(2, 1).value.split('：')[1]).strip()

        # 收货人
        revice_person = str(from_sheet.cell(4, 1).value.split('：')[1]).strip()

        # 联系电话
        tel = str(from_sheet.cell(4, 3).value.split('：')[1]).strip()

        # 共多少箱
        box_num = 0
        # 20 箱装
        if type == 20:
            remain_num = int(total_val)
            # 20、10、个位数箱
            while True:
                if remain_num >= 20:
                    remain_num -= 20
                    box_num += 1
                elif remain_num >= 10:
                    remain_num -= 10
                    box_num += 1
                elif remain_num > 0:
                    remain_num = 0
                    box_num += 1
                elif remain_num == 0:
                    break
        elif type==10:
            # 10箱装
            if int(total_val) % 10 == 0:
                box_num = (int(total_val / 10))
            else:
                box_num = (int(total_val / 10)) + 1

        for i in range(0, box_num, 1):
            # 重写到地址
            wb_sheet.write(index, 0, index+1)
            wb_sheet.write(index, 1, revice_person)
            wb_sheet.write(index, 2, tel)
            wb_sheet.write(index, 3, recive_unit)
            wb_sheet.write(index, 4, recive_address)
            wb_sheet.write(index, 5, 1)
            index += 1

        # 保存excel
        workbook.save(toPath)
        return index
    except Exception as e:
        print('处理失败 %s' %(e))




