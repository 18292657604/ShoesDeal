import xlwt
import xlrd
from xlutils.copy import copy
from format.ten.TenPack import pac_boxes_ten
from format.sheetStyle import *


'''
    导入到Excel
'''
def excel_ten(outPath):
    try:
        # 如果excel中有文件，则在后面添加
        rb = xlrd.open_workbook(outPath, formatting_info=True)
        if rb != '':
            workbook = copy(rb)
        else:
            # 创建excel
            workbook = xlwt.Workbook()
        # 读取第一个工作表中（索引顺序获取）
        sheet_read = rb.sheet_by_index(0)

        # 获取接受数量

        #修改第一个sheet中的内容
        wb_sheet = workbook.get_sheet(0)

        wb_sheet.write(5, 1, '生产厂家：际华三五一三实业有限公司', set_style(0, 0, '宋体', 220, False, 1, 0, False))

        total_index = 0
        for rownum in range(sheet_read.nrows):
            if '合计' in sheet_read.row_values(rownum):
                total_index = rownum
                break

        # 获取合计的内容
        total_content = sheet_read.cell(total_index, 5).value

        items_name = sheet_read.cell(5, 5).value

        # 收货单位
        recive_unit = str(sheet_read.cell(3, 1).value).split('：')[1]

        # 收货地址
        recive_address = str(sheet_read.cell(2, 1).value.split('：')[1]).strip()

        # 收货人
        revice_person = str(sheet_read.cell(4, 1).value.split('：')[1]).strip()

        # 联系电话
        tel = str(sheet_read.cell(4, 3).value.split('：')[1]).strip()

        wb_sheet.col(1).width = 256 * 8
        wb_sheet.col(2).width = 256 * 30
        wb_sheet.col(6).width = 256 * 20
        # 要把鞋装多少箱
        if int(total_content) % 10 == 0:
            box_num = (int(total_content / 10))
        else:
            box_num = (int(total_content / 10)) + 1
        wb_sheet.write(0, 6, '第  箱，共 '+str(box_num)+' 箱', set_style(0, 0, '宋体', 240, False, 1, 0, False))

        # 添加sheet 制作封装箱单子
        box_sheet = workbook.add_sheet("装箱单")

        box_sheet.col(0).width = 256 * 15
        box_sheet.col(1).width = 256 * 15
        box_sheet.col(2).width = 256 * 15
        box_sheet.col(3).width = 256 * 3
        box_sheet.col(4).width = 256 * 15
        box_sheet.col(5).width = 256 * 15
        box_sheet.col(6).width = 256 * 15

        content_style = set_style(0, 1, '黑体', 320, False, 1, 0, True)
        other_style = set_style(0, 0, '黑体', 300, False, 1, 0, True)

        # 页眉页脚设置 为空
        box_sheet.header_str = bytes('', encoding='utf-8')
        box_sheet.footer_str = bytes('', encoding='utf-8')

        # 合并单元格制作标题(table为标题)
        num = 0
        # 接收数量下标
        accept_index = 6
        page = 0
        for i in range(box_num):
            # 打印两份
            for j in range(2):
                #设置行高
                set_row_height(box_sheet.row(num), 36)
                set_row_height(box_sheet.row(num+1), 36)
                set_row_height(box_sheet.row(num+2), 26)
                set_row_height(box_sheet.row(num+3), 26)
                set_row_height(box_sheet.row(num+4), 26)
                set_row_height(box_sheet.row(num+5), 26)
                set_row_height(box_sheet.row(num+6), 26)
                set_row_height(box_sheet.row(num+7), 26)
                set_row_height(box_sheet.row(num+8), 26)
                set_row_height(box_sheet.row(num+9), 26)
                set_row_height(box_sheet.row(num+10), 36)
                set_row_height(box_sheet.row(num+11), 36)
                set_row_height(box_sheet.row(num+12), 26)


                box_sheet.write_merge(num, num, 0, 6, '消防救援制式服装和标志服饰装箱单', set_style(0, 1, '黑体', 440, True, 1, 0, False))
                box_sheet.write_merge(num + 1, num + 1, 0, 5, '单位：' + recive_unit, other_style)
                box_sheet.write(num + 1, 6, '共 '+ str(box_num) + ' 箱', content_style)

                box_sheet.write_merge(num + 2, num + 2, 0, 5, '品名：' + items_name + '（双）', other_style)
                box_sheet.write(num + 2, 6, '第 '+ str(i + 1) + ' 箱', content_style)

                box_sheet.write(num + 3, 0, '序号', content_style)
                box_sheet.write(num + 3, 1, '号型', content_style)
                box_sheet.write(num + 3, 2, '数量', content_style)

                box_sheet.write(num + 3, 4, '序号', content_style)
                box_sheet.write(num + 3, 5, '号型', content_style)
                box_sheet.write(num + 3, 6, '数量', content_style)

                box_sheet.write(num +4, 0, 1, content_style)
                box_sheet.write(num +5, 0, 2, content_style)
                box_sheet.write(num +6, 0, 3, content_style)
                box_sheet.write(num +7, 0, 4, content_style)
                box_sheet.write(num +8, 0, 5, content_style)

                box_sheet.write(num + 4, 4, 6, content_style)
                box_sheet.write(num + 5, 4, 7, content_style)
                box_sheet.write(num + 6, 4, 8, content_style)
                box_sheet.write(num + 7, 4, 9, content_style)
                box_sheet.write(num + 8, 4, 10, content_style)

                # ===================================
                # 如果是最后一箱
                if (box_num - 1) == i:
                    # 返回下一个接受数量的开始
                    accept_index = pac_boxes_ten(j, accept_index, total_index, sheet_read, box_sheet, num + 4, True,
                                                 content_style)
                    # 最后一箱的数量
                    if (int(total_content) % 10) == 0:
                        last_num = 10
                    else:
                        last_num = int(total_content) % 10
                    box_sheet.write_merge(num + 9, num + 9, 0, 6, '本箱内合计数量:' + str(last_num) + '双', content_style)
                else:
                    accept_index = pac_boxes_ten(j, accept_index, total_index, sheet_read, box_sheet, num + 4, False,
                                                 content_style)
                    box_sheet.write_merge(num + 9, num + 9, 0, 6, '本箱内合计数量:10双', content_style)
                    # =======================================================

                box_sheet.write_merge(num + 10, num + 10, 0, 6, '联系人：'+ revice_person + '    联系电话：' + tel, other_style)
                box_sheet.write_merge(num + 11, num + 11, 0, 6, '地址：' + recive_address, other_style)
                box_sheet.write_merge(num + 12, num + 12, 0, 6, '生产厂家：际华三五一三实业有限公司', set_style(0, 0, '黑体', 410, False, 1, 0, False))

                num += 14

            # 分页设置
            box_sheet.horz_page_breaks = [(page, page, (page+28))]
            page+=28

        # 保存excel
        workbook.save(outPath)
        return True
    except Exception as e:
        print('处理失败：%s' %(e))
        return False








