import xlwt
import xlrd
from xlutils.copy import copy

'''
设置单元格样式
0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray
'''
def set_style(fontColor, align, name,height,bold=False,background=1, underLine=0, border=True):
    style = xlwt.XFStyle()  # 初始化样式

    # 设置字体
    font = xlwt.Font()  # 为样式创建字体s
    font.name = name # 'Times New Roman'
    font.bold = bold
    font.underline = underLine
    font.colour_index = fontColor
    font.height = height

    # 设置边框
    borders = xlwt.Borders()
    if border==True:
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN

    # 设置居中
    alignment = xlwt.Alignment()
    if align==1:
        alignment.horz = xlwt.Alignment.HORZ_CENTER
    else:
        alignment.horz = xlwt.Alignment.HORZ_LEFT
    alignment.vert = xlwt.Alignment.VERT_CENTER

    # 等于1为自动换行
    alignment.wrap = 1

    # 设置背景颜色
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = background # 背景颜色

    style.pattern=pattern
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


'''
    设置行高
'''
def set_row_height(row_obj, height):
    row_obj.height_mismatch = True
    row_obj.height = 20 * height


'''
    导入到Excel
'''
def excel(outPath):
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

        wb_sheet.col(1).width= 256 * 8
        wb_sheet.col(2).width= 256 * 30
        wb_sheet.col(6).width= 256 * 20
        # 要把鞋装多少箱
        box_num = (int(total_content/10))+1

        wb_sheet.write(0, 6, '第  箱，共 '+str(box_num)+' 箱', set_style(0, 0, '宋体', 240, False, 1, 0, False))

        if int(total_content) % 10 == 0:
            box_num = (int(total_content / 10))
        else:
            box_num = (int(total_content / 10)) + 1

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

        #页眉页脚设置 为空
        box_sheet.header_str = bytes('', encoding='utf-8')
        box_sheet.footer_str = bytes('', encoding='utf-8')

        # 合并单元格制作标题(table为标题)
        num = 0
        # 接收数量下标
        accept_index = 6
        page = 0
        for i in range(box_num):
            #设置行高
            set_row_height(box_sheet.row(num), 40)
            set_row_height(box_sheet.row(num+1), 27)
            set_row_height(box_sheet.row(num+2), 27)
            set_row_height(box_sheet.row(num+3), 27)
            set_row_height(box_sheet.row(num+4), 27)
            set_row_height(box_sheet.row(num+5), 27)
            set_row_height(box_sheet.row(num+6), 27)
            set_row_height(box_sheet.row(num+7), 27)
            set_row_height(box_sheet.row(num+8), 27)
            set_row_height(box_sheet.row(num+9), 27)
            set_row_height(box_sheet.row(num+10), 27)
            set_row_height(box_sheet.row(num+11), 40)
            set_row_height(box_sheet.row(num+12), 27)


            box_sheet.write_merge(num, num, 0, 6, '消防救援制式服装和标志服饰装箱单', set_style(0, 1, '黑体', 440, True, 1, 0, False))
            box_sheet.write_merge(num + 1, num + 1, 0, 5, '单位：消防高等专科学校 教学保障大队 教学保障大队干部', other_style)
            box_sheet.write(num + 1, 6, '共 '+ str(box_num) +' 箱', content_style)

            box_sheet.write_merge(num + 2, num + 2, 0, 5, '品名：19消防春秋作训鞋（双）', other_style)
            box_sheet.write(num + 2, 6, '第 '+ str(i + 1) +' 箱', content_style)

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
            # 返回下一个接受数量的开始
            #如果是最后一箱
            if (box_num-1)==i:
                accept_index = pac_boxes(accept_index, total_index, sheet_read, box_sheet, num + 4, True, content_style)
                # 最后一箱的数量
                last_num = int(total_content) % 10
                box_sheet.write_merge(num + 9, num + 9, 0, 6, '本箱内合计数量:' + str(last_num) + '双', content_style)
            else:
                accept_index = pac_boxes(accept_index, total_index, sheet_read, box_sheet, num + 4, False, content_style)
                box_sheet.write_merge(num + 9, num + 9, 0, 6, '本箱内合计数量:10双', content_style)
            # ===============================================

            box_sheet.write_merge(num + 10, num + 10, 0, 6, '联系人： 王宁，朱忠奎    联系电话：13769160919，18208805630', other_style)
            box_sheet.write_merge(num + 11, num + 11, 0, 6, '地址： 云南省 昆明市 经开区 云南省昆明市官渡区阿拉乡小石坝公安消防部队高等专科学校教学保障大队收', other_style)
            box_sheet.write_merge(num + 12, num + 12, 0, 6, '生产厂家：际华三五一三实业有限公司', set_style(0, 0, '黑体', 410, False, 1, 0, False))

            num += 14

            # 分页设置
            box_sheet.horz_page_breaks = [(page, page, (page+28))]
            page+=28

        # 保存excel
        workbook.save(outPath)
    except Exception as e:
        print('sheet重复：%s' %(e))
'''
last 最后一次接收数量的下标，用于处理最后几双鞋，不满10双的  True为最后一箱
'''
# 将10只鞋装一箱子
def pac_boxes(j, total_index, sheet_read, box_sheet, num, last, content_style):
    shoes_list = []
    # 10双一盒
    # 专门存放型号
    model_sex = set()
    box_total = 0;
    index = 0
    for i in range((j+1), total_index, 1):
        #存放1盒中的型号
        model_sex.add(sheet_read.cell(i, 3).value)

        shoes_dict = {}

        shoes_dict['model'] = str(sheet_read.cell(i, 3).value)
        shoes_dict['num'] = int(sheet_read.cell(i, 5).value)
        shoes_list.append(shoes_dict)

        box_total += int(sheet_read.cell(i, 5).value)
        # 如果装够10双鞋，结束本次循环，重新开始装
        if last==False:
            if box_total == 10:
                box_total = 0
                modelNum(model_sex, shoes_list, box_sheet, num, content_style)
                index = i
                break
    if last==True:
        modelNum(model_sex, shoes_list, box_sheet, num, content_style)

    # 加空白格的边框
    if len(model_sex)<5:
        model_length = num + len(model_sex)
        for i in range(model_length, (num + 10), 1):
            if i<(num+5):
                box_sheet.write(i, 1, '', content_style)
                box_sheet.write(i, 2, '', content_style)
            else:
                box_sheet.write(i-5, 5, '', content_style)
                box_sheet.write(i-5, 6, '', content_style)
    else:
        model_length = len(model_sex)-5
        for j in range(model_length, (num + 5), 1):
            box_sheet.write(j, 5, '', content_style)
            box_sheet.write(j, 6, '', content_style)
    return index
'''
    将型号数量写入列表中
'''
def modelNum(model_sex, shoes_list, box_sheet, num, content_style):
    # 将10只鞋装箱盒中
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
        else:
            model_col = 5
            num_col = 6
        box_sheet.write(num, model_col, model, content_style)
        box_sheet.write(num, num_col, model_num, content_style)
        num += 1





