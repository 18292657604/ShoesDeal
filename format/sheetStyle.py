import xlwt

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