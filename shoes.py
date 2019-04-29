from format.ten.ExcelTenDeal import excel_ten
from format.twenty.ExcelTwentyDeal import excel_twenty
from util import getAllFile


# 开始转换
def startConvert(mkdir_path, type):
    type = int(type)
    if type == 10 or type == 20:
        fileNameLest = getAllFile(mkdir_path)
        for t in range(len(fileNameLest)):
            fileName = mkdir_path + str(fileNameLest[t])

            if type == 10:
                # 10装箱
                result = excel_ten(fileName)
            else:
                # 20装箱
                result = excel_twenty(fileName)

            if result == True:
                print(fileName + "====处理成功")
            else:
                print(fileName + "====处理失败")
        print('所有文件处理成功')
    else:
        print("不符合规定请重新输入：")




if __name__ == "__main__":
    mkdir_path = input("请输入目录路径:")
    type = input("请输入规则的类型(注10、20):")

    type = int(type)

    if type == 10 or type == 20:
        fileNameLest = getAllFile(mkdir_path)
        for t in range(len(fileNameLest)):
            fileName = mkdir_path + str(fileNameLest[t])

            if type == 10:
                # 10装箱
                result = excel_ten(fileName)
            else:
                # 20装箱
                result = excel_twenty(fileName)

            if result == True:
                print(fileName + "====处理成功")
            else:
                print(fileName + "====处理失败")
        print('所有文件处理成功')
    else:
        print("不符合规定请重新输入：")







