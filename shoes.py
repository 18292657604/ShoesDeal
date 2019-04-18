from ExcelDeal import excel
import os


# 获取目录中所有的文件列表
def getAllFile(mkdir_path):
    for root, dirs, files in os.walk(mkdir_path):
        return files

if __name__ == "__main__":
    mkdir_path = input("请输入目录路径:")
    fileNameLest = getAllFile(mkdir_path)

    print('======================')
    for t in range(len(fileNameLest)):
        fileName = mkdir_path + str(fileNameLest[t])
        excel(t, fileName)
        print(fileName + "====处理成功")
        print('======================')
    print('所有文件处理成功')







