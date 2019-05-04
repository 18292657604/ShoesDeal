import os

# 获取目录中所有的文件列表
def getAllFile(mkdir_path):
    for root, dirs, files in os.walk(mkdir_path):
        print(files)
        print('===========')
        return files

# 顺序访问文件夹
def getFilelist(mkdir_path):
    fileList = os.listdir(mkdir_path)
    # 汉子排序问题
    fileList.sort()

    fileList = [i.encode('GBK') for i in fileList]

    fileList.sort()

    fileList = [i.decode('GBK') for i in fileList]

    return fileList





if __name__ == '__main__':
    getAllFile('C:/Users/LSY/Desktop/消防学院春秋作训鞋/')