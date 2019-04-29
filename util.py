import os

# 获取目录中所有的文件列表
def getAllFile(mkdir_path):
    for root, dirs, files in os.walk(mkdir_path):
        return files