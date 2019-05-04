from util import getFilelist
from format.post.postoffice import postoffice


# 打包成邮局的格式
def post(mkdir_path, out_path, type):
    # 下标
    index = 0
    fileNameLest = getFilelist(mkdir_path)
    for t in range(len(fileNameLest)):
        fileName = mkdir_path + '/' + str(fileNameLest[t])
        index = postoffice(index ,fileName, out_path, type)

if __name__ == "__main__":

    mkdir_path = input("请输入要遍历的文件夹:")
    out_path = input("请输入目标Excel:")

    type = input("请输入规格（10/20）")

    post(mkdir_path, out_path, type)