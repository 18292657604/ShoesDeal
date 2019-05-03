from util import getAllFile
from format.post.postoffice import postoffice

# 打包成邮局的格式
def post(mkdir_path, out_path, type):
    fileNameLest = getAllFile(mkdir_path)
    for t in range(len(fileNameLest)):
        fileName = mkdir_path + '/' + str(fileNameLest[t])
        postoffice(t, fileName, out_path, type)


if __name__ == "__main__":

    mkdir_path = input("请输入要遍历的文件夹:")
    out_path = input("请输入目标Excel:")

    type = input("请输入规格（10/20）")

    post(mkdir_path, out_path, type)