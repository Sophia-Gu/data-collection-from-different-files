import os
import win32com.client as win32


def list_dir(path, neededFileType='.xlsx', list_name=[], list_dir_name=[]):
    """
    :param path: 输入需要处理的文件夹（根目录）
    :param neededFileType: 需要寻找的文件类型(默认寻找.xlsx类型)
    :return: 返回目标文件所在的文件夹（不包含所寻找的目标文件本身）
    """
    # filetype_str = neededFileType
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        ft_n = file_path.rfind('.')
        filetype = file_path[ft_n:]
        if filetype == neededFileType:
            dir_file_name = os.path.dirname(file_path)
            list_dir_name.append(dir_file_name)
            break
        elif os.path.isdir(file_path):
            list_dir(file_path, neededFileType, list_name, list_dir_name )
        else:
            list_name.append(file_path)

    return list_dir_name


def xls_to_xlsx(filename):
    '''
    将xls文件转换为xlsx
    :param filename: xls文件路径
    :return: None
    注: xls  FileFormat = 56   xlsx FileFormat = 51
    '''
    xlApp = win32.gencache.EnsureDispatch('Excel.Application')
    xlBook = xlApp.Workbooks.Open(filename)
    xlBook.SaveAs(filename + 'x', FileFormat=51)
    xlBook.Close()
    xlApp.Application.Quit()
    os.remove(filename)


def file_path_rename(path, str='yF='):
    """
    功能：将文件目录转换为带下划线的文件名，方便之后的表格存储
    :param path: 待转换的文件路径
    :param str: 从路径中出现str开始之后的修改文件名
    :return: 返回以str开头，分隔符\被替代为_的字符串
    """
    path_rename = path.replace('\\', '_')
    file_star_num = path_rename.index(str)
    file_rename = path_rename[file_star_num:]
    return file_rename


def file_name(file_dir):
    """
    :param file_dir:文件路路径
    :return: 所有子文件的文件名(非文件夹)
    """
    for root, dirs, files in os.walk(file_dir):
        return files


def change_suffix(mainpath):
    """
    :param _mainpath: 文件所在路径
    :param suffix: 目标文件类型
    :return:
    """
    files = os.listdir(mainpath)
    for filename in files:
        full_path = os.path.join(mainpath, filename)
        dot_position = full_path.rfind(".")
        os.rename(full_path, full_path[:dot_position] + '.txt')
        # 查找最后一个 dot 出现的位置，作为分割文件名的依据，防止文件名中就含有dot的情况


# 判断字符串中是否存在关键词
def check(string, sub_str):
    if (string.find(sub_str) == -1):
        print("不存在！")
    else:
        print("存在！")