import dir_output
import sum0_of_excel
"""
功能说明：
1.对于经过预处理的excel近行数据的查询与整合，形成一级sum文件
2.存储名为文件目录的变形，以第一个特征文件字符为起点
3.可重复使用（会将之前生成的sum文件删除）
4. excel以.xls格式打开或存储，便于数据的读取，并防止因为兼容问题出错
"""

if __name__ == "__main__":

    print('******************************************说明******************************************')
    print('第一次输入文件路径为目标整理数据所在文件夹，可以是最外层目录，也可以是靠近目标文件的目录，按需输入')
    print('汇总后的表格会以 文件路径_SUM.xlsx 命名，故第二次输入的字符为路径上的需要保留的标志性起始字符，按需输入')
    print(r'例 数据文件所在目录为C:\Users\yujun\Desktop\DR-PSA\yF=0.1\PL-A……')
    print(r'若DR-PSA对于汇总文件无意义则忽略，若yF=有意义则输入它，无意义则可以选择PL-A')
    print('注意：输入的起始字符必须是文件路径中存在的连续字符!!!!! ')
    print('*****************************************************************************************\n')
    path = input('大哥快点输入文件路径：')
    # 当文件以yF=开头时，不必用该语句
    str_start = input('输入起始字符')
    print('loading~~~~~')
    print('pu->pu->>pu->>>pu->>>>~~~~~~')
    # 找到excel所在的文件路径
    xls_paths = dir_output.list_dir(path, '.xlsx')
    # 对所有的excel进行处理
    for i in xls_paths:
        sum0_of_excel.sum_of_values(i, str_start)

    print('数据处理已完成，细心的你可以检查一下哦！')
    input('输入任意键结束！')
