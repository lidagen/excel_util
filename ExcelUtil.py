import xlrd
import xlwt


def read_excel1():
    path = "D:\\"
    # 生产文件名称、位置
    full_path = path + "perform1.sql"
    file = open(full_path, 'w')
    # 打开文件
    workbook = xlrd.open_workbook(r'D:\\代客下单订单导入20200601184020.xls')
    # sheet 0
    sheet = workbook.sheet_by_index(0)
    rows = sheet.nrows
    # 排除第一列循环每行
    for i in range(rows)[1:]:
        # 获取每行的第二列
        cell2 = sheet.cell_value(i, 2)
        # 获取每行的第三列
        cell3 = sheet.cell_value(i, 3)
        file.writelines("update " + cell2 + " set DEL_FLAG = 'Y' where id = " + cell3 + ";")
        file.writelines("\n")


def read_excel(num):
    # 打开文件
    workbook = xlrd.open_workbook(r'D:\\代客下单订单导入20200601184020.xls')
    # sheet 0
    sheet = workbook.sheet_by_index(0)

    rows = sheet.nrows  # 获取有多少行
    cols = sheet.ncols  # 获取有多少列
    print("共有", rows, "行")
    print("共有", cols, "列")

    path = "D:\\"
    # 生产文件名称、位置
    full_path = path + "perform.sql"
    file = open(full_path, 'w')

    # 获取第num列上所有数据
    cols = sheet.col_values(num)
    # 循环处理选择的列,第二行开始遍历
    for col in cols[1:]:
        file.writelines("select * from user where id =" + col + ";")
        file.writelines("\n")


if __name__ == '__main__':
    # read_excel(3)
    read_excel1()

