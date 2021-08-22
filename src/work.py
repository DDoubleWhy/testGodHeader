# pip install gitpython
# pip install xlrd

import xlrd
import re
from git.repo import Repo

print("helloWorld!")
book = xlrd.open_workbook("D:\warning.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
dline = sh.col_values(3, start_rowx=0, end_rowx=None)
eline = sh.col_values(4, start_rowx=0, end_rowx=None)
res = {}
for i in range(len(dline)):
    if (None == re.search("inval*", dline[i])):
        res[dline[i]] = eline[i]

for s in res.keys():
    print(s)
    with open(s, mode = 'r+') as myFile:
        print("open it !")
        for i,l in enumerate(myFile.readlines()):
            if (re.search(res[s], l) != None):
                print("get it !")
                del_line = i + 1
    

    with open(s, 'r') as old_file:
        with open(s, 'r+') as new_file:

            current_line = 0

            # 定位到需要删除的行
            while current_line < (del_line - 1):
                old_file.readline()
                current_line += 1

            # 当前光标在被删除行的行首，记录该位置
            seek_point = old_file.tell()

            # 设置光标位置
            new_file.seek(seek_point, 0)

            # 读需要删除的行，光标移到下一行行首
            old_file.readline()

            # 被删除行的下一行读给 next_line
            next_line = old_file.readline()

            # 连续覆盖剩余行，后面所有行上移一行
            while next_line:
                new_file.write(next_line)
                next_line = old_file.readline()

            # 写完最后一行后截断文件，因为删除操作，文件整体少了一行，原文件最后一行需要去掉
            new_file.truncate()


print("#####")
print(res)