'''
    1. 扫描多个作业文件夹，将其名称添加到数组中
    2. 扫描单个作业文件夹所有提交的文件，将这些文件名添加到文件名数组中
    3. 根据学生表一一比较已提交的文件名列表中的名称
    4. 根据学生的学号，在统计表中修改其提交的次数
'''

import os
import xlrd
from xlutils.copy import copy
import pandas as pd


# 获取学生提交文件的名称列表
def get_file_name(sub_dir):
    del_ip_file_name_list = list()
    # 1. 读取作业1文件夹下的文件，将获取到的文件名存储在列表中
    # 该方法返回的是 列表
    file_name_submitted = os.listdir(sub_dir)
    # 输出列表中的索引元素
    for file in file_name_submitted:
        # 将各文件名中的首部的IP部分删除掉
        del_ip_file_name_list.append(file[file.find("_") + 1:])
    # print(del_ip_file_name_list)
    return del_ip_file_name_list


# 获取每个提交作业的文件夹中的文件列表
def get_dir_subbed_all_filename(subbed_dir):
    for sub_index in range(len(subbed_dir)):
        print('++++++++++++++++++++++++++++++')
        print(subbed_dir[sub_index])
        # 获取学生提交文件的名称列表
        subbed_file_name_list = get_file_name(subbed_dir[sub_index])
        statistics_method(sub_index, subbed_file_name_list)
        print("-----------------------------")


# 统计学生提交的
def statistics_method(index, subbed_file_name_list):
    # 获取学生字典， 学号:姓名
    df = pd.read_excel(stu_list_path)
    # 获取学生名单，以字典形式
    stu_dict = dict(zip(df['学号'], df['姓名']))
    # 统计是否提交的学生
    for item in stu_dict.items():
        flag = 0
        row_value = df[df['学号'] == item[0]].index.values[0]
        # 判断提交的文件列表是否为空
        if subbed_file_name_list:
            for f in subbed_file_name_list:
                if f.find(str(item[0])) >= 0 and f.find(item[1]) >= 0:
                    ws_stu.write(int(row_value) + 1, index + 2, 1)
                    # 从文件列表中删除已提交的学生文件，减少循环次数
                    subbed_file_name_list.remove(f)
                    # 已查找，将flag设为1
                    flag = 1
                    # 查找到就退出内层文件列表循环
                    break
            # 未查找到该学生的文件
            if flag == 0:
                print("未提交：", item)
                ws_stu.write(int(row_value) + 1, index + 2, 0)
        else:
            print("未提交：", item)
            # xlrd 行和列都是从0开始计数
            ws_stu.write(int(row_value) + 1, index + 2, 0)
    wb_stu_workbook.save('test_1.xls')


if __name__ == '__main__':

    # 作业的数量
    task_num = 3
    # 各作业文件夹的路径名称列表
    sub_path_list = list()
    # 学生提交作业的文件夹路径
    sub_path = "./TestDirectory/作业"
    for i in range(1, task_num + 1):
        sub_path_list.append('{0}{1}{2}'.format(sub_path, i, '/'))
    print(sub_path_list)
    # 学生名单文件路径
    stu_list_path = r"./TestDirectory/test.xls"

    # 打开学生名单Excel
    rb_stu_workbook = xlrd.open_workbook(stu_list_path)
    # 复制一份学生名单
    wb_stu_workbook = copy(rb_stu_workbook)
    # 选择第一个sheet表
    ws_stu = wb_stu_workbook.get_sheet(0)
    # 开始统计
    get_dir_subbed_all_filename(sub_path_list)
