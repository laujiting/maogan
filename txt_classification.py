'''
Author: Liu Jiting
Date: 2019/10/06
Description: Traverse the Excel-directory,
and read the content about txt name and its label,
move the txt to the corrtct directory.
'''

import os
import shutil
from auto import *


def read_excel(file):
    wb = app_readExcel.books.open(file)
    sht = wb.sheets['Sheet1']
    # print('A2:{}'.format(sht.range('A2').value))
    # print(type(sht.range('A2').value))
    rows = sht.api.UsedRange.Rows.count
    file_name = [int(item) for item in sht.range('B2:B{}'.format(rows)).value]
    file_name = [item for item in file_name if item]# [:4]  # 去除空值，[:4]是暂时只取前4个值进行测试
    label = sht.range('K2:K{}'.format(rows)).value
    label = [item.rstrip('类') for item in label if item]# [:4]  # 去除空值
    print(len(label), label)
    for item in label:
        if item in 'Ⅰ':
            print(len(file_name), file_name)
            print(len(label), label)
    '''prj_mrt_path = prj_path + '\\' + 'prj{}'.format(i)
    mrt_name = [item for item in os.listdir(prj_mrt_path) if item[-3:] == 'mrt']
    print(mrt_name)
    for j, file in enumerate(mrt_name[:4]):
        # print(prj_path + '\\' + 'prj{}'.format(i) + '\\' + file)
        file = file.replace('mrt', 'txt')  # 修改后缀
        print(file)
        src_path = prj_mrt_path + '\\' + file
        dst_path = prefix + '\\' + '..' + '\\' + 'classification_data'
        print(label[j])
        move_txt(src_path, dst_path, file, label[j])'''

    wb.close()
    app_readExcel.quit()
    return file_name, label


if __name__ == '__main__':
    # 所有工程目录的父目录
    # path = 'F:\\automaogan\\splitmpj'
    # Excel整合在一个文件夹下
    excel_dir = 'E:\\研究生学习\\项目\\锚杆识别\\锚杆Excel'
    # 导出每个工程目录的名称
    # prj_list = os.listdir(path)
    # print('Origin list:{}'.format(prj_list))
    # 保证项目列表升序
    # prj_list = sorted(prj_list,
    #                   key=lambda item: int(item.lstrip('prj').rstrip('.mpj')),
    #                   reverse=False)
    # print('Sorted list:{}'.format(prj_list))
    # 对每个工程目录遍历得到其mrt数据文件列表，元素为文件的绝对路径
    '''for prj in prj_list:
        print('project:{}'.format(prj))
        prj_path = path + '\\' + prj
        prj_file = os.listdir(prj_path)
        print('project file:{}'.format(prj_file))
        mrt_dir = prj_path + '\\' + prj_file[0]
        # 打开mrt目录读取数据文件名
        if os.path.isdir(mrt_dir):
            mrt_list = os.listdir(mrt_dir)
            mrt_list = [mrt_dir + '\\' + item for item in mrt_list]
            print('mrt:{}'.format(mrt_list))
            print(len(mrt_list))
        #'''
    length = 0
    for i, file in enumerate(os.listdir(excel_dir)):
        excel_path = excel_dir + '\\' + file.lstrip('~$')
        txt_list, label = read_excel(excel_path)
        length = i
    print(length)