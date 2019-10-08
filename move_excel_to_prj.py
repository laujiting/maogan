import os
import shutil


def move_excel(src, dst):
    assert os.path.isabs(dst)
    shutil.copy(src, dst)


if __name__ == '__main__':
    excel_dir = 'E:\\研究生学习\\项目\\锚杆识别\\锚杆Excel'
    prj_dir = 'E:\\研究生学习\\项目\\锚杆识别\\test2'  # 根据实际情况改变，改为所有工程的父目录
    excel_list = [item for item in os.listdir(excel_dir) if item[:2] != '~$']
    excel_list = sorted(excel_list,
                        key=lambda item: int(item.lstrip('prj').rstrip('.xlsx')),
                        reverse=False)
    print(excel_list)
    for excel in excel_list[1:]:
        excel_abs_src = excel_dir + '\\' + excel
        midfix = excel.rstrip('.xlsx')
        print(midfix)
        # 构造excel在对应工程中的绝对路径
        excel_abs_path = prj_dir + '\\' + midfix + '.mpj' + '\\' + excel
        print(excel_abs_path)
        move_excel(excel_abs_src, excel_abs_path)
    # 先修改这个程序里的路径，把excel移动到工程下，再运行auto，将txt分类
    # print(excel_list)