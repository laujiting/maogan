import os
import shutil
import xlwings as xw
from auto import move_txt
from txt_classification import read_excel


# 程序不可见，只打开不新建工作薄，屏幕更新关闭
app_readExcel = xw.App(visible=False, add_book=False)
app_readExcel.display_alerts = False
app_readExcel.screen_updating = False


if __name__ == '__main__':
    excel_dir = 'F:\\maogan_final\\锚杆Excel'
    txt_dir = 'F:\\mrtM'
    classification_dir = 'F:\\automaogan\\classification_data'
    length = 0
    excel_list = [item for item in os.listdir(excel_dir) if item[:2] != '~$']
    excel_list = sorted(excel_list,
                        key=lambda item: int(item.lstrip('prj').rstrip('.xlsx')),
                        reverse=False)
    '''print(len(excel_list[:10]))
    print(len(excel_list[10:20]))
    print(len(excel_list[20:30]))
    print(len(excel_list[30:40]))
    print(len(excel_list[40:50]))
    print(len(excel_list[50:70]))
    print(len(excel_list[70:]))'''
    for i, file in enumerate(excel_list):
        excel_path = excel_dir + '\\' + file.lstrip('~$')
        txt_list, label = read_excel(excel_path)
        txt_list = sorted(txt_list)
        for index, txt in enumerate(txt_list):
            txt_abs_path = txt_dir + '\\' + '{}.txt'.format(txt)
            move_txt(txt_abs_path, classification_dir, '{}.txt'.format(txt), label[index])
        # length = i + 1
            print(txt_abs_path)
    # print(length)