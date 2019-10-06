import os
import shutil
import xlwings as xw


# 程序不可见，只打开不新建工作薄，屏幕更新关闭
app_readExcel = xw.App(visible=False, add_book=False)
app_readExcel.display_alerts = False
app_readExcel.screen_updating = False


def make_dir(prefix, name):
    dir_path = prefix + '\\' + 'prj{}.mpj'.format(name)
    if os.path.exists(dir_path):
        pass
    else:
        os.mkdir(dir_path)
    return dir_path


def move_txt(src, dst, file, label):
    assert os.path.isabs(src)
    if label == 'Ⅰ':
        d_name = '1'
    elif label == 'Ⅱ':
        d_name = '2'
    elif label == 'Ⅲ':
        d_name = '3'
    else:
        d_name = '4'
    print(d_name)
    print(dst + '\\' + d_name + '\\' + file)
    shutil.copy(src, dst + '\\' + d_name + '\\' + file)


def create_prj(prefix_path):
    # mrt转入工程目录
    for index, i in enumerate(range(41, 88)):
        prj_path = make_dir(path, name=i)  # 创建单个工程目录
        print('dir_path:{}'.format(prj_path))
        shutil.copy(prefix_path, prj_path + '\\' + 'prj{}.mpj'.format(i))  # 移动工程文件到刚创建的目录并重命名
        mrt_dst = prj_path + '\\' + 'prj{}'.format(i)  # 构造mrt的绝对路径
        # if os.path.exists(mrt_dst):
        #     pass
        # else:
        #     mrt_dst = os.mkdir(mrt_dst)
        # 创建数据文件的目录
        # 这里的问题是如果已存在目录，会抛出错误，因此不再创建目录

        print('mrt_dst:{}'.format(mrt_dst))  # 显示文件dst绝对路径
        print('mrt_list{}:{}'.format(index, mrt_path + '\\' + mrt_list[index]))  # 显示文件src绝对路径
        shutil.copytree(mrt_path + '\\' + mrt_list[index], mrt_dst)  # 移动文件
    # print(os.listdir(path))  # 显示所有创建的工程文件


def txt_classification(prefix_path):
    # prefix_path为所有工程目录的父目录
    for index, i in enumerate(range(41, 42)):
        prj_path = make_dir(prefix_path, name=i)  # 已创建的目录，直接返回目工程目录路径
        file = [item for item in os.listdir(prj_path) if '.xlsx' in item][0]
        print(file)
        excel_path = prj_path + '\\' + file.lstrip('~$')
        print(excel_path)
        wb = app_readExcel.books.open(excel_path)
        sht = wb.sheets['Sheet1']
        # print('A2:{}'.format(sht.range('A2').value))
        # print(type(sht.range('A2').value))
        rows = sht.api.UsedRange.Rows.count
        file_name = [int(item) for item in sht.range('B2:B{}'.format(rows)).value]
        file_name = [item for item in file_name if item][:4]  # 去除空值
        label = sht.range('K2:K{}'.format(rows)).value
        label = [item.rstrip('类') for item in label if item][:4]  # 去除空值
        print(len(file_name), file_name)
        print(len(label), label)
        prj_mrt_path = prj_path + '\\' + 'prj{}'.format(i)
        mrt_name = [item for item in os.listdir(prj_mrt_path) if item[-3:] == 'mrt']
        print(mrt_name)
        for j, file in enumerate(mrt_name[:4]):
            # print(prj_path + '\\' + 'prj{}'.format(i) + '\\' + file)
            file = file.replace('mrt', 'txt')  # 修改后缀
            print(file)
            src_path = prj_mrt_path + '\\' + file
            dst_path = prefix_path + '\\' + '..' + '\\' + 'classification_data'
            print(label[j])
            move_txt(src_path, dst_path, file, label[j])

        wb.close()
        app_readExcel.quit()


if __name__ == '__main__':
    path = 'F:\\automaogan\\splitmpj'  # 在当前目录下创建所有工程的文件夹
    file_path = 'F:\\automaogan\\1DY.mpj\\1DY.mpj'  # 一个模板工程文件的绝对路径
    mrt_path = 'F:\\automaogan\\mrt_fenzu'  # 吴分组好的数据的父目录
    # print(os.listdir(mrt_path))
    mrt_list = os.listdir(mrt_path)  # 列表，元素为所有mrt的文件名
    print(mrt_list)
    # create_prj(file_path)
    # txt_classification(path)
