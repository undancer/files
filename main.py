__author__ = 'undancer'

import os
from xlrd import open_workbook


def findExcelFile(path):
    """
    查找指定路径下的excel文件

    :param path:
    :return:
    """
    _list = []
    for root, dirs, files in os.walk(path):
        clean(dirs)
        _list += [os.path.join(root, name) for name in files if
                  not (not endswith(name, '.xls', 'xlsx') or name.startswith('~'))]
    return _list


def clean(dirs):
    """
    文件目录里不包含.svn

    :param dirs:
    """
    for name in ['.svn']:
        if name in dirs:
            dirs.remove(name)


def endswith(name, *suffixs):
    """

    判断文件名结尾，excel文件有2种扩展名

    :param name:
    :param suffixs:
    :return:
    """
    for suffix in suffixs:
        if str.endswith(name, suffix):
            return True
    return False


def main():
    paths = findExcelFile('/share/svn')
    print(paths)

    count, error = 0, 0

    for path in paths:
        workbook = open_workbook(path)
        for sheet in workbook.sheets():
            for x in range(1, sheet.nrows):
                for y in range(sheet.ncols):
                    title = str(sheet.cell_value(0, y))
                    value = str(sheet.cell_value(x, y))
                    if len(title) > 0 and len(value) > 0:
                        try:
                            msg = '{0:s} | {1:d}:{2:d} -> {3:s}:{4:s}'.format(path, x, y, title, value)
                            print(msg)
                            count += 1
                        except UnicodeEncodeError:
                            error += 1
                            pass
    print('count:%d error:%d' % (count, error))


if __name__ == '__main__':
    main()
