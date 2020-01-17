# coding:utf-8

import os

if __name__ == '__main__':
    path = r'E:\马原视频'
    fileList = os.listdir(path)
    for f in fileList:
        old = os.path.join(path, f)
        if os.path.isdir(old):
            continue
        filename, ext = os.path.splitext(f)
        # if ext == '.ans':
        #     print('filename: %s\text: %s'%(filename, ext))
        filename = filename.replace('+', '-')
        if filename[2] != '组':
            filename = filename[:2] + '组' + filename[2:]
        new = os.path.join(path, filename + ext)
        os.rename(old, new)
