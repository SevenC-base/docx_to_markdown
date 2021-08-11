#!/usr/bin/env python
# -*- coding:UTF-8 -*-
# 2021/08/10 周二 15:23:14
# By  Hasaki-h1
import os, sys, subprocess, uuid


def get_file_list(directory):
    """单独一个目录"""
    files_list = []
    files_path_list = []
    files_p_list = []

    if os.path.exists (directory):
        directory_n = directory
    else:
        print ("%s 不是一个有效的目录！！！" % directory)
        sys.exit ()

    # 遍历目录下读取可读文件
    all_files_directory = os.walk (directory_n, topdown=True, followlinks=True)
    for root, dirs, files in all_files_directory:
        # 获取文件路径
        for f_name in files:
            if ".docx" in f_name:
                # print(f_name)
                file_path_d = os.path.join (root, f_name)
                file_path_m = os.path.join (root, f_name.replace (".docx", ".md"))
                files_path_list.append (file_path_d)
                files_p_list.append (file_path_m)

    return files_path_list, files_p_list


def convert_md(pandoc_path, f_docx, f_md, images_store_p):
    try:
        image_store_path_create = images_store_p + str (f_docx.split ('\\')[-1]).replace (".docx", "") + str (
            uuid.uuid4 ())
        cmd = pandoc_path + " \"{}\" -f docx  -t markdown  -o \"{}\"  --extract-media=\"{}\"".format (f_docx, f_md,
                                                                                                  image_store_path_create)
        res = subprocess.Popen (cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        sout, serr = res.communicate ()
        res.wait ()
        sout.decode ("gbk")
    except Exception as e:
        serr.decode ("gbk")
    finally:
        pass


def main(pandoc_path, dirctory, images_store_path):
    fpl_docx, fpl_md = get_file_list (dirctory)
    print ('🚀-----------------起飞------------------🚀')
    for fpl_d in fpl_docx:
        progress = '{:.2f}%'.format (((fpl_docx.index (fpl_d) + 1) / len (fpl_docx)) * 100)
        # print(fpl_d,fpl_md[fpl_docx.index(fpl_d)])
        convert_md (pandoc_path, fpl_d, fpl_md[fpl_docx.index (fpl_d)], images_store_path)
        print ('✈ 进度：{} ------  {}  ->>>--->>> 转换完成!'.format (progress, str (fpl_d.split ('\\')[-1])))


if __name__ == '__main__':
    # 配置pandoc 路径
    pandoc_path = "C:\\Users\\Administrator\\Downloads\\pandoc.exe"
    # 配置docx文件路径，会遍历该路径下的所有docx文件
    dirctory = "D:\\笔记\\"
    # 配置图片存储路径
    images_store_path = "D:\\笔记\\images\\"

    main(pandoc_path, dirctory, images_store_path)
