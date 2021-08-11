#### 1.简介

一个辅助脚本，利用工具[pandoc](https://github.com/jgm/pandoc/releases/tag/2.14.1)，以单线程的方式将目录下所有docx文件转为markdown格式。

pandocx：https://github.com/jgm/pandoc/releases/tag/2.14.1

#### 2.使用配置：

修改如下配置：

```python
# 配置pandoc 路径
pandoc_path = "C:\\Users\\Administrator\\Downloads\\pandoc.exe"
# 配置docx文件存在的路径，会遍历该路径下的所有docx文件
dirctory = "D:\\笔记\\"
# 配置图片存储路径
images_store_path = "D:\\笔记\\images\\"
```

#### 3.运行结果如下

![](E:\9-Program\GIT\Docx_to_markdown\1.png)

