June.14
========

June.14 is the 165th day of the year in the Gregorian calendar, and a mailbot created by oncemore2020 on June.14

对指定Excel文档中记录的人事信息，逐个发送给内容行中的邮箱与其进行核对。当然，也可以改造用于其他用途，比如发工资单;-)

## 1. 准备开发环境

以 Python3.6.x 版本为例，下载32位版本的安装程序，以便支持发布的二进制程序在32位Windows上运行。

安装包下载地址：[https://www.python.org/ftp/python/3.6.8/python-3.6.8.exe](https://www.python.org/ftp/python/3.6.8/python-3.6.8.exe)。

安装路径设置到 `C:/Python36`，或者安装时勾选添加 Python 到环境变量。

安装完成后，需要安装依赖的 Python 包，使用 pip 包管理器安装即可，首先 CMD 进入项目根目录，然后执行

```bash
c:\python36\python.exe -m pip install -r requirements.txt
```

## 2. 构建发布应用安装程序

CMD 进入到项目根目录，然后执行

```bash
c:\python36\python.exe setup.py build
```

生成的二进制程序在 `build` 目录下。

若要发布安装包，则运行

```bash
c:\python36\python.exe setup.py build bdist_msi
```

生成的安装包在 `dist` 目录下。
