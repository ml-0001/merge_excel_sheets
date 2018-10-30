

@echo 
@echo 
@echo hello,world!

@echo 我是小白用户,我的电脑已经安装了anaconda, o(∩_∩)o 哈哈,我现在想初始化环境(只需要执行一次).


cd openpyxl-2.5.9

python setup.py install




@echo 安装完成

@echo 我现在要测试下工具是否可用,所以运行了如下命令(我要把三个文件合并到一起):
cd ../
@echo python merge_sheets.py 测试合并结果文件.xlsx 测试01.xlsx 测试02.xlsx 测试03.xlsx

python merge_sheets.py 测试合并结果文件.xlsx 测试01.xlsx 测试02.xlsx 测试03.xlsx

@pause
