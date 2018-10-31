# merge_excel_sheets
支持合并多个excel工作簿的工作表数据


**合并后数据请检查下下格式和显示，可能有bug验证不充分，此外，如果需要自动扫描目录，请使用下个版本 。**


使用python2.7 环境下 openpyxl2.5.9 读写excel . 主程序代码是merge_sheets.py

小白用户可以直接执行以下脚本（如果已经安装了anaconda, 以下脚本初始化环境和测试）:
```
	install_env.bat  
```


#注意事项:
	1. 需要安装python 环境(有anaconda也行), 需要安装openpyxl2.5.9 (进入到目录中执行python setup.py install)
	2. 需要将merge_sheets.py  文件放在和excel 文件一起（同目录） 。
	3. 确保各个excel工作簿中的sheet 名字一样，并且每个sheet 的表头是一样的顺序（列顺序）！！




*以下运行命令都是在运行目录（最好是excel所在目录）执行CMD 命令运行 。
打开CMD的方式是， 在文件所在目录，按住Shift 然后点击右键，选择“在此处打开命令窗口”，会弹出黑框 ，在里面可以输入命令 。*


#使用方法1.（黑框中运行命令）:
```
	python merge_sheets.py test.xlsx test01.xlsx test02.xlsx test03.xlsx
```
	其中: test.xlsx 是最终合并的文件名,  test01,xlsx , test02.xlsx , test03.xlsx 是待合并的文件列表.


#使用方法2.如果要合并的文件非常多, 可以运行:
```
	python merge_sheets.py 最终文件.xlsx -dir testdir
```
	这样会扫描testdir目录的所有xlsx文件,合并到"最终文件.xlsx"
    
    假设需要把测试01.xlsx , 测试02.xlsx , 测试03.xlsx , 测试04.xlsx 合并为 最终文件.xlsx ，那么命令为：
```
	python merge_sheets.py  最终文件.xlsx  测试01.xlsx  测试02.xlsx 测试03.xlsx  测试04.xlsx 
```



