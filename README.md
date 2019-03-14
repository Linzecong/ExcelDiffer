# An Excel Differ

# 如何使用源码运行

1. Install Python3
2. pip install PyQt5
3. python3 MainWindow.py

# 如何打包

pyinstaller MainWindow.spec

# 使用文档

本程序实现的主要功能有：
比较两个Excel的行增删和交换信息
比较两个Excel的列增删和交换信息
比较两个Excel的合并单元格的增删信息
可以自定义字体和颜色，自带一个夜间模式
窗口之间可以随心所欲的拖动

1.先打开需要比较的Excel表格

![image](https://github.com/Linzecong/ExcelDiffer/docimages/openexcel.png)

可以在菜单栏打开也可以在工具栏打开
我们先打开一个旧的Excel表格，然后再打开一个新的Excel表格

2.打开后点击菜单栏或工具栏的开始比较按钮开始Diff

![image](https://github.com/Linzecong/ExcelDiffer/docimages/ana.png)

具体Diff时间见后面算法分析

3.点击列表内的结果可以动态的查看差异

![image](https://github.com/Linzecong/ExcelDiffer/docimages/result.png)

4.点击格式菜单或工具栏可以自定义颜色和字体

![image](https://github.com/Linzecong/ExcelDiffer/docimages/font.png)

5.点击切换模式可以更改为夜间模式，再次点击可以更改回来
可以通过修改style.qss文件自定义样式

![image](https://github.com/Linzecong/ExcelDiffer/docimages/night.png)


# 关于算法

主要是算法是 ``最长公共子序列``

设旧表格 行数量为 N，列数量为M
设新表格 行数量为 A，列数量为B

一般而言，列的数量较少，我们先求出列的增删信息

新建一个数组 map 代表 新表格的每一列 与 旧表格的每一列的对应关系，默认为-1

然后对于新表格的每一列，与旧表的每一列都进行比较，求一个``最长公共子序列``
我们将公共子序列长度最长的那个，作为新表与旧表的对应
因此 总的时间复杂度为 O(M * B * N * A) 其中 N * A 为``最长公共子序列``的时间复杂度


我们得到了对应关系后，可以通过一个遍历，求出旧表中没有被对上的列，那些列就是被删除的列
同理，新表中没有对应的列，就是新增的列。


我们得到了对应关系后，将列重排，然后开始对行做同样的操作
对行求增删信息的时间复杂度为 O(N * A * M * B) 其中 M * B 为``最长公共子序列``的时间复杂度


得到了所有的对应关系后，我们可以对map数组求``最长上升子序列``然后不在该子序列中的，就是被交换了的行或列

然后我们有了对应数组，我们只需要遍历整个对应数组，然后比较单元格的值，就可以知道单元格是否被修改 时间复杂度为 O(max(N,A)*max(M,B))


总的复杂度还是很大的，在行数为300左右，列数为10左右的时候，比较需要5s左右的时间



因此考虑优化，我们可以设定一个匹配阈值PP，当``最长公共子序列长度``比例达到这个阈值的时候，我们可以认定此行已被匹配，不再对剩下的行做比较。
在代码中可以通过修改 ``Algorithm.py`` 来获取更好的效果
```
class MyAlg(QObject):
    YZ = 0.5 #匹配成功的阈值 0~1越高精准度越高，
    PP = 0.8 #用于不用跑完全部行，加速比较 ，PP >= YZ
```

这里还有一个优化，就旧表中已经被匹配的行，我们不再与它做匹配，这里可是省下一半的时间。

但是总的时间复杂度还是很高。因此考虑优化 ``最长公共子序列长``算法，一般的求``最长公共子序列``的时间复杂度为 O(N*M)
是平方级别的。但是在应用当中，我们对表格的修改不是很大，因此可以使用更为高效的求``最长公共子序列``的算法

http://www.xmailserver.org/diff2.pdf

在这篇论文中介绍了一种时间复杂度为 O(ND)的算法，其中 D为两个序列的不同个数。我们可以使用深搜+二分优化这个算法，似的时间复杂度接近O(N)的级别

这里有一个C++实现版本，有机会翻译成Python，并加入到此软件中

https://github.com/abcdabcd987/sit/blob/master/src/Diff.cc#L150


欢迎讨论更优秀的算法！！