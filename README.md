# excelReader Excel文件读取
a simple python script that read the Microsoft Excel file.

一个读取Excel文件的简单python脚本

## What it can do 它能干啥
parsing a single sheet `*.xlsx` file into a python list. An abstract usage:

把一个单表格的`*.xlsx`文件读取为一个python的列表。 下面是一个抽象一点的用法实例：

```
from ThisScript import TheReaderFunction
target_file = "path/to/file.xlsx"
table = TheReaderFunction(target_file)
C6 = table[5][2]
print C6   # type(C6) == str
```

For detailed sample, see the last part of the script.

具体的看python脚本的最后，会有实际情况下的用法，其实就是调用个函数的事

## Limits 不足

* only supports single sheet file without any graph, excel programs. And the sheet should have default name `sheet1` (Yes, the simplest type you could imagine)  
* 只支持单表格的文件，并且不能有任何图表、Excel编程。就是最简单的那种。并且那个表格的名字必须是默认的`sheet1`
* not fully tested, may have some issues among different version of Excels. Good luck.
* 未经过完整测试，可能在不同版本的Excel上有问题

## xlsx file format  xlsx文件格式
If you are just looking for a handy tool, skip this part.

这部分简单叙述一下xlsx的文件格式，没兴趣就跳过吧。

xlsx is actually a zip file. The details could be found by using :

xlsx文件实际上是一个zip压缩包。用unzip或者随便什么就可以把它打开：

```
$ unzip sample.xlsx
```

The inner content contain some data. And we only need two of them. `xl/sharedStrings.xml` and `xl/worksheets/*.xml`. Well, maybe more than two.

压缩包里的内容，我们只需要两个：`xl/sharedStrings.xml` 和 `xl/worksheets/*.xml`

Each *.xml under folder `xl/worksheets/` stands for a sheet. By default, the sheet has name `sheet1`. Thus the xml file has name `sheet1.xml`.

`xl/worksheets/`下的每个xml文件都对应一个表格，例如名为`sheet1`的表格就对应`sheet1.xml`

All data in xlsx are regarded as strings, and stored in the sharedString.xml. Duplicated strings are dereplicated. Maybe that is just why it's called `sharedString.xml`.

xlsx文件中所有数据都被认为是字符串，存储在`sharedString.xml`中。重复的字符串会被消除，难怪文件叫做共享字符串。

The `sheet1.xml` specified the width and height of a sheet, and link every grid to a string in `sharedString.xml`. A sheet then could be retrieved.

`sheet1.xml`中指定了表格的大小，并且将每一个格子链接到sharedString.xml的一个字符串中。有了这些知识，就可以读取一个简单的xlsx文件了。

## Improvement 改进
I mean, improvement by YOUR hand. For me, this script is just enough for application.

这里我说的是你对这个脚本做出的改进。对我而言这个脚本已经够用了。

#### Multi sheet support 多表格支持
This is simple. Find all the *.xml under `xl/worksheets/`, and return multiple list. Done.

很简单，把`xl/worksheets/`下的所有xml文件都挨个处理一遍，返回多个列表就完了。

#### More 可能还有其他的。。。
