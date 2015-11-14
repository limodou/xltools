# xltools - Excel 模板化读取和写入工具

## 什么是xltools

它的创建是为了减少在读取和写入Excel时对格式,位置以及数据处理的重复性工作. 比如我们在读取一个Excel
文件时,要先分析文件的格式,比如哪几列是数据,从哪里开始,数据类型是什么,提取出来的数据如何组织.再比如
写入一个Excel时,我们需要分析Excel哪些单元格要写入,格式,样式,用什么数据来填充.在大部分情况下,我们
会直接打开Excel,用程序方式来直接读取或写入,但是这种方式既不直观,也不方便修改.因此,我设计是通过在
Excel中定义一些特殊的文本作为标记,然后通过对这些标记进行解析,识别出它的字段名,有哪些过滤处理,
还包括对数据的循环处理等格式内容.因此这样的Excel可以称之为模板,通过模板实现格式与数据的有效结合,
对于常见的Excel处理可以减轻开发工作量.

使用方法:

1. 将要处理的Excel拷贝,然后加工成模板.(读与写的模板在标签定义上略有差异)
1. 使用Reader或Writer类对模板进行处理,得到相应的结果

## 读处理

### 标签格式

在用于读取的Excel模板中通过 `{{xxx}}` 这样的形式来定义标签,常见标签格式为:

{{xxx}} --
    其中xxx为要抽取数据对应的数据项名称
{{for xxx}} --
    for开始的标签表示下面定义的是循环.此标签需要定义在第一列,并且独占一行.它最后需要 `{{end}}` 表示结束.
    在它之下,在 `{{end}}` 之前的为循环的内容.可以是多行. `xxx` 为循环所对应的数据项名称.
{{end}} --
    表示循环结束

### 简单循环模板

下面我们定义一个简单的循环数据的模板

A | B | C
{{for items}} | |
{{f1}} | {{f2}} | {{f3}}
{{end}} | |

通过上面的模板我们就定义了一个简单的循环. 当数据是:

A | B | C
A1 | B1 | C1
A2 | B2 | C2
A3 | B3 | C3

我们可以使用下面的代码:

```
>>> from xltools import Reader
>>> from pprint import pprint
>>> x = Reader('template_1.xlsx', 'sheet1', 'data_1.xlsx')
>>> pprint(x.result)
[{u'items': [{u'f1': u'A1', u'f2': u'B1', u'f3': u'C1'},
             {u'f1': u'A2', u'f2': u'B2', u'f3': u'C2'},
             {u'f1': u'A3', u'f2': u'B3', u'f3': u'C3'}]}]
```

我们可以看到整个结果是一个 list ,然后每个元素是一个dict,其中,key `items`就是模板中设定的循环数据项,它
的值就是对应的每行数据.每行数据又组织为一个dict,key值分别对应模板中的数据项名.

所以我们最终得到的是一个多级的dict数组.

`Reader` 的原型是:

```
Reader(template_file, sheet_name, input_file,
                 use_merge=False, merge_keys=None, merge_left_join=True,
                 merge_verbose=False, callback=None)
```

`template_file` --
    是模板文件
`sheet_name` --
    为指定的sheet名
`input_file` --
    数据文件,如果只有一个文件,则可以为字符串,如果是多个,则可以是tuple或list.

    `Reader` 在处理数据文件时,通常会根据sheet名,在数据文件中也查找同名的sheet页.如果没找到,则整个文件
    所有sheet页都会进行搜索.如果需要指定要处理的sheet页名称,还可以将 `input_file` 定义为:

    ```
    [('filename', '*'), ('filename', 'sheetname1', 'sheetname2'), 'filename']
    ```

    当值为tuple时,第一个元素为数据文件名,后面的元素表示对应的sheet名,'*'表示所有sheet页.
`callback` --
    回调,当处理完毕时,执行回调函数.回调函数原型是:

    ```
    def callback(data)
    ```

    其中 `data` 就是 `Reader` 处理之后的数据.

关于merge的相关参数,具体内容参见 `Merge` 的说明.

### 复杂例子

在简单例子中,我们只定义了一个循环,并且所有变量都定义在循环体中,下面我们定义一个更复杂一些的例子.

Section | {{section}} |
Input | {{input}} |
A | B | C
{{for request}} | |
{{f1}} | {{f2}} | {{f3}}
{{end}} | |
 | |
Output | {{output}} |
A | B | C
{{for response}} | |
{{f1}} | {{f2}} | {{f3}}
{{end}} | |

可以看到这个模板有两个循环,同时还有一些非循环项.我们可以把这个例子看成报文规范的示例,下面我们的
数据将保存多条报文,如:

Section | Package1 |
Input | Input-001 |
A | B | C
A1 | B1 | C1
A2 | B2 | C2
A3 | B3 | C3
 | |
Output | Output-001 |
A | B | C
X1 | Y1 | Z1
X2 | Y2 | Z2
X3 | Y3 | Z3
 | |
 | |
Section | Package2 |
Input | Input-002 |
A | B | C
A4 | B4 | C4
A5 | B5 | C5
A6 | B6 | C6
 | |
Output | Output-002 |
A | B | C
X4 | Y4 | Z4
X5 | Y5 | Z5
X6 | Y6 | Z6

让我们执行程序:

```
>>> from xltools import Reader
>>> from pprint import pprint
>>> x = Reader('template_2.xlsx', 'sheet1', 'data_2.xlsx')
>>> pprint(x.result)
[{u'input': u'Input-001',
  u'output': u'Output-001',
  u'request': [{u'f1': u'A1', u'f2': u'B1', u'f3': u'C1'},
               {u'f1': u'A2', u'f2': u'B2', u'f3': u'C2'},
               {u'f1': u'A3', u'f2': u'B3', u'f3': u'C3'}],
  u'response': [{u'f1': u'A', u'f2': u'B', u'f3': u'C'},
                {u'f1': u'X1', u'f2': u'Y1', u'f3': u'Z1'},
                {u'f1': u'X2', u'f2': u'Y2', u'f3': u'Z2'},
                {u'f1': u'X3', u'f2': u'Y3', u'f3': u'Z3'}],
  u'section': u'Package1'},
 {u'input': u'Input-002',
  u'output': u'Output-002',
  u'request': [{u'f1': u'A4', u'f2': u'B4', u'f3': u'C4'},
               {u'f1': u'A5', u'f2': u'B5', u'f3': u'C5'},
               {u'f1': u'A6', u'f2': u'B6', u'f3': u'C6'}],
  u'response': [{u'f1': u'A', u'f2': u'B', u'f3': u'C'},
                {u'f1': u'X4', u'f2': u'Y4', u'f3': u'Z4'},
                {u'f1': u'X5', u'f2': u'Y5', u'f3': u'Z5'},
                {u'f1': u'X6', u'f2': u'Y6', u'f3': u'Z6'}],
  u'section': u'Package2'}]
```

可以看到我们按照预期抽取出来了两条数据.其中 `requst` 和 `response` 分别对应报文的请求和输出字段列表.