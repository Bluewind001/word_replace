# word_replace

###程序功能：
使用多个替换规则，替换多个docx文件

###运行说明：
1. 这是一个Maven的项目，clone下代码使用mvn package运行程序进行打包，在target目录生成两个jar文件：docx-replace-1.0-SNAPSHOT.jar， original-docx-replace-1.0-SNAPSHOT.jar。
2. 使用docx-replace-1.0-SNAPSHOT.jar进行文档替换，在命令行下切换到此jar文件所在的目录，同时此目录下需要一个名为config.txt文件，这是程序所需要的配置文件。
3. config.txt文件里有且只有两行：第一行是替换规则文件所在的完整路径；第二行是替换文件所在的文件夹的完整路径（里面也可以包含文件夹）。config.txt文件内容如下（可参考config.txt)：
    Linux:
    &nbsp;&nbsp;/home/wyz/Desktop/word/替换规则.txt
    &nbsp;&nbsp;/home/wyz/Desktop/word/new
    Windows:
    &nbsp;&nbsp;F:\桌面\wyz\替换规则.txt
    &nbsp;&nbsp;F:\桌面\wyz\程序及表单<br/>
    __注__：替换规则文件和需要替换的文件在files文件夹下。

4. 在jar文件所在的目录下运行：java -jar docx-replace-1.0-SNAPSHOT.jar。 __注意__：程序只能处理docx文件，其他后缀的文件自动跳过。
5. 会在替换的文件夹同级目录下生成一个新的文件夹：“原文件夹名-替换后”，就是替换后的文件。

__注__：也可以直接下载docx-replace-1.0-SNAPSHOT.jar文件，更新程序的同时也会更这个jar文件。


###程序说明：
<p>&nbsp;&nbsp;读取docx文件使用的是POI库，它是对文件进行段落切分，利用最小的‘run’进行查找和替换，所以你要替换的文字需要使用特殊字符包裹起来，如：‘[ ]’,‘{ }’等，
才能让它们在同一个run中，否则替换不成功。<br/>&nbsp;&nbsp;程序可以处理段落，表格，页眉的段落，页眉的表格的文本替换。</p>


