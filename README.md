# pickExcel


基于某老哥([excel](https://github.com/wwhgtt/excel))已有成果改进的浏览器内javascript提取excel数据的工具，excel→json
<br>

<br>

### 目前已提供的功能
=
+ 抓取整列数据
+ 抓取整行数据
+ 抓取**符合某些条件**的整行数据
<br>
<br>

### 使用方法
拷贝项目至本地后，在浏览器内打开index.html
<br>
<br>
### 使用配置
所有涉及到数字的地方请填写阿拉伯数字!<br>
所有涉及到数字的地方请填写阿拉伯数字!<br>
所有涉及到数字的地方请填写阿拉伯数字!<br>


无论使用哪种功能，请先填写需要抓取Excel表格的页码<br>默认的命名方式为"sheet1,sheet2,sheet3..."<br>不排除有些表格自定义了命名，所以请提供页码名称,默认为"sheet1"
<br>
#### 抓取整列数据
选择"导出整列"选项<br>
填写需要导出的列数以及每一列的名称,例如:<br><br>
!["demo"](https://github.com/wyh369352887/pickExcel/raw/master/image/2.jpg)
<br><br>
将xlxs文件拖进虚线框内<br><br>
![虚线框](https://github.com/wyh369352887/pickExcel/raw/master/image/1.jpg)
<br><br>

#### 抓取整行数据
选择"导出整行"选项<br>
填写excel数据起始的行数(_有些表格的会有总表头或者其他的内容会影响数据的读取_),填写抓取的起止行数,例如:
![demo](https://github.com/wyh369352887/pickExcel/raw/master/image/3.jpg)
<br>
<br>
在这个表格中，第一行为表头，实际数据是从第二行开始的,所以起始行数为2
<br>
<br>
![demo](https://github.com/wyh369352887/pickExcel/raw/master/image/4.jpg)
<br>
<br>

#### 抓取符合条件的整行
选择"导出符合条件的整行",填写需要满足几个条件(注意，每个条件之间的关系是"且&&"),接着输入每个需要限制的条目和值,例如:
<br>
<br>
![demo](https://github.com/wyh369352887/pickExcel/raw/master/image/5.jpg)
<br>
<br>

这个项目只是个初版，是在工作之余根据以往的惨痛教训编写的，为了加快工作效率，暂时只支持这三个功能，后面陆续会增加，比如符合条件的行的其他值？抓取不连续的行的数据等等，个人感觉有些功能实现可能并没有什么实际意义...  <br>样式也会优化...