# pickExcel
基于某老哥(https://github.com/wwhgtt/excel)
已有成果改进的纯js提取excel数据的小工具，将excel数据转化为json数据
<br/>
<br/>
<br/>
<h2>目前已提供的功能</h2>
<ul>
  <li>抓取整列数据(多列)</li>
  <li>抓取整行数据(连续多行)</li>
  <li>抓取符合某些条件的整行数据</li>
</ul>
<br/>
<h2>使用方法</h2>
在浏览器内打开index.html<br/>
<h2>使用配置</h2>
所有涉及到数字的地方请填写阿拉伯数字!<br/>
所有涉及到数字的地方请填写阿拉伯数字!<br/>
所有涉及到数字的地方请填写阿拉伯数字!<br/>
<br/>
<br/>
<br/>
<p>无论使用哪种功能，请先填写需要抓取Excel表格的页码，默认的命名方式为Sheet1,Sheet2,Sheet3...<br/>
不排除有些表格自定义了命名，所以请提供页的名称<br/>
默认为Sheet1</p>
<br/>
<h4>抓取整列数据</h4>
<p>选择“导出整列”选项，填写需要导出的列数以及每一列的名称,将excel拖进虚线框内</p>
<p>例如，导出3列，每一列的名字分别为“姓名”，“性别”，“年龄”</p>
<br/>
<h4>抓取整行数据</h4>
<p>选择“导出整行”选项，填写excel数据起始的行数(有些表格的会有总表头或者其他的内容会影响数据的读取),填写抓取的起止行数</p>
<p>例如，表格从第2行开始,抓取从第20到第30行的数据</p>
<br/>
<h4>抓取符合条件的整行</h4>
<p>选择“导出符合条件的整行”选项,填写需要满足几个条件(注意，每个限制之间的要求是“且” &&)，接着输入每个需要限制的条目和值</p>
<p>例如，需要满足2个条件，分别是“年龄”为“20”和“性别”为“男”</p>
<br/>
<br/>
<br/>
<p>这个项目只是个初版，是在工作之余根据以往的惨痛教训编写的，为了加快工作效率，暂时只支持这三个功能，后面陆续会增加，比如符合条件的行的其他值？抓取不连续的行的数据等等，个人感觉有些功能实现可能并没有什么实际意义...</p>
<p>样式也会优化...</p>
