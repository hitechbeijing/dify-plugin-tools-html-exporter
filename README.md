# html_exporter
由bowenliang123的md_exporter实现了markdown转docx，我在他的基础上重写方法，实现html转docx
<p><img width="1366" alt="image" src="https://github.com/user-attachments/assets/10380232-f855-4b1d-9dd9-4e0f5e9df9aa" /></p>


<h3>支持以下HTML标签的转换</h3>

```
<p></p>
<div></div>
<span></span>
<ul></ul>
<li></li>
<table></table>
<br\>
<h1></h1>
<h2></h2>
<h3></h3>
<h4></h4>
<h5></h6>
<h6></h6>
<span></span>
<b></b>
<strong></strong>
<i></i>
<u></u>
<em></em>
<u></u>
<small></small>
<ol></ol>
<mark></mark>
```
<h3>支持css样式：</h3>

```
front-size
front-color
font-weight
text_align
margin
line-height
text-decoration
background-color
```
<h3>支持完整141色HTML颜色名称 详见：</h3>
<a href='https://www.runoob.com/tags/html-colorname.html'>HTML颜色名</a>
<h3>注意：LLM 生成的HTML需遵循以下规范：</h3>
<ul>
  <li>1、文档标题请使用h1-h6的HTML标题标签</li>
  <li>2、文档段落请使用html p标签</li>
  <li>3、span、mark标签需放置在块级标签内</li>
  <li>4、不支持html style标签，css样式请使用html标签style属性</li>
  <li>5、不支持javascript</li>
</ul>
