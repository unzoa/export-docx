# export-docx

> 浏览器导出docx文件，利用实现准备好的模版docx文件

## 导出image：
 * 1. 利用 open-docxtemplater-image-module 编辑imageModule函数
 * 2. 在new Docxtemplater()实例上利用api：attachModule加载ImageModule实例
 * 3. doc.setData中绑定的图片字段应该是base64的图片
 * 4. temp.docx 中图片的字段区别与字符串{variable} 是 {%imageVariable}
 * If you would like to choose which images should be centered one by one:
   - Set the global switch to false opts.centered = false.
   - Use {%image} for images that shouldn't be centered.
   - Use {%%image} for images that you would like to see centered.
 *