# export-docx

> 浏览器导出docx文件。利用有你编写文档格式的模版 temp.docx，在其中插入变量名，待执行导出时赋值。

**感谢**: [docxtemplater](https://github.com/open-xml-templating/docxtemplater)

## Use

### 模版文件

> temp.docx 移步[github](https://github.com/unzoa/export-docx)获取。

**注意变量绑定**：
- 图片：{%lineViewDataImg}
- 字符串：{lineViewDataTitle}

### 安装

```bash
npm i export-docx -D
```

### 代码使用

```js
import exportWord from 'export-docx'

// 字段参照temp.docx, 可自定义匹配编辑
const data = {
  "title": "Report Title: 2021-12-18 20:00:00-->2021-12-18 21:00:00",
  "totalNum": 15915,
  "highLevelCount": 20,
  "middleLevelCount": 4,
  "lowLevelCount": 374,
  "outsideAttackCount": "aaa: 15726 | bbb: 369 | ccc: 140 | ddd: 20 | total: 16255",
  "lineViewDataImg": "data:image/png;base64, ···", // base64 格式
  "lineViewDataTitle": "图表：aaa"
}

/**
* [导出word]
* @param  {String} tempDocUrl 你的模版文件地址
* @param  {Object} docData    你的导出数据
* @param  {String} fileName   你的导出文件名, 默认值：sample.docx
* @param  {[Number, Number]}  imgSize   统一设置图片尺寸， [width, height]
*/
exportWord('/temp.docx', data, 'test.docx', [520, 520 * 272 / 609])
```
