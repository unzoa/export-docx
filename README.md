# export-docx

> 浏览器导出docx文件，利用实现准备好的模版 temp.docx 文件

## Use

```js
import exportWord from 'export-docx'

// 字段参照temp.docx, 可自定义匹配编辑
const data = {
  "title": "Report Title: 2021-12-18 20:00:00-->2021-12-18 21:00:00",
  "totalNum": 15915,
  "highLevelCount": 20,
  "middleLevelCount": 4,
  "lowLevelCount": 374,
  "outsideSourceCount": 397,
  "outsideNum": 15915,
  "countryCount": 23,
  "insideSourceCount": 0,
  "insideNum": 0,
  "outsideAttackCount": "aaa: 15726 | bbb: 369 | ccc: 140 | ddd: 20 | total: 16255",
  "lineViewDataImg": "data:image/png;base64, ···", // base64 格式
  "lineViewDataTitle": "图表：aaa",
  "pieViewDataImg": "",
  "pieViewDataTitle": "图表：bbb",
  "pie2ViewDataImg": "",
  "pie2ViewDataTitle": "图表：ccc",
  "srcPortDataImg": "",
  "srcPortDataTitle": "图表：ddd",
  "barViewDataImg": "",
  "barViewDataTitle": "图表：eee",
  "insideAttackCount": "fff",
  "inLineViewDataImg": "",
  "inLineViewDataTitle": "无数据",
  "inPieViewDataImg": "",
  "inPieViewDataTitle": "无数据",
  "inPie2ViewDataImg": "",
  "inPie2ViewDataTitle": "无数据",
  "inSrcPortDataImg": "",
  "inSrcPortDataTitle": "无数据"
}

exportWord('/temp.docx', data, 'test.docx', [520, 520 * 272 / 609])
```
