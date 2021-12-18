const JSZipUtils = require('jszip-utils')
const Docxtemplater = require('docxtemplater')
const { saveAs } = require('file-saver')
const JSZip = require('jszip')
const ImageModule = require('open-docxtemplater-image-module')

/**
* [导出word]
* @param  {String} tempDocUrl [模版文件地址]
* @param  {Object} docData    [导出数据]
* @param  {String} fileName   [导出文件名]
* @param  {Array}  imgSize    [统一设置图片尺寸]
*/

module.exports = (tempDocUrl, docData, fileName = 'sample Word', imgSize) => {
  // 读取并获得模板文件的二进制内容
  JSZipUtils.getBinaryContent(tempDocUrl, function (error, content) {
    // input.docx是模板。我们在导出的时候，会根据此模板来导出对应的数据
    // 抛出异常
    if (error) {
      throw error
    }

    // 创建一个JSZip实例，内容为模板的内容
    let zip = new JSZip(content)
    // 创建并加载docxtemplater实例对象
    let doc = new Docxtemplater()
      .loadZip(zip)
      .attachModule(imageModule(imgSize))

    // 设置模板变量的值
    doc.setData(docData)

    try {
      // 用模板变量的值替换所有模板变量
      doc.render()
    } catch (error) {
      // 抛出异常
      let e = {
        message: error.message,
        name: error.name,
        stack: error.stack,
        properties: error.properties
      }
      console.log(e)
      throw error
    }

    // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
    let out = doc.getZip().generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    })
    // 将目标文件对象保存为目标类型的文件，并命名
    saveAs(out, fileName + '.docx')
  })
}

function imageModule (size = [200, 200]) {
  let opts = {}
  opts.centered = true // 图片居中，在word模板中定义方式为{%%image}
  opts.fileType = 'docx'

  /**
  * [description]
  * @param  {[type]} tagValue [base64 或者 图片的url，但是这里对url不做支持]
  * @param  {[type]} tagName  [docData中图片的key]
  * @return {[type]}          [buffer]
  */
  opts.getImage = (tagValue, tagName) => {
    return base64DataURLToArrayBuffer(tagValue)
  }

  /**
  * [getSize description]
  * @param  {[type]} img      [buffer 对象]
  * @param  {[type]} tagValue [base64 或者 图片的url，但是这里对url不做支持]
  * @param  {[type]} tagName  [docData中图片的key]
  * @return {[type]}          [图片width, height]
  */
  opts.getSize = (img, tagValue, tagName) => {
    return size
  }

  return new ImageModule(opts)
}

// 导出图片，格式转换
function base64DataURLToArrayBuffer (dataURL) {
  const base64Regex = /^data:image\/(png|jpg|svg|svg\+xml);base64,/
  if (!base64Regex.test(dataURL)) {
    return false
  }
  const stringBase64 = dataURL.replace(base64Regex, '')
  let binaryString
  if (typeof window !== 'undefined') {
    binaryString = window.atob(stringBase64)
  } else {
    binaryString = Buffer.from(stringBase64, 'base64').toString('binary')
  }
  const len = binaryString.length
  const bytes = new Uint8Array(len)
  for (let i = 0; i < len; i++) {
    const ascii = binaryString.charCodeAt(i)
    bytes[i] = ascii
  }
  return bytes.buffer
}
