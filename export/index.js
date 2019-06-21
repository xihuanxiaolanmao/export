// 云函数入口文件
const cloud = require('wx-server-sdk')
const nodeExcel = require('excel-export')
const fs = require('fs')
const path = require('path')

cloud.init({
  env: "com-2ovdl"   // 你的环境
})

const db = cloud.database()
const MAX_LIMIT = 80
// 生成分数项并且下载对应的excel
exports.main = async (event, context) => {

  let collectionId = 'zl'                 // 模拟的集合名
  let openId = 'sda6248daa888764'             // 模拟openid
  let confParams = ['单价','日工资', '部门', '工作内容','备注','数量','项目','员工名','时间'] // 模拟表头

  let jsonData = []
  
  let time=event
  //db.collection(collectionId).count
  const countResult = await db.collection('zl').where({time:event.time}).count()
  const total = countResult.total
  const ci = Math.ceil(total / MAX_LIMIT);
  
  // 获取数据
  
  for (let i = 0; i < ci; i++) {
    await db.collection(collectionId).where({ time: event.time }).orderBy('staffName', 'desc').skip(i).limit(MAX_LIMIT).get().then(res => {
      if(i!=0){
       jsonData=jsonData.concat(res.data)
      }else{
        jsonData=res.data
      }
    })
  }
  
  // 等待所有
  /*return (await Promise.all(jsonData)).reduce((acc, cur) => {
    return {
      data: jsonData,
      errMsg: acc.errMsg,
      ci:ci,
      total:total,
      countResult: countResult,
      time:time
    }
  })*/

  // 转换成excel流数据
  let conf = {
    stylesXmlFile: path.resolve(__dirname, 'styles.xml'),
    name: 'sheet',
    cols: confParams.map(param => {
      return { caption: param, type: 'string' }
    }),
    rows: jsonToArray(jsonData)
  }
  let result = nodeExcel.execute(conf) // result为excel二进制数据流

  // 上传到云存储
  return await cloud.uploadFile({
    cloudPath: `download/sheet${openId}.xlsx`,  // excel文件名称及路径，即云存储中的路径
    fileContent: Buffer.from(result.toString(), 'binary'),
  })
  
  // json对象转换成数组填充
  function jsonToArray(arrData) {
    let arr = new Array()
    arrData.forEach(item => {
      let itemArray = new Array()
      for (let key in item) {
        if (key === '_id' || key === '_openid'||key==='id') { continue }
        itemArray.push(item[key])
      }
      arr.push(itemArray)
    })
    return arr
  }

}


