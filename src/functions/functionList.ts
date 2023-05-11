/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-11-11 11:36:04
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2023-04-03 11:35:52
 * @Description: 导入数据
 */
import axios from './axios'
import { getTime } from './functions'
import { transformToStr,strToNamber,reg } from './../taskpane/unit'
//公用函数
const PAGESIZE: number = 500;//最大条数
let CONCURRENTSIZE: number = 5;//最大并发数
/**
 * @description: 迭代实现
 * @param {Object} result 第一个值的输出结果
 * @param {string} tableName 请求表明
 * @param {string} codesList string[][] code集合
 * @param {string} columnsStr 请求入参
 * @param {string} startDate 
 * @param {string} endDate
 * @param {string} dateList 日期列表
 * @param {any[]} indexArr 第几个0,插入地址，1串行一共是多少条，2第几个并发，3第几个请求
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @return {*}
 */
async function* creatiterator(result:Object,tableName:string,codesList:string[][],columnsStr:string[],startDate:string,endDate:string,dateList:string[],indexArr:any[],nullType:number,axis:number) {
    for (let i = 0; i < codesList.length; i++) {
        const codes = codesList[i];
        yield await getData(result,tableName,codes,columnsStr,startDate,endDate,dateList,[...indexArr,i],nullType,axis);
    }
  }
/**
 * @description: 请求数据
 * @param {Object} result 第一个值的输出结果
 * @param {string} tableName 表明
 * @param {string} codes codesList
 * @param {string[]} columnsStr 入参
 * @param {string} startDate 
 * @param {string} endDate
 * @param {string} dateList 日期列表
 * @param {any[]} indexArr 第几个0,插入地址，1串行一共是多少条，2第几个并发，3第几个请求
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @return {*}
 */
function getData(result:Object,tableName:string,codes:string[],columnsStr:string[],startDate:string,endDate:string,dateList:string[],indexArr:any[],nullType:number,axis:number){
    const extendsParams = {
        codes,//股票代码
        columns:columnsStr.toString()//因子
      }
    const params = {
        tableName,//表名
        extendsParams: JSON.stringify(extendsParams),
        startDate,
        endDate
    }
    return new Promise((resolve,reject)=>{
        axios.post('/',params).then((res:any)=>{
            //设置地址
            const addrStart:string = indexArr[0]//鼠标位置坐标
            const numIndex:number = indexArr[1]//一个串行里有几条
            const sIndex:number = indexArr[2]//第几个串行
            const iIndex:number = indexArr[3]//每串中的第几个

            // 例子A1
            const addrStr = (addrStart.match(reg) as any)[0];//拿到地址的字母 A
            const addrNum:number =Number(addrStart.replace(/[^\d.]/g, ''))//拿到地址的数字 1 
            const endNumer = strToNamber(addrStr) + dateList.length - 1;//拿到矩阵结尾的横坐标（此时为数字需要转换成AAB形式）
            let startNum:number = 0//矩阵结尾纵坐标数字
            let endNum:number = 0//矩阵结尾纵坐标数字
            console.log(sIndex);
            
            if(sIndex < CONCURRENTSIZE - 1 && codes.length == PAGESIZE ){
                startNum = numIndex*sIndex*PAGESIZE + iIndex * PAGESIZE 
                endNum = numIndex*sIndex*PAGESIZE + (iIndex + 1) * PAGESIZE - 1
            }else{
                if(codes.length == PAGESIZE){
                    startNum = numIndex*(CONCURRENTSIZE - 1)*PAGESIZE + iIndex * PAGESIZE
                    endNum = numIndex*(CONCURRENTSIZE - 1)*PAGESIZE + (iIndex + 1) * PAGESIZE - 1
                }else{
                    startNum = numIndex*(CONCURRENTSIZE - 1)*PAGESIZE + iIndex * PAGESIZE
                    endNum = numIndex*(CONCURRENTSIZE - 1)*PAGESIZE + iIndex * PAGESIZE + codes.length - 1
                }
            }
            let addrRange:string = ''
            if(axis == 1){
                addrRange = `${addrStr}${addrNum + startNum}:${transformToStr(endNumer)}${addrNum + endNum }`//矩阵开始结尾横纵坐标
            }else{
                addrRange = `${transformToStr(strToNamber(addrStr) + startNum)}${addrNum}:${transformToStr(strToNamber(addrStr) + endNum)}${addrNum + dateList.length - 1}`//矩阵开始结尾横纵坐标
            }
            
            console.log(addrRange);
            //地址end
            setResult(result,res.data.data,codes,dateList,addrRange,columnsStr,sIndex === 0 && iIndex === 0,nullType,axis)
            resolve('success')
        }).catch((err:any)=>{
            reject(err)
        })        
    })

}
/**
 * @description: 请求数据
 * @param {object} result 第一个值的输出结果
 * @param {string} tableName 表明
 * @param {string} codes codesList
 * @param {string} columnsStr 入参
 * @param {string} startDate 
 * @param {string} endDate
 * @param {string} addr 插入地址
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @return {*}
 */
 function getAllData(result:{val:string},tableName:string,codes:string[],columnsStr:string[],startDate:string,endDate:string,dateList:string[],addr:string,nullType:number,axis:number){
    if(codes.length < PAGESIZE){
        CONCURRENTSIZE = 1
    }else if(codes.length < PAGESIZE * CONCURRENTSIZE ){
        CONCURRENTSIZE = Math.ceil(codes.length / PAGESIZE)
    }
    const size = Math.ceil(codes.length / PAGESIZE)//切割数组
    let newArr: string[][] = []
    for (let i = 0; i < size; i++) {
      newArr[i] = []
      newArr[i] = codes.slice(i * PAGESIZE, (i + 1) * PAGESIZE)
    }
  //串行请求功能实现
  const axiosSize = Math.ceil(newArr.length / CONCURRENTSIZE)
  let concurrentArr: any[][][] = [] //[[[],[]],[[],[]]]
  for (let i = 0; i < CONCURRENTSIZE; i++) {
    concurrentArr[i] = []
    concurrentArr[i] = newArr.slice(i * axiosSize, (i + 1) * axiosSize)
  }
  console.log(concurrentArr);
  concurrentArr.forEach((codesList,index) => {
    if(codesList.length == 0 ) return
    const gin = creatiterator(result,tableName,codesList, columnsStr,startDate,endDate,dateList,[addr,axiosSize,index],nullType,axis)
    codesList.map(() => {
      gin.next()
    })
  })
  return new Promise((resolve,reject)=>{
    const timer = setInterval(()=>{
        if(result.val){
            clearInterval(timer)
            resolve(result.val)
        }
    },500)
  })
}
/**
 * @description: 分部设置结果
 * @param {Object} result 第一个值的输出结果
 * @param {any} webData web请求数据
 * @param {string} codesList 输入codes集合
 * @param {string} dateList 时间集合
 * @param {string} addrRange 插入地址
 * @param {string[]} columnsStr 参数
 * @param {boolean} isFirst 是否第一个
* @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @return {*}
 */
async function setResult(result:any,webData:any[],codesList:string[],dateList:string[],addrRange:string,columnsStr:string[],isFirst:boolean,nullType:number,axis:number){
    const dataMap = new Map()

    webData.map(item=>{
        if(axis == 1){
            dataMap.set(item[columnsStr[0]]+','+item[columnsStr[1]],item[columnsStr[2]])
        }else{
            dataMap.set(item[columnsStr[1]]+','+item[columnsStr[0]],item[columnsStr[2]])
        }
        
    })
    console.log(dataMap);
    let setData:any[][] = []
    if(axis == 1){
        for (let i = 0; i < codesList.length; i++) {
            setData[i] = []
            for (let j = 0; j < dateList.length; j++) {
                if(i === 0 && j === 0 && isFirst){
                    setData[i][j] = null;
                    result.val = dataMap.get(codesList[i]+','+dateList[j])||'NaN'
                }else{
                    if(nullType == 1){
                        setData[i][j] = dataMap.get(codesList[i]+','+dateList[j])||'NaN' 
                    }else{
                        if(dataMap.has(codesList[i]+','+dateList[j])){
                            setData[i][j] = dataMap.get(codesList[i]+','+dateList[j])
                        }else{
                            const num = dateList.indexOf(dateList[j])
                            const timeListSort = num > -1 ? dateList.slice(0,num):[]
                            const timeList = timeListSort.reverse()
                            let btn = false
                            for (let time = 0; time < timeList.length; time++) {
                                const timeObj = timeList[time];
                                console.log(timeObj);
                                if(dataMap.has(codesList[i]+','+timeObj)){
                                    console.log(codesList[i],timeObj);
                                    setData[i][j] = dataMap.get(codesList[i]+','+timeObj)
                                    btn = true
                                    break;
                                }
                            } 
                            if(!btn){
                                setData[i][j] = 'NaN'
                            }                        
                        }
                    }
                }
            }
        }        
    }else{
        for (let i = 0; i < dateList.length; i++) {
            setData[i] = []
            for (let j = 0; j < codesList.length; j++) {
                if(i === 0 && j === 0 && isFirst){
                    setData[i][j] = null;
                    result.val = dataMap.get(dateList[i]+','+codesList[j])||'NaN'
                }else{
                    if(nullType == 1){
                        setData[i][j] = dataMap.get(dateList[i]+','+codesList[j])||'NaN' 
                    }else{
                        if(dataMap.has(dateList[i]+','+codesList[j])){
                            setData[i][j] = dataMap.get(dateList[i]+','+codesList[j])
                        }else{
                            const num = dateList.indexOf(dateList[i])
                            const timeListSort = num > -1 ? dateList.slice(0,num):[]
                            const timeList = timeListSort.reverse()
                            let btn = false //是否找到有值
                            for (let time = 0; time < timeList.length; time++) {
                                const timeObj = timeList[time];
                                console.log(timeObj);
                                if(dataMap.has(timeObj+','+ codesList[j])){
                                    setData[i][j] = dataMap.get(timeObj+','+ codesList[j])
                                    btn = true 
                                    break;
                                }
                            }       
                            if(!btn){
                                setData[i][j] = 'NaN'
                            }                 
                        }
                    }
                }
            }
        }  
    }

    console.log(setData);
    //设置地址
    const context = new Excel.RequestContext();
    let sheet = context.workbook.worksheets.getActiveWorksheet()
    let range = sheet.getRange(addrRange);
    range.values = setData;
    //创建提示
    if(isFirst){
        let rangePrompt = sheet.getRange(addrRange.split(':')[0]);
        rangePrompt.format.fill.color = "#ffeb3b";
        rangePrompt.dataValidation.prompt = {
            message: "Excel插件—中诚信指数—函数编辑",
            showPrompt: true, // The default is 'false'.
            title: "如需修改请使用"
        };       
    }
    //end
    await context.sync();

}
/**
 * @description: 设置时间
 * @param {string} codes
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {string} addr
 * @param {number} axis
 * @return {*}
 */
const setFunPubilcData = (codes:string[][],time:string,addr:string|undefined,axis:number)=>{
    return new Promise(async (resolve,reject)=>{
        if(typeof codes != 'object') return reject({code:0,msg:'codes输入不正确'})
        if(typeof time != 'string') return reject({code:0,msg:'time输入不正确'})
        const timeStr = time.split(',')
        if(timeStr.length != 5) return reject({code:0,msg:'time输入不正确'})
        //获取输入地址
        const getAddr:string|undefined = addr
        const addrStart:string = getAddr && getAddr.split('!')[1] || 'A1'
        console.log(addrStart);
        //end
        const codesListStr:string = codes.toString();//将获取的数据转成字符串 code
        const codeList = codesListStr.split(',')//转换成一维数组
        /**日期处理**/
        const dataFormat:number = Number(timeStr[4])
        let startDate = timeStr[0]
        let endDate = timeStr[1]
        if(dataFormat === 1){
            startDate = timeStr[0]
            endDate = timeStr[1]
        }else if(dataFormat === 2){
            startDate = timeStr[0].replace(/\//g,'')
            endDate = timeStr[1].replace(/\//g,'')
        }else if(dataFormat === 3){
            startDate = timeStr[0].replace(/-/g,'')
            endDate = timeStr[1].replace(/-/g,'')
        }
        const timeList =await getTime(startDate,endDate,timeStr[2],timeStr[3])
        
        let dateListReverse:string[] = timeList.data.data.map((item:any)=>item.date).reverse()
        if(dateListReverse.length == 0) return reject({code:0,msg:'时间获取为空'})
        let dateList:string[] = []
        if(dataFormat == 1){
            dateList =dateListReverse.map((item:any)=>item.replace(/-/g,''));
        }else if(dataFormat == 2){
            dateList =dateListReverse.map((item:any)=>item.replace(/-/g,'/'));
        }else if(dataFormat == 3){
            dateList =dateListReverse
        }
        /**日期处理end**/
        let setData:string[][] = []
        //设置地址
        const addrStr = (addrStart.match(reg) as any)[0];//拿到地址的字母A1
        const addrNum:number =Number(addrStart.replace(/[^\d.]/g, ''))//拿到地址的数字 A1 
        console.log(addrNum);
        let addrRange:string = ''
        if(axis == 1){
            if(addrNum < 2) return '输入位置不能选第一行'
            const endNumer = strToNamber(addrStr) + dateList.length - 1;//拿到矩阵结尾的横坐标（此时为数字需要转换成AAB形式）
            addrRange = `${addrStr}${addrNum - 1}:${transformToStr(endNumer)}${addrNum - 1}`//矩阵开始结尾横纵坐标 
            setData = [dateList]           
        }else{
            if(addrStr == 'A') return '输入位置不能选第一列'
            const endNumer = addrNum + dateList.length - 1;//拿到矩阵结尾的横坐标（此时为数字需要转换成AAB形式）
            addrRange = `${transformToStr(strToNamber(addrStr) - 1)}${addrNum}:${transformToStr(strToNamber(addrStr) - 1)}${endNumer}`//矩阵开始结尾横纵坐标  
            dateList.map(item=>{
                setData.push([item])
            })     
        }  

        console.log(addrRange);
        const context = new Excel.RequestContext();
        let sheet = context.workbook.worksheets.getActiveWorksheet()
        let range = sheet.getRange(addrRange);
        range.values = setData;
        range.format.autofitColumns()
        await context.sync();
        resolve({code:200,msg:'成功！',obj:{codeList,dateList,dateListReverse,addrStart}})
    })
}
//公用函数end
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} invocation Custom function invocation
 * @return {string}
 */
export async function exposureList(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'exposuresw',codeList,['SECU_CODE','TRADINGDAY',functionArg],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function specificReturnList(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'specific_ret_sw',codeList,['SECU_CODE','TRADINGDAY','SPRET'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function stock_style_vg_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_stock_style_class',codeList,['secu_code','dt','style_class'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function stock_style_cap_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_stock_style_class',codeList,['secu_code','dt','size_class'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 股票成长价值风格得分
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function stock_style_score_vg_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_stock_style_class',codeList,['secu_code','dt','style_score'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 股票市值风格得分
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function stock_style_score_cap_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_stock_style_class',codeList,['secu_code','dt','size_score'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/*****************基金************************/
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function fund_type_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'fund_type_ccx',codeList,['FUNDCODE','TRADE_DT','FUND_TYPE_NAME'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function fund_style_vg_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_style_class_'+functionArg,codeList,['fund_code','dt','fund_style_score_r6_class'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function fund_style_cap_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_style_class_'+functionArg,codeList,['fund_code','dt','fund_size_score_class'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function stock_specific_risk_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'d_srisk_vra',codeList,['SECU_CODE','TRADE_DATE','SRISK'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function fund_exposure_style_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'fund_barra_exposure_sw',codeList,['FUNDCODE','TRADE_DT',functionArg],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}

/**
 * @description: 基金市值风格得分
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function fund_style_score_cap_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_style_class_'+functionArg,codeList,['fund_code','dt','fund_size_score'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金市值风格得分
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function fund_style_score_vg_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_style_class_'+functionArg,codeList,['fund_code','dt','fund_style_score_r6'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金久期
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_dur_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_duration_type',codeList,['fund_code','dt','dur_mean_rolling_'+functionArg],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金久期标签
 * @param {string} codes 股票code集合
 * @param {string} scrollNumber 滚动期数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_dur_label_List(codes:string[][],scrollNumber:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_duration_type',codeList,['fund_code','dt','fund_dur_type_rolling_'+scrollNumber],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金持有的第N大行业名称
 * @param {string} codes 股票code集合
 * @param {string} scrollNumber 滚动期数
 * @param {string} industrySort 行业排名
 * @param {string} industryType 行业类型
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_hold_top_industry_name_List(codes:string[][],scrollNumber:number,industrySort:number,industryType:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const outParams:string = (scrollNumber == 1?
        `top${industrySort}_industry_name_${industryType}`:
        `top${industrySort}_industry_name_${industryType}_rolling_${scrollNumber}`)
    const showResult =await getAllData(result,'open_ccx_fund_industry_top_weight_rolling_'+industryType,codeList,['fund_code','dt',outParams],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金持有的第N大行业名称
 * @param {string} codes 股票code集合
 * @param {string} scrollNumber 滚动期数
 * @param {string} industrySort 行业排名
 * @param {string} industryType 行业类型
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_hold_top_industry_ratio_List(codes:string[][],scrollNumber:number,industrySort:number,industryType:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const outParams:string = (scrollNumber == 1?
        `top${industrySort}_industry_weight_${industryType}`:
        `top${industrySort}_industry_weight_${industryType}_rolling_${scrollNumber}`)
    const showResult =await getAllData(result,'open_ccx_fund_industry_top_weight_rolling_'+industryType,codeList,['fund_code','dt',outParams],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金行业标签
 * @param {string} codes 股票code集合
 * @param {string} industryType 滚动期数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_industry_label_List(codes:string[][],industryType:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_industry_label_'+industryType,codeList,['fund_code','dt','industry_label_name'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金打新收益
 * @param {string} codes 股票code集合
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_hold_newstock_ret_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_new_stock_data',codeList,['fund_code','dt','hold_new_stock_ret'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/**
 * @description: 基金经理卸任标签
 * @param {string} codes 股票code集合
 * @param {string} functionArg 函数参数
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
export async function ccx_fund_manager_dimision_List(codes:string[][],functionArg:string,time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fund_manager_change_dimissiondate',codeList,['fund_code','dt','change_'+functionArg],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}
/***********************指数****************************/
/**
 * @description: 因子暴露表
 * @param {string} codes 股票code集合
 * @param {string} time 时间start,end,dataType,frequency,dataFormat/开始，结束，日历日或交易日，日期频率，时间格式
 * @param {number} nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {string} addr Custom function invocation
 * @return {string}
 */
 export async function index_levels_List(codes:string[][],time:string,nullType:number,axis:number,addr:string|undefined){    
    const obj:any = await setFunPubilcData(codes,time,addr,axis)
    if(obj.code === 0 ) return obj.msg
    const { codeList,dateList,dateListReverse,addrStart } = obj.obj
    let result = { val:'' }
    const showResult =await getAllData(result,'open_ccx_fof_index_levels',codeList,['index_code','trade_date','index_level'],dateList[0],dateList[dateList.length-1],dateListReverse,addrStart,nullType,axis)
    return showResult
}