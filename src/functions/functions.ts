/*
/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-08-24 16:32:45
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2023-04-04 09:10:41
 * @Description: excel自定义函数
 */
/* global clearInterval, console, CustomFunctions, setInterval */
import axios from './axios'
import SaveDataMap from './SaveDataMap'
import {
  index_rebalance_frequencyList,
  index_weight_methodList,
  formatDate,transformToStr,
  strToNamber,
  reg
} from '../taskpane/unit'
import { exposureList,
  specificReturnList,
  stock_style_vg_List,
  stock_style_cap_List,
  stock_style_score_vg_List,
  stock_style_score_cap_List,
  fund_type_List,
  fund_style_vg_List,
  fund_style_cap_List,
  fund_exposure_style_List,
  fund_style_score_cap_List,
  fund_style_score_vg_List,
  ccx_fund_dur_List,
  stock_specific_risk_List,
  index_levels_List,
  ccx_fund_dur_label_List,
  ccx_fund_hold_top_industry_name_List,
  ccx_fund_hold_top_industry_ratio_List,
  ccx_fund_industry_label_List,
  ccx_fund_hold_newstock_ret_List,
  ccx_fund_manager_dimision_List
 } from './functionList'
 import index_weight_List from './index_weight'
//公共配置
const PAGESIZE: number = 500;//最大条数
let CONCURRENTSIZE: number = 5;//最大并发数
//end
//获取交易日
/**
 * @description: 
 * @param {string} start 开始时间
 * @param {string} end  结束时间
 * @param {number} dataType 1自然日2交易日
 * @param {number} frequency 频率
 * @return {*}
 */
export const getTime = (start:string,end:string,dataType:number|string,frequency:number|string)=>{
  let params:any = {}
  let type:string = 'is_daily'
  console.log(frequency);
  
  switch (Number(frequency)) {
    case 1:
      type = 'is_daily'
      break;
    case 2:
      type = 'is_weekly'
      break;
    case 3:
      type = 'is_monthly'
      break;
    case 4:
      type = 'is_yearly'
      break;
    default:
      type = 'is_daily'
      break;
  }
  let obj:any = {}
  obj[type] = '1'
  if(dataType == 1){
    params = {
      tableName:'open_ccx_excel_calander_days',//自然日
      extendsParams: JSON.stringify(obj),
      startDate:start,
      endDate:end,
      startIndex:0,
      pageSize:200
    }
  }else{
    params = {
      tableName:'open_ccx_excel_stock_tradingdays',//交易日
      extendsParams: JSON.stringify(obj),
      startDate:start,
      endDate:end,
      startIndex:0,
      pageSize:200
    }
  }
  
  return axios.post('/', params)
};
(window as any).getTime = getTime;
//获取股票代码
const getCodeListSZ = (startIndex=0)=>{
  const params = {
    tableName:'open_ccx_excel_stock_constants',//表名
    extendsParams: JSON.stringify({columns:"secu_code,chinameabbr,ashare_valid,ashare_with_delist,sh_market,sz_market,bj_market,main_board,second_board,stib_board,is_hs300,is_zz500,is_zz800,is_zz1000"}),
    startIndex,
    pageSize:200
  }
  return axios.post('/', params)
};
(window as any).getCodeListSZ = getCodeListSZ;
//获取基金代码
const getCodeListOF = (startIndex=0)=>{
  const params = {
    tableName:'open_ccx_excel_fund_constants',//表名
    extendsParams: JSON.stringify({columns:"fund_code,chinameabbr,fund_type_ccx_name,fund_valid_main"}),
    startIndex,
    pageSize:200
  }  
  return axios.post('/', params)
};
(window as any).getCodeListOF = getCodeListOF;
//获取中诚信指数
const getCodeListIndex = (startIndex=0)=>{
  const params = {
    tableName:'open_ccx_fof_index_basic',//表名
    extendsParams: JSON.stringify({columns:"index_code,short_name,label"}),
    startIndex,
    pageSize:200
  }  
  return axios.post('/', params)
};
(window as any).getCodeListIndex = getCodeListIndex;
//end
//公用函数
/**
 * @description: 
 * @param {any} ceilArr Excel插入的数据数组
 * @param {atring} tableName 表名
 * @param {atring} columnsStr 表列名
 * @return {*}
 */
function apiData(ceilArr: any[], tableName: string, columnsStr: string|object,callBack:Function) {
  const codesArr: string[] = ceilArr.map(item => {//股票
    return item.codes
  })
  const codeSet = new Set(codesArr)
  const nameArr: string[] = ceilArr.map(item => {//因子
    return item.factorName
  })
  const nameSet = new Set(nameArr)
  let extendsParams = {}
  if(typeof columnsStr === 'string'){
    console.log(codesArr);
    
    if(codesArr[0] == undefined){
      extendsParams = {
        columns: Array.from(nameSet).toString() + columnsStr//因子
      }
    }else{
      extendsParams = {
        codes: Array.from(codeSet),//股票代码
        columns: Array.from(nameSet).toString() + columnsStr//因子
      }
    }
    
  }else{
    extendsParams = columnsStr
  }
  //时间
  const startDate = ceilArr[0].date && formatDate(ceilArr[0].date)
  const endDate = ceilArr[ceilArr.length - 1].date && formatDate(ceilArr[ceilArr.length - 1].date)
  const params = {
    tableName,//表名
    extendsParams: JSON.stringify(extendsParams),
    startDate,
    endDate
  }
  // console.log(params);  
  return new Promise((resolve,reject)=>{
    try {
      axios.post('/', params).then((obj:any) => {
        if(!obj) return reject(obj)
        const data = obj.data
        if (data.code === 2000) {
          if (!data.data || data.data.length == 0) {
            // console.log(`未查询到${params.extendsParams}相关信息`)
            ceilArr.forEach((item: any) => {
              try {
                item.invocation.setResult('NaN')
              } catch (error) {
                console.log(error);
                
              }
              
            })
          } else {
            ceilArr.forEach((item: any) => {
              let btn = true
              const webArr = data.data;//请求得数据
              btn = callBack({btn,webArr,item})
              if (btn) {
                item.invocation.setResult('NaN')
              }
            })
          }
        } else {
          console.log(`服务器错误` + data.code)
        }
        resolve('success')
        
      }).catch((err:any) => {
        console.log(JSON.stringify(err))
        reject(err)
      })
  
    } catch (error) {
      reject(error)
      return JSON.stringify(error)
    }
  })

}
/**
 * @description: 
 * @param {any} arr Excel插入的数据数组
 * @param {atring} tableName 表名
 * @param {atring} columnsStr 表列名
 * @return {*}
 */
async function* creatiterator(arr: any[][], tableName: string, columnsStr: string|object,callBack:Function) {
  for (let item of arr) {
    yield await apiData(item, tableName, columnsStr,callBack);
  }
}
/**
 * @description: 
 * @param {any} allArr 集合数组
 * @param {string} tableName 数据库表名
 * @param {string} columnsStr 查询列名
 * @return {*}
 */
function box(allArr:any[],tableName:string,columnsStr:string|object,callBack:Function){
  if(!window.token.userName || !window.token.passWord) return window.alertErr('账号密码不存在，请重新登录！')
  //数组排序
  const allArrSort = allArr.sort(function(a,b){
    return a.date - b.date
  })
  const size = Math.ceil(allArrSort.length / PAGESIZE)//切割数组
  let newArr: any[][] = []
  for (let i = 0; i < size; i++) {
    newArr[i] = []
    newArr[i] = allArrSort.slice(i * PAGESIZE, (i + 1) * PAGESIZE)
  }
  //串行请求功能实现
  const axiosSize = Math.ceil(newArr.length / CONCURRENTSIZE)
  let concurrentArr: any[][] = [] //[[[],[]],[[],[]]]
  for (let i = 0; i < CONCURRENTSIZE; i++) {
    concurrentArr[i] = []
    concurrentArr[i] = newArr.slice(i * axiosSize, (i + 1) * axiosSize)
  }
  // console.log(concurrentArr);
  concurrentArr.forEach(arr => {
    const gin = creatiterator(arr, tableName, columnsStr,callBack)
    arr.forEach(() => {
      gin.next()
    })
  })
}
//end
let saveData = new SaveDataMap(10000*10*5)
////////////////////////////股票函数
//因子暴露表
interface ccx_stock_exposureArr {
  codes: string;
  date: string | number;
  factorName: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_exposureArr: ccx_stock_exposureArr[] = []
let ccx_stock_exposureTimer: any;//time用来控制事件的触发
/**
 * 该表反映了个股对各风格因子和行业因子的暴露值，即股票在各个风格因子和行业因子上的取值,其中行业因子为 0-1 变量，1 为该个股属于该行业分类，0 为不属于该行业。
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes string 股票代码
 * @param date Date 交易日期
 * @param factorName string 因子名称
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string} The sum of the two numbers.
 */
export function ccx_stock_exposure(codes: string, date: number | string, factorName: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`exposuresw${codes}${date}${factorName}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_exposureArr.push({ codes, date, factorName, invocation })
  if (ccx_stock_exposureTimer !== null) {
    clearTimeout(ccx_stock_exposureTimer);
  }
  ccx_stock_exposureTimer = setTimeout(() => {
    box(ccx_stock_exposureArr, 'exposuresw', ',SECU_CODE,TRADINGDAY',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for (let i = 0; i < webArr.length; i++) {
        if (item.codes == webArr[i].SECU_CODE && formatDate(item.date) == webArr[i].TRADINGDAY.replace(/-/g,'')) {
          item.invocation.setResult(webArr[i][item.factorName])
          btn = false;
          saveData.put(`exposuresw${item.codes}${item.date}${item.factorName}`,webArr[i][item.factorName])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.SECU_CODE == item.codes).reverse()
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0][item.factorName])
          saveData.put(`exposuresw${item.codes}${item.date}${item.factorName}`,filterCode[0][item.factorName]);
        }else{
          saveData.put(`exposuresw${item.codes}${item.date}${item.factorName}`,'NaN');
        }
      }
      return btn
    });
    ccx_stock_exposureArr = []
    ccx_stock_exposureTimer = null
  }, 500)
}
(window as any).saveData = saveData;
//end 因子暴露表

//因子收益表start
interface ccx_stock_factor_returnArr {
  codes:string;
  date: string | number;
  factorName: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_factor_returnArr: ccx_stock_factor_returnArr[] = []
let ccx_stock_factor_returnTimer: any;//time用来控制事件的触发
/**
 * 风格因子和行业因子对应的收益率
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param date Date交易日期
 * @param factorName string因子名称
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_factor_return(date: number | string, factorName: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`fac_ret_sw${factorName}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_factor_returnArr.push({codes:factorName, date, factorName:'', invocation })
  if (ccx_stock_factor_returnTimer !== null) {
    clearTimeout(ccx_stock_factor_returnTimer);
  }
  ccx_stock_factor_returnTimer = setTimeout(() => {
    box(ccx_stock_factor_returnArr, 'fac_ret_sw', 'TradeDate,Factor,DlyReturn',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].Factor && formatDate(item.date) == webArr[i].TradeDate.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['DlyReturn'])
          btn = false
          saveData.put(`fac_ret_sw${item.codes}${item.date}`,webArr[i]['DlyReturn'])
          break;
        }
      }
      if(btn){
        saveData.put(`fac_ret_sw${item.codes}${item.date}`,'NaN')
      }
      return btn
    })
    ccx_stock_factor_returnArr = []
    ccx_stock_factor_returnTimer = null
  }, 500)
}
//因子收益表end

//特质收益表start

interface ccx_stock_specific_returnArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_specific_returnArr: ccx_stock_specific_returnArr[] = []
let ccx_stock_specific_returnTimer: any;//time用来控制事件的触发
/**
 * 个股的特质收益率，即残差部分，代表着股票收益不能被风险因子解释的部分。
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_specific_return(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`specific_ret_sw${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_specific_returnArr.push({ codes, date, invocation })
  if (ccx_stock_specific_returnTimer !== null) {
    clearTimeout(ccx_stock_specific_returnTimer);
  }
  ccx_stock_specific_returnTimer = setTimeout(() => {
    box(ccx_stock_specific_returnArr,'specific_ret_sw','SECU_CODE,TRADINGDAY,SPRET',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].SECU_CODE && formatDate(item.date) == webArr[i].TRADINGDAY.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['SPRET'])
          btn = false;
          saveData.put(`specific_ret_sw${item.codes}${item.date}`,webArr[i]['SPRET'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.SECU_CODE == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['SPRET'])
          saveData.put(`specific_ret_sw${item.codes}${item.date}`,filterCode[0]['SPRET']);
        }else{
          saveData.put(`specific_ret_sw${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })
    ccx_stock_specific_returnArr = []
    ccx_stock_specific_returnTimer = null
  }, 500)
}

//特质收益表end

//险因子协方差矩阵表start

interface ccx_stock_factor_covarianceArr {
  date: string | number;
  factorName1: string;
  factorName2: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_factor_covarianceArr: ccx_stock_factor_covarianceArr[] = []
let ccx_stock_factor_covarianceTimer: any;//time用来控制事件的触发
/**
 * 日度级别的风险因子协方差矩阵
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param date Date交易日期
 * @param factorName1 string因子名称
 * @param factorName2 string因子名称
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_factor_covariance(date: number | string, factorName1: string, factorName2: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`d_cov_vra${factorName1}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_factor_covarianceArr.push({ date, factorName1, factorName2, invocation })
  if (ccx_stock_factor_covarianceTimer !== null) {
    clearTimeout(ccx_stock_factor_covarianceTimer);
  }
  ccx_stock_factor_covarianceTimer = setTimeout(() => {
    const extendsParams = {
      FACTOR:factorName1,
      columns:`TRADE_DATE,FACTOR,${factorName2}`//因子
    }
    box(ccx_stock_factor_covarianceArr,'d_cov_vra',extendsParams,function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.factorName1 == webArr[i].FACTOR && formatDate(item.date) == webArr[i].TRADE_DATE.replace(/-/g,'')){
          item.invocation.setResult(webArr[i][item.factorName2])
          btn = false
          saveData.put(`d_cov_vra${item.factorName1}${item.date}`,webArr[i][item.factorName2])
          break;
        }
      }
      if(btn){
        saveData.put(`d_cov_vra${item.factorName1}${item.date}`,'NaN')
      }
      return btn
    })
    ccx_stock_factor_covarianceArr = []
    ccx_stock_factor_covarianceTimer = null
  }, 500)
}

//险因子协方差矩阵表end

//特质风险表start

interface ccx_stock_specific_riskArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_specific_riskArr: ccx_stock_specific_riskArr[] = []
let ccx_stock_specific_riskTimer: any;//time用来控制事件的触发
/**
 * 日度级别的个股特质风险，特质收益表中特质收益的波动
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_specific_risk(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`d_srisk_vra${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_specific_riskArr.push({ codes, date, invocation })
  if (ccx_stock_specific_riskTimer !== null) {
    clearTimeout(ccx_stock_specific_riskTimer);
  }
  ccx_stock_specific_riskTimer = setTimeout(() => {
    box(ccx_stock_specific_riskArr,'d_srisk_vra','SECU_CODE,TRADE_DATE,SRISK',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].SECU_CODE && formatDate(item.date) == webArr[i].TRADE_DATE.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['SRISK'])
          btn = false
          saveData.put(`d_srisk_vra${item.codes}${item.date}`,webArr[i]['SRISK'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.SECU_CODE == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['SRISK'])
          saveData.put(`d_srisk_vra${item.codes}${item.date}`,filterCode[0]['SRISK']);
        }else{
          saveData.put(`d_srisk_vra${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })
    ccx_stock_specific_riskArr = []
    ccx_stock_specific_riskTimer = null
  }, 500)
}
//特质风险表end
//股票风格属性start
interface ccx_stock_style_vgArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_style_vgArr: ccx_stock_style_vgArr[] = []
let ccx_stock_style_vgTimer: any;//time用来控制事件的触发
/**
 * 将基金定报披露的持仓股票的价值风格得分加权计算得到基金的风格得分，进而得到基金的风格属性
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_style_vg(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`ccx_stock_style_class${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_style_vgArr.push({ codes, date, invocation })
  if (ccx_stock_style_vgTimer !== null) {
    clearTimeout(ccx_stock_style_vgTimer);
  }
  ccx_stock_style_vgTimer = setTimeout(() => {
    box(ccx_stock_style_vgArr,'open_ccx_stock_style_class','secu_code,dt,style_class',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].secu_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['style_class'])
          btn = false
          saveData.put(`ccx_stock_style_class${item.codes}${item.date}`,webArr[i]['style_class'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.secu_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['style_class'])
          saveData.put(`ccx_stock_style_class${item.codes}${item.date}`,filterCode[0]['style_class']);
        }else{
          saveData.put(`ccx_stock_style_class${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })

    ccx_stock_style_vgArr = []
    ccx_stock_style_vgTimer = null
  }, 500)
}
//股票风格属性end
//股票市值属性start
interface ccx_stock_style_capArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_style_capArr: ccx_stock_style_capArr[] = []
let ccx_stock_style_capTimer: any;//time用来控制事件的触发
/**
 * 将基金定报披露的持仓股票的价值风格得分加权计算得到基金的风格得分，进而得到基金的风格属性
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_style_cap(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`ccx_stock_style_class_1${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_style_capArr.push({ codes, date, invocation })
  if (ccx_stock_style_capTimer !== null) {
    clearTimeout(ccx_stock_style_capTimer);
  }
  ccx_stock_style_capTimer = setTimeout(() => {
    box(ccx_stock_style_capArr,'open_ccx_stock_style_class','secu_code,dt,size_class',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].secu_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['size_class'])
          btn = false
          saveData.put(`ccx_stock_style_class_1${item.codes}${item.date}`,webArr[i]['size_class'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.secu_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['size_class'])
          saveData.put(`ccx_stock_style_class_1${item.codes}${item.date}`,filterCode[0]['size_class']);
        }else{
          saveData.put(`ccx_stock_style_class_1${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })

    ccx_stock_style_capArr = []
    ccx_stock_style_capTimer = null
  }, 500)
}
//股票市值属性end
//股票成长价值风格得分srart
interface ccx_stock_style_score_vgArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_style_score_vgArr: ccx_stock_style_score_vgArr[] = []
let ccx_stock_style_score_vgTimer: any;//time用来控制事件的触发
/**
 * 股票成长价值风格得
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_style_score_vg(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`ccx_stock_style_class_2${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_style_score_vgArr.push({ codes, date, invocation })
  if (ccx_stock_style_score_vgTimer !== null) {
    clearTimeout(ccx_stock_style_score_vgTimer);
  }
  ccx_stock_style_score_vgTimer = setTimeout(() => {
    box(ccx_stock_style_score_vgArr,'open_ccx_stock_style_class','secu_code,dt,style_score',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].secu_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['style_score'])
          btn = false
          saveData.put(`ccx_stock_style_class_2${item.codes}${item.date}`,webArr[i]['style_score'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.secu_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['style_score'])
          saveData.put(`ccx_stock_style_class_2${item.codes}${item.date}`,filterCode[0]['style_score']);
        }else{
          saveData.put(`ccx_stock_style_class_2${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })

    ccx_stock_style_score_vgArr = []
    ccx_stock_style_score_vgTimer = null
  }, 500)
}
//股票成长价值风格得分end
//股票市值风格得分start
interface ccx_stock_style_score_capArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_stock_style_score_capArr: ccx_stock_style_score_capArr[] = []
let ccx_stock_style_score_capTimer: any;//time用来控制事件的触发
/**
 * 股票市值风格得分
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_stock_style_score_cap(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`ccx_stock_style_class_3${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_stock_style_score_capArr.push({ codes, date, invocation })
  if (ccx_stock_style_score_capTimer !== null) {
    clearTimeout(ccx_stock_style_score_capTimer);
  }
  ccx_stock_style_score_capTimer = setTimeout(() => {
    box(ccx_stock_style_score_capArr,'open_ccx_stock_style_class','secu_code,dt,size_score',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].secu_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['size_score'])
          btn = false
          saveData.put(`ccx_stock_style_class_3${item.codes}${item.date}`,webArr[i]['size_score'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.secu_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['size_score'])
          saveData.put(`ccx_stock_style_class_3${item.codes}${item.date}`,filterCode[0]['size_score']);
        }else{
          saveData.put(`ccx_stock_style_class_3${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })

    ccx_stock_style_score_capArr = []
    ccx_stock_style_score_capTimer = null
  }, 500)
}
//股票市值风格得分end


// /////////////////////////////////////基金函数
//基金类型start
interface ccx_fund_typeArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_typeArr: ccx_fund_typeArr[] = []
let ccx_fund_typeTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_type(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`fund_type_ccx${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_typeArr.push({ codes, date, invocation })
  if (ccx_fund_typeTimer !== null) {
    clearTimeout(ccx_fund_typeTimer);
  }
  ccx_fund_typeTimer = setTimeout(() => {
    box(ccx_fund_typeArr,'fund_type_ccx','FUNDCODE,TRADE_DT,FUND_TYPE_NAME',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].FUNDCODE && formatDate(item.date) == webArr[i].TRADE_DT.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['FUND_TYPE_NAME'])
          btn = false
          saveData.put(`fund_type_ccx${item.codes}${item.date}`,webArr[i]['FUND_TYPE_NAME'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.FUNDCODE == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['FUND_TYPE_NAME'])
          saveData.put(`fund_type_ccx${item.codes}${item.date}`,filterCode[0]['FUND_TYPE_NAME']);
        }else{
          saveData.put(`fund_type_ccx${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_typeArr = []
    ccx_fund_typeTimer = null
  }, 500)
}
//基金类型end

//基金风格属性start
interface ccx_fund_style_vgArr {
  codes: string;
  date: string | number;
  positionType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_style_vgArr: ccx_fund_style_vgArr[] = []
let ccx_fund_style_vgTimer: any;//time用来控制事件的触发
/**
 * 将基金定报披露的持仓股票的价值风格得分加权计算得到基金的风格得分，进而得到基金的风格属性
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param positionType 持仓类型
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_style_vg(codes: string, date: number | string,positionType = 'detail', invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`ccx_fund_style_class_detail${codes}${date}${positionType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_style_vgArr.push({ codes, date,positionType, invocation })
  if (ccx_fund_style_vgTimer !== null) {
    clearTimeout(ccx_fund_style_vgTimer);
  }
  ccx_fund_style_vgTimer = setTimeout(() => {
    box(ccx_fund_style_vgArr,'open_ccx_fund_style_class_'+positionType,'fund_code,dt,fund_style_score_r6_class',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['fund_style_score_r6_class'])
          btn = false
          saveData.put(`ccx_fund_style_class_detail${item.codes}${item.date}${item.positionType}`,webArr[i]['fund_style_score_r6_class'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['fund_style_score_r6_class'])
          saveData.put(`ccx_fund_style_class_detail${item.codes}${item.date}${item.positionType}`,filterCode[0]['fund_style_score_r6_class']);
        }else{
          saveData.put(`ccx_fund_style_class_detail${item.codes}${item.date}${item.positionType}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_style_vgArr = []
    ccx_fund_style_vgTimer = null
  }, 500)
}
//基金风格属性end

//基金市值属性start
interface ccx_fund_style_capArr {
  codes: string;
  date: string | number;
  positionType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_style_capArr: ccx_fund_style_capArr[] = []
let ccx_fund_style_capTimer: any;//time用来控制事件的触发
/**
 * 将基金定报披露的持仓股票的市值风格得分加权计算得到基金的市值得分，进而得到基金的市值属性
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param positionType 持仓类型
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_style_cap(codes: string, date: number | string,positionType = 'detail', invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${positionType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_style_capArr.push({ codes, date,positionType, invocation })
  if (ccx_fund_style_capTimer !== null) {
    clearTimeout(ccx_fund_style_capTimer);
  }
  ccx_fund_style_capTimer = setTimeout(() => {
    box(ccx_fund_style_capArr,'open_ccx_fund_style_class_'+positionType,'fund_code,dt,fund_size_score_class',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['fund_size_score_class'])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,webArr[i]['fund_size_score_class'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['fund_size_score_class'])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,filterCode[0]['fund_size_score_class']);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_style_capArr = []
    ccx_fund_style_capTimer = null
  }, 500)
}
//基金市值属性end

//基金风格因子暴露start
interface ccx_fund_exposure_styleArr {
  codes: string;
  date: string | number;
  factorName: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_exposure_styleArr: ccx_fund_exposure_styleArr[] = []
let ccx_fund_exposure_styleTimer: any;//time用来控制事件的触发
/**
 * 将基金定报披露的持仓股票在风险模型的行业因子上的暴露加权计算得到基金的行业因子暴露
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param factorName factorName因子
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_exposure_style(codes: string, date: number | string, factorName: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`fund_barra_exposure_sw${codes}${date}${factorName}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_exposure_styleArr.push({ codes, date, factorName, invocation })
  if (ccx_fund_exposure_styleTimer !== null) {
    clearTimeout(ccx_fund_exposure_styleTimer);
  }
  ccx_fund_exposure_styleTimer = setTimeout(() => {
    box(ccx_fund_exposure_styleArr,'fund_barra_exposure_sw',',FUNDCODE,TRADE_DT',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].FUNDCODE && formatDate(item.date) == webArr[i].TRADE_DT.replace(/-/g,'')){
          item.invocation.setResult(webArr[i][item.factorName])
          btn = false
          saveData.put(`fund_barra_exposure_sw${item.codes}${item.date}${item.factorName}`,webArr[i][item.factorName])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.FUNDCODE == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0][item.factorName])
          saveData.put(`fund_barra_exposure_sw${item.codes}${item.date}${item.factorName}`,filterCode[0][item.factorName]);
        }else{
          saveData.put(`fund_barra_exposure_sw${item.codes}${item.date}${item.factorName}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_exposure_styleArr = []
    ccx_fund_exposure_styleTimer = null
  }, 500)
}
//基金风格因子暴露end

//基金市值风格得分start
interface ccx_fund_style_score_capArr {
  codes: string;
  date: string | number;
  positionType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_style_score_capArr: ccx_fund_style_score_capArr[] = []
let ccx_fund_style_score_capTimer: any;//time用来控制事件的触发
/**
 * 基金市值风格得分
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param positionType 持仓类型
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_style_score_cap(codes: string, date: number | string,positionType = 'detail', invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${positionType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_style_score_capArr.push({ codes, date,positionType, invocation })
  if (ccx_fund_style_score_capTimer !== null) {
    clearTimeout(ccx_fund_style_score_capTimer);
  }
  ccx_fund_style_score_capTimer = setTimeout(() => {
    box(ccx_fund_style_score_capArr,'open_ccx_fund_style_class_'+positionType,'fund_code,dt,fund_size_score',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['fund_size_score'])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,webArr[i]['fund_size_score'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['fund_size_score'])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,filterCode[0]['fund_size_score']);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_style_score_capArr = []
    ccx_fund_style_score_capTimer = null
  }, 500)
}
//基金市值风格得分end

//基金成长价值风格得分start
interface ccx_fund_style_score_vgArr {
  codes: string;
  date: string | number;
  positionType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_style_score_vgArr: ccx_fund_style_score_vgArr[] = []
let ccx_fund_style_score_vgTimer: any;//time用来控制事件的触发
/**
 * 基金成长价值风格得分
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param positionType 持仓类型
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_style_score_vg(codes: string, date: number | string,positionType = 'detail', invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${positionType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_style_score_vgArr.push({ codes, date,positionType, invocation })
  if (ccx_fund_style_score_vgTimer !== null) {
    clearTimeout(ccx_fund_style_score_vgTimer);
  }
  ccx_fund_style_score_vgTimer = setTimeout(() => {
    box(ccx_fund_style_score_vgArr,'open_ccx_fund_style_class_'+positionType,'fund_code,dt,fund_style_score_r6',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['fund_style_score_r6'])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,webArr[i]['fund_style_score_r6'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['fund_style_score_r6'])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,filterCode[0]['fund_style_score_r6']);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.positionType}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_style_score_vgArr = []
    ccx_fund_style_score_vgTimer = null
  }, 500)
}
//基金成长价值风格得分end

//基金久期start
interface ccx_fund_durArr {
  codes: string;
  date: string | number;
  scrollNumber:number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_durArr: ccx_fund_durArr[] = []
let ccx_fund_durTimer: any;//time用来控制事件的触发
/**
 * 基金久期指的是基金持有债券组合的久期，根据基金半年报及年报中披露的利率风险敏感分析数据和久期凸度计算公式以及债券投资占比计算得到单期久期，然后求最近若干期的久期平均值
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param scrollNumber 滚动期数
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_dur(codes: string, date: number | string,scrollNumber :number, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${scrollNumber}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_durArr.push({ codes, date,scrollNumber, invocation })
  if (ccx_fund_durTimer !== null) {
    clearTimeout(ccx_fund_durTimer);
  }
  ccx_fund_durTimer = setTimeout(() => {
    box(ccx_fund_durArr,'open_ccx_fund_duration_type','fund_code,dt,'+('dur_mean_rolling_'+scrollNumber),function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['dur_mean_rolling_'+scrollNumber])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,webArr[i]['dur_mean_rolling_'+scrollNumber])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['dur_mean_rolling_'+scrollNumber])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,filterCode[0]['dur_mean_rolling_'+scrollNumber]);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_durArr = []
    ccx_fund_durTimer = null
  }, 500)
}
//基金久期end
//基金久期标签start
interface ccx_fund_dur_labelArr {
  codes: string;
  date: string | number;
  scrollNumber:number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_dur_labelArr: ccx_fund_dur_labelArr[] = []
let ccx_fund_dur_labelTimer: any;//time用来控制事件的触发
/**
 * 基金久期指的是基金持有债券组合的久期，根据基金半年报及年报中披露的利率风险敏感分析数据和久期凸度计算公式以及债券投资占比计算得到单期久期，然后求最近若干期的久期平均值
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param scrollNumber 滚动期数
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_dur_label(codes: string, date: number | string,scrollNumber :number, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${scrollNumber}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_dur_labelArr.push({ codes, date,scrollNumber, invocation })
  if (ccx_fund_dur_labelTimer !== null) {
    clearTimeout(ccx_fund_dur_labelTimer);
  }
  ccx_fund_dur_labelTimer = setTimeout(() => {
    box(ccx_fund_dur_labelArr,'open_ccx_fund_duration_type','fund_code,dt,'+('fund_dur_type_rolling_'+scrollNumber),function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['fund_dur_type_rolling_'+scrollNumber])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,webArr[i]['fund_dur_type_rolling_'+scrollNumber])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['fund_dur_type_rolling_'+scrollNumber])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,filterCode[0]['fund_dur_type_rolling_'+scrollNumber]);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_dur_labelArr = []
    ccx_fund_dur_labelTimer = null
  }, 500)
}
//基金久期标签end

//基金持有的第N大行业名称start
interface ccx_fund_hold_top_industry_nameArr {
  codes: string;
  date: string | number;
  scrollNumber:number;
  industrySort:number;
  industryType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_hold_top_industry_nameArr: ccx_fund_hold_top_industry_nameArr[] = []
let ccx_fund_hold_top_industry_nameTimer: any;//time用来控制事件的触发
/**
 * 基金久期指的是基金持有债券组合的久期，根据基金半年报及年报中披露的利率风险敏感分析数据和久期凸度计算公式以及债券投资占比计算得到单期久期，然后求最近若干期的久期平均值
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param scrollNumber 滚动期数
 * @param industrySort 行业排名
 * @param industryType 行业分类
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_hold_top_industry_name(codes: string, date: number | string,scrollNumber :number,industrySort:number,industryType:string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${scrollNumber}${industrySort}${industryType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_hold_top_industry_nameArr.push({ codes, date,scrollNumber,industrySort,industryType,invocation })
  if (ccx_fund_hold_top_industry_nameTimer !== null) {
    clearTimeout(ccx_fund_hold_top_industry_nameTimer);
  }
  ccx_fund_hold_top_industry_nameTimer = setTimeout(() => {
    const outParams:string = (scrollNumber === 1?
      `top${industrySort}_industry_name_${industryType}`:
      `top${industrySort}_industry_name_${industryType}_rolling_${scrollNumber}`)
    box(ccx_fund_hold_top_industry_nameArr,'open_ccx_fund_industry_top_weight_rolling_'+industryType,'fund_code,dt,'+outParams,function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i][outParams])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}${industrySort}${industryType}`,webArr[i][outParams])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0][outParams])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,filterCode[0][outParams]);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_hold_top_industry_nameArr = []
    ccx_fund_hold_top_industry_nameTimer = null
  }, 500)
}
//基金持有的第N大行业名称end

//基金持有的第N大行业比例start
interface ccx_fund_hold_top_industry_ratioArr {
  codes: string;
  date: string | number;
  scrollNumber:number;
  industrySort:number;
  industryType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_hold_top_industry_ratioArr: ccx_fund_hold_top_industry_ratioArr[] = []
let ccx_fund_hold_top_industry_ratioTimer: any;//time用来控制事件的触发
/**
 * 基金久期指的是基金持有债券组合的久期，根据基金半年报及年报中披露的利率风险敏感分析数据和久期凸度计算公式以及债券投资占比计算得到单期久期，然后求最近若干期的久期平均值
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param scrollNumber 滚动期数
 * @param industrySort 行业排名
 * @param industryType 行业分类
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_hold_top_industry_ratio(codes: string, date: number | string,scrollNumber :number,industrySort:number,industryType:string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${scrollNumber}${industrySort}${industryType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_hold_top_industry_ratioArr.push({ codes, date,scrollNumber,industrySort,industryType,invocation })
  if (ccx_fund_hold_top_industry_ratioTimer !== null) {
    clearTimeout(ccx_fund_hold_top_industry_ratioTimer);
  }
  ccx_fund_hold_top_industry_ratioTimer = setTimeout(() => {
    const outParams:string = (scrollNumber === 1?
      `top${industrySort}_industry_weight_${industryType}`:
      `top${industrySort}_industry_weight_${industryType}_rolling_${scrollNumber}`)
    box(ccx_fund_hold_top_industry_ratioArr,'open_ccx_fund_industry_top_weight_rolling_'+industryType,'fund_code,dt,'+outParams,function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i][outParams])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}${industrySort}${industryType}`,webArr[i][outParams])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0][outParams])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,filterCode[0][outParams]);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_hold_top_industry_ratioArr = []
    ccx_fund_hold_top_industry_ratioTimer = null
  }, 500)
}
//基金持有的第N大行业比例end

//基金行业标签start
interface ccx_fund_industry_labelArr {
  codes: string;
  date: string | number;
  industryType:string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_industry_labelArr: ccx_fund_industry_labelArr[] = []
let ccx_fund_industry_labelTimer: any;//time用来控制事件的触发
/**
 * 根据基金半年报及年报披露全部持仓股票及股票所属的行业计算基金持有行业的比例，滚动最近4期的持仓占比超过50%的行业即为该基金的行业标签
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param industryType 滚动期数
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_industry_label(codes: string, date: number | string,industryType :string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${industryType}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_industry_labelArr.push({ codes, date,industryType, invocation })
  if (ccx_fund_industry_labelTimer !== null) {
    clearTimeout(ccx_fund_industry_labelTimer);
  }
  ccx_fund_industry_labelTimer = setTimeout(() => {
    box(ccx_fund_industry_labelArr,'open_ccx_fund_industry_label_'+industryType,'fund_code,dt,industry_label_name',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['industry_label_name'])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.industryType}`,webArr[i]['industry_label_name'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的值
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['industry_label_name'])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.industryType}`,filterCode[0]['industry_label_name']);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.industryType}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_industry_labelArr = []
    ccx_fund_industry_labelTimer = null
  }, 500)
}
//基金行业标签end

//基金打新收益start
interface ccx_fund_hold_newstock_retArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_hold_newstock_retArr: ccx_fund_hold_newstock_retArr[] = []
let ccx_fund_hold_newstock_retTimer: any;//time用来控制事件的触发
/**
 * 根据新上市股披露的新股中签明细以及基金打新开板即卖的策略估计基金每日的稀有新股的收益
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_hold_newstock_ret(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`fund_new_stock_data${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_hold_newstock_retArr.push({ codes, date, invocation })
  if (ccx_fund_hold_newstock_retTimer !== null) {
    clearTimeout(ccx_fund_hold_newstock_retTimer);
  }
  ccx_fund_hold_newstock_retTimer = setTimeout(() => {
    box(ccx_fund_hold_newstock_retArr,'open_ccx_fund_new_stock_data','fund_code,dt,hold_new_stock_ret',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].secu_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['hold_new_stock_ret'])
          btn = false
          saveData.put(`fund_new_stock_data${item.codes}${item.date}`,webArr[i]['hold_new_stock_ret'])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.secu_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['hold_new_stock_ret'])
          saveData.put(`fund_new_stock_data${item.codes}${item.date}`,filterCode[0]['hold_new_stock_ret']);
        }else{
          saveData.put(`fund_new_stock_data${item.codes}${item.date}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_hold_newstock_retArr = []
    ccx_fund_hold_newstock_retTimer = null
  }, 500)
}
//基金打新收益end
//基金经理卸任标签start
interface ccx_fund_manager_dimisionArr {
  codes: string;
  date: string | number;
  scrollNumber:number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_fund_manager_dimisionArr: ccx_fund_manager_dimisionArr[] = []
let ccx_fund_manager_dimisionTimer: any;//time用来控制事件的触发
/**
 * 基金最近一段时间是否有基金经理卸任
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param scrollNumber 滚动期数
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_fund_manager_dimision(codes: string, date: number | string,scrollNumber :number, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`${codes}${date}${scrollNumber}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_fund_manager_dimisionArr.push({ codes, date,scrollNumber, invocation })
  if (ccx_fund_manager_dimisionTimer !== null) {
    clearTimeout(ccx_fund_manager_dimisionTimer);
  }
  ccx_fund_manager_dimisionTimer = setTimeout(() => {
    box(ccx_fund_manager_dimisionArr,'open_ccx_fund_manager_change_dimissiondate','fund_code,dt,'+('change_'+scrollNumber),function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].fund_code && formatDate(item.date) == webArr[i].dt.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['change_'+scrollNumber])
          btn = false
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,webArr[i]['change_'+scrollNumber])
          break;
        }
      }
      if(btn){
        //如果值为空 择去前一天的指
        const filterCode:any[] = webArr.filter(obj=>obj.fund_code == item.codes).reverse();
        console.log(filterCode);
        if(window.nullValue === 2 && filterCode && filterCode.length != 0){
          btn = false;
          item.invocation.setResult(filterCode[0]['change_'+scrollNumber])
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,filterCode[0]['change_'+scrollNumber]);
        }else{
          saveData.put(`ccx_fund_style_class_key${item.codes}${item.date}${item.scrollNumber}`,'NaN');
        }
      }
      return btn
    })

    ccx_fund_manager_dimisionArr = []
    ccx_fund_manager_dimisionTimer = null
  }, 500)
}
//基金经理卸任标签end
/*******中诚信指数函数start***********/
//指数代码
interface ccx_index_codeArr {
  type: string;
  invocation: string|undefined//单元格信息
}
let ccx_index_codeArr: ccx_index_codeArr[] = []
let ccx_index_codeTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param type 基金代码
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
*/
export function ccx_index_code(type='index_basic',invocation: CustomFunctions.Invocation) {
  ccx_index_codeArr.push({ type, invocation:invocation.address })
  if (ccx_index_codeTimer !== null) {
    clearTimeout(ccx_index_codeTimer);
  }
  ccx_index_codeTimer = setTimeout(() => {
    box(ccx_index_codeArr,'open_ccx_fof_'+type,'index_code,full_name,label',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}){
        if(true){
          console.log('webArr',webArr);
          
          const data = webArr.filter(a=>a.label == 'prod')
          const context = new Excel.RequestContext();
          let sheet = context.workbook.worksheets.getActiveWorksheet()
          //设置地址
          const addrStart:string = item.invocation && item.invocation.split('!')[1] || 'A1'
          // const addrStart:string = 'A1'
          const addrStr = (addrStart.match(reg) as any)[0];//拿到地址的字母A1
          const addrNum:number =Number(addrStart.replace(/[^\d.]/g, ''))//拿到地址的数字 A1 
          const endAddr = `${addrStart}:${transformToStr(strToNamber(addrStr) + 1)}${addrNum +data.length - 1 + 1 + 1}`
          console.log(endAddr);
          
          let setData:any[][] = []
          data.map((str:any)=>{
            setData.push([str.index_code,str.full_name])
          })
          console.log(setData);
          setData.unshift(['code','name'])
          setData.unshift([null,item.codes])
          
          let range = sheet.getRange(endAddr);
          // range.format.autofitColumns()
          range.values = setData;
          (async()=>{
            await context.sync();
          })();
        }
    })

    ccx_index_codeArr = []
    ccx_index_codeTimer = null
  }, 500)

  return 'type:'+type
}
//指数名称
interface ccx_index_nameArr {
  codes: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_index_nameArr: ccx_index_nameArr[] = []
let ccx_index_nameTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_index_name(codes: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`index_basic_info${codes}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_index_nameArr.push({ codes, invocation })
  if (ccx_index_nameTimer !== null) {
    clearTimeout(ccx_index_nameTimer);
  }
  ccx_index_nameTimer = setTimeout(() => {
    box(ccx_index_nameArr,'open_ccx_fof_index_basic','index_code,full_name',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].index_code){
          item.invocation.setResult(webArr[i]['full_name'])
          btn = false
          saveData.put(`index_basic_info${item.codes}`,webArr[i]['full_name'])
          break;
        }
      }
      return btn
    })

    ccx_index_nameArr = []
    ccx_index_nameTimer = null
  }, 500)
}
//指数基期
interface ccx_index_base_dateArr {
  codes: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_index_base_dateArr: ccx_index_base_dateArr[] = []
let ccx_index_base_dateTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_index_base_date(codes: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`index_basic_info_base_date${codes}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_index_base_dateArr.push({ codes, invocation })
  if (ccx_index_base_dateTimer !== null) {
    clearTimeout(ccx_index_base_dateTimer);
  }
  ccx_index_base_dateTimer = setTimeout(() => {
    box(ccx_index_base_dateArr,'open_ccx_fof_index_basic','index_code,base_date',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].index_code){
          item.invocation.setResult(webArr[i]['base_date'])
          btn = false
          saveData.put(`index_basic_info_base_date${item.codes}`,webArr[i]['base_date'])
          break;
        }
      }
      return btn
    })

    ccx_index_base_dateArr = []
    ccx_index_base_dateTimer = null
  }, 500)
}
//换仓频率
interface ccx_index_rebalance_frequencyArr {
  codes: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_index_rebalance_frequencyArr: ccx_index_rebalance_frequencyArr[] = []
let ccx_index_rebalance_frequencyTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_index_rebalance_frequency(codes: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`index_basic_info_rebalance_period${codes}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_index_rebalance_frequencyArr.push({ codes, invocation })
  if (ccx_index_rebalance_frequencyTimer !== null) {
    clearTimeout(ccx_index_rebalance_frequencyTimer);
  }
  ccx_index_rebalance_frequencyTimer = setTimeout(() => {
    box(ccx_index_rebalance_frequencyArr,'open_ccx_fof_index_basic','index_code,rebalance_period',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].index_code){
          const str = index_rebalance_frequencyList[webArr[i]['rebalance_period']]
          item.invocation.setResult(str)
          btn = false
          saveData.put(`index_basic_info_rebalance_period${item.codes}`,str)
          break;
        }
      }
      return btn
    })

    ccx_index_rebalance_frequencyArr = []
    ccx_index_rebalance_frequencyTimer = null
  }, 500)
}
//加权方式
interface ccx_index_weight_methodArr {
  codes: string;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_index_weight_methodArr: ccx_index_weight_methodArr[] = []
let ccx_index_weight_methodTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_index_weight_method(codes: string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`index_basic_info_weight_method${codes}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_index_weight_methodArr.push({ codes, invocation })
  if (ccx_index_weight_methodTimer !== null) {
    clearTimeout(ccx_index_weight_methodTimer);
  }
  ccx_index_weight_methodTimer = setTimeout(() => {
    box(ccx_index_weight_methodArr,'open_ccx_fof_index_basic','index_code,weight_method',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].index_code){
          const str = index_weight_methodList[webArr[i]['weight_method']]
          item.invocation.setResult(str)
          btn = false
          saveData.put(`index_basic_info_weight_method${item.codes}`,str)
          break;
        }
      }
      return btn
    })

    ccx_index_weight_methodArr = []
    ccx_index_weight_methodTimer = null
  }, 500)
}
//指数成分权重
interface ccx_index_weightArr {
  codes: string;
  date: string | number;
  invocation: string|undefined//单元格信息
}
let ccx_index_weightArr: ccx_index_weightArr[] = []
let ccx_index_weightTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
*/
export function ccx_index_weight(codes: string, date: number | string, invocation: CustomFunctions.Invocation) {
  ccx_index_weightArr.push({ codes, date, invocation:invocation.address })
  console.log(codes);
  
  if (ccx_index_weightTimer !== null) {
    clearTimeout(ccx_index_weightTimer);
  }
  ccx_index_weightTimer = setTimeout(() => {
    box(ccx_index_weightArr,'open_ccx_fof_index_weight','index_code,trade_date,weight_content',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].index_code && formatDate(item.date) == webArr[i].trade_date.replace(/-/g,'')){
          const data = JSON.parse(webArr[i]['weight_content']) 
          const context = new Excel.RequestContext();
          let sheet = context.workbook.worksheets.getActiveWorksheet()
          //设置地址
          const addrStart:string = item.invocation && item.invocation.split('!')[1] ||item.invocation || 'A1'
          // const addrStart:string = 'A1'
          const addrStr = (addrStart.match(reg) as any)[0];//拿到地址的字母A1
          const addrNum:number =Number(addrStart.replace(/[^\d.]/g, ''))//拿到地址的数字 A1 
          const endAddr = `${addrStart}:${transformToStr(strToNamber(addrStr) + 2)}${addrNum +data.length - 1 + 1 + 1}`
          console.log(endAddr);
          
          let setData:any[][] = []
          data.map((str:any)=>{
            setData.push([webArr[i].trade_date,str.components_code,str.weight])
          })
          setData.unshift(['date','code','weight'])
          if(i === 0){
            setData.unshift([null,null,null])
          }else{
            setData.unshift([item.codes,null,null])
          }
          let range = sheet.getRange(endAddr);
          range.values = setData;
          (async()=>{
            await context.sync();
          })();
          btn = false
          break;
        }
      }
      return btn
    })

    ccx_index_weightArr = []
    ccx_index_weightTimer = null
  }, 500)
  console.log('ccx_index_weightArr[0].',codes);
  
  return codes
}
//指数点位
interface ccx_index_levelsArr {
  codes: string;
  date: string | number;
  invocation: CustomFunctions.StreamingInvocation<string>//单元格信息
}
let ccx_index_levelsArr: ccx_index_levelsArr[] = []
let ccx_index_levelsTimer: any;//time用来控制事件的触发
/**
 * 根据基金合同以及基金定报披露的股票仓位确认基金类型
 * @customfunction
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 基金代码
 * @param date Date交易日期
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 * @return {string}
 */
export function ccx_index_levels(codes: string, date: number | string, invocation: CustomFunctions.StreamingInvocation<string>) {
  let saveIs = saveData.getVal(`index_level${codes}${date}`)
  if(saveIs){
    return invocation.setResult(saveIs)
  }
  ccx_index_levelsArr.push({ codes, date, invocation })
  if (ccx_index_levelsTimer !== null) {
    clearTimeout(ccx_index_levelsTimer);
  }
  ccx_index_levelsTimer = setTimeout(() => {
    box(ccx_index_levelsArr,'open_ccx_fof_index_levels','index_code,trade_date,index_level',function ({ btn,webArr,item }:{btn:boolean,webArr:any[],item:any}):boolean{
      for(let i=0;i<webArr.length;i++){
        if(item.codes == webArr[i].index_code && formatDate(item.date) == webArr[i].trade_date.replace(/-/g,'')){
          item.invocation.setResult(webArr[i]['index_level'])
          btn = false
          saveData.put(`index_level${item.codes}${item.date}`,webArr[i]['index_level'])
          break;
        }
      }
      return btn
    })

    ccx_index_levelsArr = []
    ccx_index_levelsTimer = null
  }, 500)
}
//中诚信指数函数end

//导入数据
/**
 * description 因子暴露表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
  export function ccx_stock_exposureList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
    return exposureList(codes,functionArg,time,nullType,axis,invocation.address)
  }
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
  export function ccx_stock_specific_returnList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
    return specificReturnList(codes,time,nullType,axis,invocation.address)
  } 
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
 export function ccx_stock_style_vgList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return stock_style_vg_List(codes,time,nullType,axis,invocation.address)
}  
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
 export function ccx_stock_style_capList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return stock_style_cap_List(codes,time,nullType,axis,invocation.address)
} 
/**
 * description 股票成长价值风格得分
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_stock_style_score_vgList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return stock_style_score_vg_List(codes,time,nullType,axis,invocation.address)
} 
/**
 * description 股票成长价值风格得分
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_stock_style_score_capList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return stock_style_score_cap_List(codes,time,nullType,axis,invocation.address)
} 
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */

export function ccx_stock_specific_riskList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return stock_specific_risk_List(codes,time,nullType,axis,invocation.address)
}  
/*********************基金*************************/
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
 export function ccx_fund_typeList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return fund_type_List(codes,time,nullType,axis,invocation.address)
}  
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_style_vgList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return fund_style_vg_List(codes,functionArg,time,nullType,axis,invocation.address)
}
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */

export function ccx_fund_style_capList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return fund_style_cap_List(codes,functionArg,time,nullType,axis,invocation.address)
}
/**
 * description 因子暴露表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
 export function ccx_fund_exposure_styleList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return fund_exposure_style_List(codes,functionArg,time,nullType,axis,invocation.address)
}  
/**
 * description 基金市值风格得分
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_style_score_capList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return fund_style_score_cap_List(codes,functionArg,time,nullType,axis,invocation.address)
} 
/**
 * description 基金成长价值风格得分
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_style_score_vgList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return fund_style_score_vg_List(codes,functionArg,time,nullType,axis,invocation.address)
} 
/**
 * description 基金久期指的是基金持有债券组合的久期，根据基金半年报及年报中披露的利率风险敏感分析数据和久期凸度计算公式以及债券投资占比计算得到单期久期，然后求最近若干期的久期平均值
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_durList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_dur_List(codes,functionArg,time,nullType,axis,invocation.address)
} 
/**
 * description 基金久期标签
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_dur_labelList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_dur_label_List(codes,functionArg,time,nullType,axis,invocation.address)
} 
/**
 * description 基金持有的第N大行业名称
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param {number} scrollNumber 滚动期数
 * @param {number} industrySort 行业排名
 * @param {string} industryType 行业类型
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_hold_top_industry_nameList (codes:string[][],scrollNumber:number,industrySort:number,industryType:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_hold_top_industry_name_List(codes,scrollNumber,industrySort,industryType,time,nullType,axis,invocation.address)
} 
/**
 * description 基金持有的第N大行业比例
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param {number} scrollNumber 滚动期数
 * @param {number} industrySort 行业排名
 * @param {string} industryType 行业类型
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_hold_top_industry_ratioList (codes:string[][],scrollNumber:number,industrySort:number,industryType:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_hold_top_industry_ratio_List(codes,scrollNumber,industrySort,industryType,time,nullType,axis,invocation.address)
} 
/**
 * description 根据基金半年报及年报披露全部持仓股票及股票所属的行业计算基金持有行业的比例，滚动最近4期的持仓占比超过50%的行业即为该基金的行业标签
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_industry_labelList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_industry_label_List(codes,functionArg,time,nullType,axis,invocation.address)
} 
/**
 * description 根据基金半年报及年报披露全部持仓股票及股票所属的行业计算基金持有行业的比例，滚动最近4期的持仓占比超过50%的行业即为该基金的行业标签
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_hold_newstock_retList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_hold_newstock_ret_List(codes,time,nullType,axis,invocation.address)
} 
/**
 * description 基金久期指的是基金持有债券组合的久期，根据基金半年报及年报中披露的利率风险敏感分析数据和久期凸度计算公式以及债券投资占比计算得到单期久期，然后求最近若干期的久期平均值
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param functionArg 函数参数
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {number} axis 1 x:时间y:codes
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */
export function ccx_fund_manager_dimisionList (codes:string[][],functionArg:string,time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return ccx_fund_manager_dimision_List(codes,functionArg,time,nullType,axis,invocation.address)
} 
/**********中诚信***************/    
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */

 export function ccx_index_levelsList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return index_levels_List(codes,time,nullType,axis,invocation.address)
}
/**
 * description 特质收益表
 * @customfunction 
 * @helpurl https://excel.ccxindices.com/help.html
 * @param codes 股票code集合
 * @param time 时间start,end,dataFormat,dataType,frequency/开始，结束，时间格式，日历日或交易日，日期频率
 * @param nullType
 * @param {CustomFunctions.Invocation} invocation Custom function invocation
 * @requiresAddress
 */

 export function ccx_index_weightList (codes:string[][],time:string,nullType:number,axis:number,invocation:CustomFunctions.Invocation){
  return index_weight_List(codes,time,invocation.address)
}
//导入数据