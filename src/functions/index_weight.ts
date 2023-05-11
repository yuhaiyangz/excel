/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-11-22 12:17:29
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2023-02-22 15:08:00
 * @Description: 指数权重
 */
import { getTime } from './functions'
import { transformToStr,strToNamber,reg } from './../taskpane/unit'
import axios from './axios'
export default async function index_weight_List(codes:string[][],time:string,addr:string|undefined){    
    if(typeof codes != 'object') return 'codes输入不正确'
    if(typeof time != 'string') return 'time输入不正确'
    const timeStr = time.split(',')
    if(timeStr.length != 5) return 'time输入不正确'
    //获取输入地址
    console.log(addr);
    
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
    const timeList =await getTime(startDate,endDate,timeStr[2],timeStr[3])
    let dateListReverse:string[] = timeList.data.data.map((item:any)=>item.date).reverse()
    if(dateListReverse.length == 0) return '时间获取为空'
    let dateList:string[] = []
    if(dataFormat == 1){
        dateList =dateListReverse.map((item:any)=>item.replace(/-/g,''));
    }else if(dataFormat == 2){
        dateList =dateListReverse.map((item:any)=>item.replace(/-/g,'/'));
    }else if(dataFormat == 3){
        dateList =dateListReverse
    }
    const context = new Excel.RequestContext();
    let sheet = context.workbook.worksheets.getActiveWorksheet()
    //创建提示
    let rangePrompt = sheet.getRange(addrStart);
    console.log(rangePrompt);
    
    rangePrompt.format.fill.color = "#ffeb3b";
    rangePrompt.dataValidation.prompt = {
        message: "Excel—中诚信指数—函数编辑",
        showPrompt: true, // The default is 'false'.
        title: "如需修改请使用"
    };
    await context.sync();
    //请求数据
    const extendsParams = {
        codes:codeList,//股票代码
        columns:'index_code,trade_date,weight_content'//因子
      }
    const params = {
        tableName:'open_ccx_fof_index_weight',//表名
        extendsParams: JSON.stringify(extendsParams),
        startDate:dateListReverse[0].replace(/-/g,''),
        endDate:dateListReverse[dateListReverse.length - 1].replace(/-/g,'')
    }
    interface APIType{
        index_code:string;
        trade_date:string;
        weight_content:string;
    }
    axios.post('/',params).then((res:any)=>{
        let newData:APIType[][] = []//相同code不同日期
        const data = res.data.data || [];
        codeList.map((code,codeIndex)=>{
            newData[codeIndex] = []
            newData[codeIndex] = data.filter((item:any)=>item.index_code == code).sort((a:any,b:any)=>a.trade_date - b.trade_date)
        })
        const addrStr = (addrStart.match(reg) as any)[0];//拿到地址的字母 A
        const addrNum:number =Number(addrStart.replace(/[^\d.]/g, ''))//拿到地址的数字 1 
        newData.map(async(itemArr,index)=>{
            let arr:any[][] = []
            itemArr.map(item=>{
                const weightData = JSON.parse(item.weight_content)
                arr.push(...weightData.map((obj:any)=>[item.trade_date,obj.components_code,obj.weight]))
            })
            console.log(arr);
            
            if(index == 0){
                arr.unshift([null,'code','weight'])
            }else{
                arr.unshift(['date','code','weight'])
            }
            arr.unshift([itemArr[0]?.index_code,null,null])
            const address = `${transformToStr(strToNamber(addrStr) + index*5)}${addrNum - 1}:${transformToStr(strToNamber(addrStr) + index*5 + 2)}${addrNum - 1 + arr.length - 1}`
            console.log(address);
            try{
                let range = sheet.getRange(address);
                range.values = arr;
                range.format.autofitColumns()
                await context.sync();                
            }catch(err){
                console.log(err);
                window.alertErr('数据过大，表格放不下了，请缩短日期和指数个数后重试！')
            }

        })
    })
    return 'date'
}