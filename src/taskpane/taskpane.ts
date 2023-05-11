/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-08-24 16:32:45
 * @LastEditors: yuhaiyang yuhaiyangz@163.com
 * @LastEditTime: 2023-05-11 10:05:21
 * @Description: 请填写简介
 */
/* global console, document, Excel, Office */
// The initialize function must be run each time a new page is loaded
import { transformToStr,strToNamber,reg } from './unit'
window.token = {
    userName: '',
    passWord: ''
};
// window.token = {//测试内网
//     userName: 'a0b9a4f1b5824ac7bf04b0d4321e05a0',
//     passWord: 'a0b9a4f1b5824ac7bf04b0d4321e05a0'
// };

window.ifLoadSuccess = false
try {
    Office.onReady(()=>{
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
        Office.context.document.settings.saveAsync();
        Office.addin.showAsTaskpane().then(_=>{
            console.log('加载项加载完成');
            window.ifLoadSuccess = true;
        });
    })    
} catch (error) {
    console.log(error);
    window.ifLoadSuccess = true
}
function enableButton(type:boolean) {//登录是否禁用false禁用true启用
    console.log(type);
    setTimeout(()=>{
        Office.ribbon.requestUpdate({
            tabs: [
                {
                    id: "TabHome", 
                    groups: [
                        {
                        id: "Ccxd.outLogn",
                        controls: [
                            {
                                id: "Ccxd.outLogn.Btn", 
                                enabled:type
                            }
                        ]
                        }
                    ]
                }
            ]
        });        
    },500)

}
window.enableButton = enableButton
//公用函数
    //获取区域内的内容
    async function getAddrValues(callBack:Function,addr:string){
        try {
            await Excel.run(async (context) => {
                let sheet = context.workbook.worksheets.getActiveWorksheet();            
                let range = sheet.getRange(addr);
                range.load("values");
                await context.sync();
                callBack(range.values)
            });            
        } catch (error) {
            console.log(error); 
        }
    }
//end
//打开函数搜索

function funSearch(event:any) {
    let dialog:any;
    let timer:any;
    let excAddr:string = ''
    let excAddrValueIsNull:boolean = true
    async function run(type:number) {//1代表记录在哪个位置插入公式，2代表记录定时器，查询函数参数插入的cell位置
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.load("address");
                await context.sync();
                if (type == 1) {
                    excAddr = range.address.split('!')[1]
                    getAddrValues(function(val:string){
                        if(val && val[0] != ''){
                            excAddrValueIsNull = false
                        }else{
                            excAddrValueIsNull = true
                        }
                        send(JSON.stringify({type:1,value:excAddrValueIsNull}))
                    },excAddr)
                } else {
                    send(JSON.stringify({type:2,value:range.address}))
                }
    
                console.log(`The range address was ${range.address}.`);
            });
        } catch (error) {
            console.log(error);
        }
    }
    function send(str:string) {//发送给弹窗的信息
        dialog.messageChild(str);
    }
    if(!window.token.userName || !window.token.passWord){
        window.onLogin()
        event.completed();
        return
    }
    Office.context.ui.displayDialogAsync(window.location.origin+'/functionSearch.html', { height: 65, width: 50 }, function (asyncResult) {
        event.completed();
        dialog = asyncResult.value;
        setTimeout(() => {
            run(1)
        }, 1000)

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, async function (arg:{message:string}) {
            console.log(arg);
            if (arg.message === 'start') {//开始选取地址
                run(2)
                timer = setInterval(() => {
                    run(2)
                }, 1000)
                return
            }
            if (arg.message === 'stop') {//停止选区地址
                timer && clearInterval(timer)
                return
            }
            if (arg.message === 'againGetAddr') {//重新选择插入地址
                timer = setInterval(()=>{
                    run(1)
                },1000)
                return
            }   
            if (arg.message.includes('close')) {//关闭窗口
                const msg = JSON.parse(arg.message)
                console.log('关闭窗口', msg);
                dialog.close()
                dialog = null
                timer && clearInterval(timer)
                await Excel.run(async (context) => {
                    let sheet = context.workbook.worksheets.getActiveWorksheet();
                    let dataRange = sheet.getRange(excAddr);
                    dataRange.values = [[msg.code]]
                    sheet.load()
                    dataRange.format.autofitColumns();
                    // sheet.calculate(false);//true标记为脏，如果插入数据即可引发其他数据自定更新
                    // context.application.calculationMode = 'Manual'
                    // let range = sheet.getRange('A1');
                    // range.setDirty()
                    await context.sync();
                    // sheet.enableCalculation = true
                })
                return
            }
            if (arg.message === 'end') {//直接关闭
                dialog.close()
                dialog = null
                return
            }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (err:any) {//错误
            console.log(err);
            dialog.close()
            dialog = null
        });
    });
}
//d登录或登出
function logout(event:any) {
    // 如果已登录则退出，如果未登录则登录
    if(!window.token.userName || !window.token.passWord){
        window.onLogin()
    }else{
        Office.addin.showAsTaskpane().then(res=>{
            console.log('加载项加载完成',res);
            window.onLogout() 
        })
    }
   event.completed();
}
//导入数据
const getData = async(event:any,type?:string)=>{
    let dialog:any;
    let timer:any = null
    let excAddr:string = '';//记录点开弹窗时选择的哪个单元格
    let excAddrValueIsNull:boolean = true;//记录点开弹窗时选择的哪个单元格是否为空
    let excAddrArea:string = '';//区域
    let funCodeUpdate:string|null = null//获取公式
    async function run(type:number) {
        try {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                if(type == 4){//获取公式
                    range.load("formulas");
                    await context.sync();
                    funCodeUpdate = range.formulas[0][0];
                    console.log(funCodeUpdate);
                    return
                }
                range.load("address");
                await context.sync();
                const addr = range.address.split('!')[1]
                if(type == 2){
                    send(JSON.stringify({type:3,value:addr}))
                    excAddrArea = addr
                }else if(type == 1){
                    excAddr = addr.split(':')[0];
                    getAddrValues(function(val:string){
                        if(val && val[0] != ''){
                            excAddrValueIsNull = false
                        }else{
                            excAddrValueIsNull = true
                        }
                        send(JSON.stringify({type:6,value:excAddrValueIsNull}))
                    },excAddr)
                }else if(type == 3){
                    excAddr = addr.split(':')[0]
                    send(JSON.stringify({type:5,value:excAddr}))
                }
                console.log(`The range address was ${excAddr} ${excAddrArea}.`);
            });
        } catch (error) {
            console.log(error);
        }
    }
    function send(str:string) {//发送给弹窗的信息
        dialog.messageChild(str);
    }
    /**
     * @description: 
     * @param {string} addr excel单元格地址
     * @param {string} codeList code编码列表
     * @param {any} funCode 函数信息{argList:参数列表，name:函数名，code:函数英文名}
     * @param {string} startTime 开始时间
     * @param {string} endTime 结束时间
     * @param {number} dataType 是否交易日
     * @param {number} frequency 日期频率
     * @param {number} dataFormat 日期格式
     * @param {number} nullType 是否前置空值
     * @param {number} axis 坐标轴
     * @return {*}
     */    
    const setDataForExcel = async (addr:string,codeList:{name:string,code:string}[],funCode:{argList:any[],name:string,code:string},startTime:string,endTime:string,dataType:number,frequency:number,dataFormat:number,nullType:number,axis:number)=>{
        // 地址转换
        const addrStr = (addr.match(reg) as any)[0];//拿到地址的字母A
        const addrNum:number = Number(addr.replace(/[^\d.]/g, ''))//拿到地址的数字 1
        console.log(addrStr,addrNum);
        let setCodesAddr:string = ''//name+code地址
        let setCodesAddrCode:string = ''//code地址
        let endValues:any[][] = [];//name+code值
        if(axis == 1){
            const endNumer = addrNum + codeList.length - 1;//拿到矩阵结尾的横坐标（此时为数字需要转换成AAB形式）
            setCodesAddr = `${transformToStr(strToNamber(addrStr) - 1 - 1)}${addrNum}:${transformToStr(strToNamber(addrStr) - 1)}${endNumer}`//矩阵开始结尾横纵坐标 codes位置
            setCodesAddrCode = `${transformToStr(strToNamber(addrStr) - 1)}${addrNum}:${transformToStr(strToNamber(addrStr) - 1)}${endNumer}`//矩阵开始结尾横纵坐标 codes位置
            codeList.map((item,index)=>{
                endValues[index] = []
                endValues[index].push(item.name)
                endValues[index].push(item.code)
            })            
        }else{
            setCodesAddr = `${addrStr}${addrNum - 1 - 1}:${transformToStr(strToNamber(addrStr) + codeList.length - 1)}${addrNum - 1 }`//矩阵开始结尾横纵坐标 codes位置
            setCodesAddrCode = `${addrStr}${addrNum - 1}:${transformToStr(strToNamber(addrStr) + codeList.length - 1)}${addrNum - 1 }`//矩阵开始结尾横纵坐标 codes位置
            endValues[0] = codeList.map(item=>item.name) 
            endValues[1] = codeList.map(item=>item.code) 
        }
        console.log(setCodesAddrCode);

        //end
        try {
            await Excel.run(async (context) => {//设置name+code
                let sheet = context.workbook.worksheets.getActiveWorksheet();
                let dataRange = sheet.getRange(setCodesAddr);
                dataRange.values = endValues
                dataRange.format.autofitColumns()
                await context.sync();
            })            
        } catch (error) {
            console.log(error);       
        }
        //设置公式名称，公式可选择入参
        const funStrSetAddr = `${transformToStr(strToNamber(addrStr))}${addrNum - 2 - (axis == 1?0:1) }`//函数名地址
        //参数名+值
        const argStrSetAddrStr = `${transformToStr(strToNamber(addrStr)+1)}${addrNum - 2 - (axis == 1?0:1) }:${transformToStr(strToNamber(addrStr)+1 + funCode.argList.length * 2 - 1)}${addrNum - 2 - (axis == 1?0:1) }`//函数名地址
        console.log(funCode);
        console.log(funStrSetAddr,argStrSetAddrStr);
        try {
            await Excel.run(async (context) => {
                let sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.getRange(funStrSetAddr).values = [[funCode.name]];
                if(funCode.argList && funCode.argList.length != 0){
                    let strArr:string[] = []
                    funCode.argList.forEach(item=>{
                        strArr.push(item.label)
                        strArr.push(item.arg)
                    })
                    console.log(strArr);
                    sheet.getRange(argStrSetAddrStr).values = [strArr];
                }
                
                await context.sync();
            })            
        } catch (error) {
            console.log(error);       
        }
        //公式
        const funSetAddr = `${addr}`
        let argStrSetAddr:string[] = []//参数值地址
        function setAddrArgfun(){//设置函数参数地址
            funCode.argList.forEach((item:any,index:number)=>{
                argStrSetAddr.push(`${transformToStr(strToNamber(addrStr) + 1 + (index*2 + 1))}${addrNum - 2 - (axis == 1?0:1) }`)
            })
        }
        setAddrArgfun()
        try {
            await Excel.run(async (context) => {
                let sheet = context.workbook.worksheets.getActiveWorksheet();
                let values:any = null
                if(argStrSetAddr.length != 0){
                    let argListStr:string = ''
                    argStrSetAddr.forEach((item)=>{
                        argListStr += `${item},`
                    })
                    console.log(argListStr);
                    
                    values = [[`=CCX.${(funCode.code + 'List').toUpperCase()}(${setCodesAddrCode},${argListStr}"${startTime},${endTime},${dataType},${frequency},${dataFormat}",${nullType},${axis})`]];
                }else{
                    values = [[`=CCX.${(funCode.code + 'List').toUpperCase()}(${setCodesAddrCode},"${startTime},${endTime},${dataType},${frequency},${dataFormat}",${nullType},${axis})`]];
                }
                sheet.getRange(funSetAddr).values = values
                await context.sync();
            })            
        } catch (error) {
            console.log(error);       
        }
    }
    function loopFun(API:Function,arr:any[],i:number,callBack:Function){
        API(i).then((res:any)=>{
                if(res.data.code !== 2000) return arr
                if(res.data.data.length != 0 ){
                    i = i+200
                    arr.push(...res.data.data)
                    loopFun(API,arr,i,callBack)
                }else{
                    callBack()
                }
            }) 
    }
    const getApi = (API:Function,arr:any[],i:number)=>{
        return new Promise((resolve,reject)=>{
            try {
                loopFun(API,arr,i,function(){
                    resolve(arr)
                })                  
            } catch (error) {
                reject(error)
            }
     
        })

    }
    if(!window.token.userName || !window.token.passWord){
        window.onLogin()
        event.completed();
        return
    }
        const options = {
            asyncContext: '导入数据',
            height: 70,
            width: 65
        }
        let URL = window.location.origin+'/getdata.html'
        const isFirstCookie = document.cookie.match(`[;\s+]?isFirst=([^;]*)`)?.pop();
        if(isFirstCookie){
            URL = window.location.origin+'/getdata.html?isFirst=true';
        }
        if(type){
            URL += '&type=update'
            run(4)
        }
        Office.context.ui.displayDialogAsync(URL, options,function (asyncResult) {
            event.completed(); 
            dialog = asyncResult.value
            //获取股票基金code
            interface codeListSZType {
                is_hs300:           boolean;
                is_zz800:           boolean;
                main_board:         boolean;
                is_zz1000:          boolean;
                chinameabbr:            string;
                is_zz500:           boolean;
                secu_code:          string;
                ashare_with_delist: boolean;
                sz_market:          boolean;
                ashare_valid:       boolean;
            }            
            let SZLIST:codeListSZType[] = [];//股票
            let OFLIST:any[] = [];//基金
            let INDEXLIST:any[] = [];//忠诚信指数
            let allData:any = {
                sharesList:[],//股票
                fundList:[],//基金
                indexList:[]//中诚信指数
            }
            const isFirstCookie = document.cookie.match(`[;\s+]?isFirst=([^;]*)`)?.pop();

            if (!isFirstCookie) {
                Promise.all([
                    getApi(window.getCodeListSZ,SZLIST,0),
                    getApi(window.getCodeListOF,OFLIST,0),
                    getApi(window.getCodeListIndex,INDEXLIST,0)
                ]).then((res:any[])=>{
                    console.log(res);
                    let endTime = new Date(new Date().toLocaleDateString()).getTime() +24 * 60 * 60 * 1000 -1;// 当天23:59:59
                    var exp = new Date(endTime);
                    document.cookie = "isFirst=true;expires=" + exp.toUTCString() + ";path=/";
                    allData.sharesList = res[0].map((item:any)=>{
                        let strType = ''
                        Object.keys(item).map((key:any)=>{
                            //@ts-ignore
                            if(item[key]){
                                strType +=key+','
                            }
                        })
                        return {code:item.secu_code,name:item.chinameabbr,type:strType}
                    })
                    allData.fundList = res[1].map((item:any)=>{
                        return {code:item.fund_code,name:item.chinameabbr,type:item.fund_type_ccx_name,main:item.fund_valid_main}
                    })
                    const list = INDEXLIST.filter(item=>{
                        return item.label == 'prod'
                    })
                    allData.indexList = list.map(item=>{
                        return {code:item.index_code,name:item.short_name}
                    })
                    setTimeout(()=>{
                        dialog.messageChild(JSON.stringify({type:1,list:allData}));
                    },800)
                    
                })      
            }
            //函数更新
            function getFunUpdate(){
                const reg = /^=CCX\.(\S+)\(/ //匹配公式
                const strArr = (funCodeUpdate as string).match(reg)
                let funAnalysis = {//公式解析
                    code:'',//公式
                    codesList:[],//code列表
                    functionArg:'',//获取传入的值
                    timer:{///开始，结束，日历日或交易日，日期频率，时间格式
                        start:'',
                        end:'',
                        dataType:'',
                        frequency:'',
                        dataFormat:''
                    },
                    nullValue:'',//是否fill空值
                    axis:''
                }
                console.log(strArr);
                
                if(strArr){
                    funAnalysis.code = strArr[1]
                }else{
                    return window.alertErr('获取地址非公式')
                }
                var reg1 = /\((.+)\)/ //匹配参数
                const argStr = (funCodeUpdate as string).match(reg1) && ((funCodeUpdate as string).match(reg1) as any)[1]//拿到参数字符串'A3:A302,"20221101,20221116,2,1,1",1'
                //取出双引号
                const reg2 = /"(.+)"/
                const timer = argStr.match(reg2) && argStr.match(reg2)[1] //20221101,20221116,2,1,1
                const newCode =argStr.replace(reg2,"timer")
                console.log(newCode);
                const newCodeArr = newCode.split(',')
                setTimeout(()=>{
                    getAddrValues((val:any)=>{//获取codesList
                        if(val){
                            funAnalysis.codesList = val.toString().split(',') 
                            funAnalysis.timer = timer.split(',')
                            if(newCodeArr[1] == 'timer'){
                                funAnalysis.functionArg = ''
                                funAnalysis.nullValue = newCodeArr[2]
                                funAnalysis.axis = newCodeArr[3]
                                dialog.messageChild(JSON.stringify({type:7,list:funAnalysis}));
                            }else{
                                funAnalysis.nullValue = newCodeArr[3]
                                funAnalysis.axis = newCodeArr[4]
                                getAddrValues((value:any)=>{
                                    funAnalysis.functionArg = value.toString()
                                    dialog.messageChild(JSON.stringify({type:7,list:funAnalysis}));
                                },newCodeArr[1])
                            }
                            console.log(funAnalysis);
                        }
                    },newCodeArr[0])
                },1000)

            }          
            if(type){
                //funCodeUpdate例 =CCX.CCX_STOCK_SPECIFIC_RETURNLIST(A3:A302,"20221101,20221116,2,1,1",1)
                if(funCodeUpdate != null){
                    getFunUpdate()
                }else{
                    const timerGetValue = setInterval(()=>{
                        if(funCodeUpdate != null){
                            clearInterval(timerGetValue)
                            getFunUpdate()
                        }
                    },500)
                }
            }
            setTimeout(() => {
                run(1)
            }, 1000)
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, async function(arg:{message:string}){
                console.log(arg);
                if(arg.message == 'close') {
                    dialog.close()
                    dialog = null
                    return
                }
                if(arg.message == 'start') {
                    run(2)
                    timer = setInterval(()=>{
                        run(2)
                    },1000)
                    return
                }
                if (arg.message === 'stop') {//停止选区地址
                    timer && clearInterval(timer)
                    //3获取区域地质
                    getAddrValues((value:any[])=>{
                        //区域值
                        dialog.messageChild(JSON.stringify({type:4,value}));
                    },excAddrArea)
                    return
                }
                if(arg.message == 'startEnd') {
                    run(3)
                    timer = setInterval(()=>{
                        run(3)
                    },1000)
                    return
                }
                if (arg.message === 'againGetAddr') {//重新选择插入地址
                    timer = setInterval(()=>{
                        run(1)
                    },1000)
                    return
                }                
                if (arg.message === 'stopEnd') {//停止选区地址
                    timer && clearInterval(timer)
                    return
                }
                timer && clearInterval(timer);
                const obj = JSON.parse(arg.message)
                const { funCodeList,codeList,otherOptions } = obj
                const { startTime,endTime,direction,dataType,frequency,dataFormat,nullValue,axis,insertPosition} = otherOptions
                console.log(obj);
                console.log(excAddr);
                //获取交易日
                let timeListLength:number = 0//日期长度
                if(direction + axis == 3){
                    const timeList =await window.getTime(startTime,endTime,dataType,frequency)
                    timeListLength = timeList.data.data.length
                }
                let getAddStr:string = (excAddr.match(reg) as any)[0];//拿到地址的字母A1
                let getAddrNum:string = excAddr.replace(/[^\d.]/g, '')//拿到地址的数字 A1
                let addrStr:string = getAddStr
                let addrNum:string = getAddrNum
                if(axis == 1){
                    if(getAddStr == 'A' || getAddStr == 'B'){
                        addrStr = 'C'
                    }
                    if(Number(getAddrNum) <= 2){
                        addrNum = '3'
                    }                    
                }else{
                    if(getAddStr == 'A'){
                        addrStr = 'B'
                    }
                    if(Number(getAddrNum) <= 3){
                        addrNum = '4'
                    }  
                }

                funCodeList.forEach((item:any,index:number)=>{
                    if(index === 0){
                        item.addr = addrStr+addrNum
                    }else{
                        if(direction == 1 && axis == 1){
                            item.addr = `${addrStr}${Number(addrNum) + (codeList.length + 3)*index}`
                        }else if(direction == 1 && axis == 2){
                            item.addr = `${addrStr}${Number(addrNum) + (timeListLength + 4)*index}`
                        }else if(direction == 2 && axis == 1){
                            item.addr = `${transformToStr(strToNamber(addrStr) + (timeListLength + 3)*index) }${addrNum}`
                        }else if(direction == 2 && axis == 2){
                            item.addr = `${transformToStr(strToNamber(addrStr) + (codeList.length + 2)*index) }${addrNum}`
                        }
                    }
                })  
                console.log(funCodeList);
                //创建新表
                if(insertPosition === 2 && !type ){
                    try {
                        await Excel.run(async (context) => {
                            let sheets = context.workbook.worksheets;
                            sheets.load("items/name");
                            await context.sync();
                            const num = sheets.items.filter( sheet=> {
                                return sheet.name.includes('中诚信指数')
                            }).length;
                            let sheet = sheets.add("中诚信指数"+num);
                            sheet.load("name, position");
                            await context.sync();
                            let sheetNew = context.workbook.worksheets.getItem("中诚信指数"+num);
                            sheetNew.activate();
                            sheetNew.load("name");
                            await context.sync();
                        });                          
                    } catch (error) {
                        window.alertErr('已存在具有相同名称或标识符的资源,现已插入当前页')
                    }
                  
                }

                //end
                funCodeList.map((item:any)=>{
                    const argList = item.argList.filter((item:any)=>{
                        return item.arg
                    })
                    setDataForExcel(item.addr,codeList,{code:item.code,argList,name:item.name},startTime,endTime,dataType,frequency,dataFormat,nullValue,axis)
                })
                
                dialog.close()
                dialog = null
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, function (err:any) {//错误或点击关闭
                console.log(err);
                dialog.close()
                dialog = null
            });
        });

}
function funUpdate(event:any){
    getData(event,'update')
}
// Register the function with Office.
Office.actions.associate("funSearch", funSearch);
Office.actions.associate("logout", logout);
Office.actions.associate("getData", getData);
Office.actions.associate("funUpdate", funUpdate);