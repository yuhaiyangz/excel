/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-11-11 16:53:17
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2022-12-27 10:30:55
 * @Description: 公用函数
 */
export function transformToStr(i:number):string {//excel单元格数字转AAA
    var s = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z";
    var sArray=s.split(" ");
    const map = new Map()
    sArray.map((item,index)=>{
        map.set(index+1,item)
    })
    let ArrStr:number[] = []
    
    function setArr(num:number,arr:number[],i:number):void{
        const shi = Math.floor(num/26);//十位
        const ge = num % 26;//各位
        arr[i] = ge;
        if(shi > 26){
            setArr(shi,arr,++i)
        }else{
            if(ge === 0){
                arr[i+1] = shi - 1;
            }else{
                arr[i+1] = shi;
            }
        }
    }
    setArr(i,ArrStr,0)
    let result:string = ''
    ArrStr.reverse().map((item,i)=>{
        if(i === 0 && item === 0){
            result +=''
        }else{
            if(map.get(item)){
                result += map.get(item)
            }else{
                result +='Z'
            }            
        }
    })
    return result
}
export function strToNamber(str:string):number {//excel单元格AAA转数字
    var s = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z";
    var sArray=s.split(" ");
    const map = new Map()
    sArray.map((item,index)=>{
        map.set(item,index+1)
    })
    if (str.length == 1){
        return map.get(str)
    }else{
        const strList=str.split("");//AA
        let num = 0 
        strList.forEach((item,index)=>{
            if(strList.length - 1 === index){
                num+=map.get(item)
            }else{
                // num+= map.get(item)*(strList.length - index - 1)*26
                num+= map.get(item)*Math.pow(26,strList.length - index - 1)
            }
        })
        return num
    }
}
export const reg = /^([A-Z]+)/g;
export const index_rebalance_frequencyList = ['日度','月度','季度','半年度','年度']
export const index_weight_methodList = ['','等权','市值加权','市值加权+规模权限','动态加权','其他']
//日期格式花化
export function formatDate(numb: number | string): string|number {
    // const time = new Date((numb - 1) * 24 * 3600000 + 1)
    const type = typeof numb
    if (type == 'string') return numb;
    if (numb.toString().length === 8) return numb
    let time = new Date(1900, 0, Number(numb) - 1)
    // time.setYear(time.getFullYear() - 70)
    const year = time.getFullYear()
    const month = time.getMonth() + 1
    const date = time.getDate()
    return year + '' + (month < 10 ? '0' + month : month) + '' + (date < 10 ? '0' + date : date)
  }