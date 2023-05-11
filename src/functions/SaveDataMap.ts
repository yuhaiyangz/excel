/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-11-07 11:22:53
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2022-11-07 11:34:41
 * @Description: 设置Map缓存大小
 */
export default class SaveDataMap{
    size :number
    data:Map<string,string>
    constructor(n:number){
        this.size = n; // 初始化最大缓存数据条数n
        this.data = new Map(); // 初始化缓存空间map
    }
    // 第二步代码
    put(domain:string, info:string):void{
        // if(this.data.has(domain)){
        //     this.data.delete(domain); // 移除数据
        //     this.data.set(domain, info)// 在末尾重新插入数据
        //     return;
        // }
        if(this.data.size >= this.size) {
            // 删除最不常用数据
            const firstKey= this.data.keys().next().value; // 迭代
            this.data.delete(firstKey);
        }
        this.data.set(domain, info) // 写入数据
    }
    //获取值
    getVal(key:string):string|null{
        if(this.data.has(key)){
            return (this.data.get(key) as string)
        }else{
            return null
        }
    }
}