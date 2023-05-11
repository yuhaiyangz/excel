/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-09-05 13:36:30
 * @LastEditors: Please set LastEditors
 * @LastEditTime: 2023-01-10 17:01:00
 * @Description: 请填写简介
 */
// 这个不能加export 加了就识别不到cityData了
declare interface Window {
    token: {
        userName:string;
        passWord:string
    },
    ifLoadSuccess:boolean,
    onLogin:Function,
    alertErr:Function,
    onLogout:Function,
    getTime:Function,
    getCodeListSZ:Function,
    getCodeListOF:Function,
    getCodeListIndex:Function,
    saveData:Map<string,string>,
    nullValue:number,
    enableButton:Function
  }
  declare const Vue :any;
  declare const ElementPlus :any;
  declare const axios :any;  
