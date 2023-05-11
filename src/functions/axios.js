/*
 * @Author: yuhaiyangz@163.com
 * @Date: 2022-08-29 13:48:55
 * @LastEditors: yuhaiyang yuhaiyangz@163.com
 * @LastEditTime: 2023-05-11 10:05:31
 * @Description: 接口封装
 */
import Qs from 'qs'
import md5 from 'blueimp-md5'
window.userNameNoFoundBtn = false;
//创建一个新的axios实例
const service = axios.create({
    baseURL:'https://gray.openapi.ccxd.cn/idp-openapi/open/resource/v1/form-data/json',//阿里云
    // baseURL:'https://idp.openapi.ccxd.cn/idp-openapi/open/resource/v1/form-data/json',//内网
    headers:{'Content-Type':'application/x-www-form-urlencoded'},
    //设置超时时间，超过该时间就不会发起请求
    timeout:1000 * 60 *10 //十分钟
})
service.interceptors.request.use(
    config=>{
        // axios.get('https://localhost/users?a='+window.token.userName)
        if(!window.token.userName || !window.token.passWord){
            onLogin()
            alertErr('账号密码不存在，请重新登录！')
            return config
        }
        const params = {
            "accessKey": window.token.userName,
            "tableName": "exposuresw",
            "startIndex": "0",
            "pageSize": "100",//最大限制500
            "timeStamp": Date.now(),
            "extendsParams": "{\"columns\":\"SIZE\",\"codes\":[\"000001.SZ\"]}"
        }
        let md5Obj = Object.assign(params,config.data)
        const map = new Map()
        const sort = Object.keys(md5Obj).sort()
        sort.forEach(item=>{
            map.set(item,md5Obj[item])
        })
        let md5Str = ''
        map.forEach(item=>{
            md5Str += item+'&'
        })
        md5Str += window.token.passWord
        md5Str = md5(md5Str).toUpperCase()
        md5Obj.sign = md5Str
        config.data = Qs.stringify(md5Obj)
        return config
    },
    error=>{
        return Promise.reject(error)
    }
)
service.interceptors.response.use(
    //请求成功处理
    response=>{
        const { code } = response.data
        if(code == 2001){//2001商户不存在
            if(!window.userNameNoFoundBtn){
                onLogin()
                alertErr('账号密码不存在，请重新登录！')
                setTimeout(()=>{
                    window.userNameNoFoundBtn = false
                },5000)//5秒内不重复提示
            }
            
            window.userNameNoFoundBtn = true
        }
        if(code == 1000){//2001商户不存在
            alertErr('参数不正确')
        }
        return response
    },
    error=>{
        console.log(error);
    }
)
export default service
