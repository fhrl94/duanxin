from urllib.parse import  quote
from yunpian_python_sdk.model import constant as YC
from yunpian_python_sdk.ypclient import YunpianClient
# 初始化client,apikey作为所有请求的默认值

def array_quote(array):
    datastr=[]
    if not isinstance(array, list):
        print("错误")
        quit(1)
    for one in array:
        datastr.append(quote(one))
    return datastr

clnt = YunpianClient('')

# 批量发送需要手机号码和内容的数量保持一致；多条使用+ "," + 来连接；发送内容需要使用quote编码

def sms_send(Tel, datastr):
    if not isinstance(Tel, list):
        print("程序错误")
        quit(1)
        raise Exception("程序错误")
    param = {YC.MOBILE:','.join(Tel),YC.TEXT:(','.join(array_quote(datastr)))}
    error=0
    if len(Tel) == 0:
        print("无人员")
        return error
    r = clnt.sms().multi_send(param)
    print(r.data())
    if r.data() != None:
        for one in r.data()['data']:
            if one.get('code',-1) != 0:
                error += 1
        return error
    print("程序逻辑错误")
    return 1
# 获取返回结果, 返回码:r.code(),返回码描述:r.msg(),API结果:r.data(),其他说明:r.detail(),调用异常:r.exception()
# 短信:clnt.sms() 账户:clnt.user() 签名:clnt.sign() 模版:clnt.tpl() 语音:clnt.voice() 流量:clnt.flow()


