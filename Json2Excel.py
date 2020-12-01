#!/usr/bin/python3
import requests
import json
import urllib3
import tablib


# json.text文件的格式： [{"a":1},{"a":2},{"a":3},{"a":4},{"a":5}]
# 获取json数据
def json2Excel(list_):
    header = tuple([i for i in list_[0].keys()])
    data = []
    for row in list_:
        body = []
        for v in row.values():
            body.append(v)
        data.append(tuple(body))
    data = tablib.Dataset(*data, headers=header)
    open('D:\\data.xls', 'wb').write(data.xls)


def POST(url, cookie):
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    post_data = {
        "REQUEST": {
            "API_ATTRS": {
                "Api_ID": "snowbeer.dmp.OCMS.getDeferredTransferIncome",
                "Api_Version": "1.0",
                "App_Sub_ID": "1200000702PG",
                "App_Token": "804a43506bc1423eb28eb548acae5608",
                "Partner_ID": "12000000",
                "Sign": "NO_SIGN",
                "Sys_ID": "12000007",
                "Time_Stamp": "2020-11-30 15:04:43:876",
                "User_Token": ""
            },
            "REQUEST_DATA": {
                "billCode": "HP20200720000606",
                "billDateEnd": "2020-07-31",
                "billDateStart": "2020-07-01",
                "orgCodeList": ["BZJ00"]
            }
        }
    }
    # post 请求头
    headers = {
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Mobile Safari/537.36",
        "Cookie": cookie}

    req = requests.post(url, json=post_data, headers=headers, verify=False)
    data = json.loads(req.text)
    response_ = data['RESPONSE']
    data_ = response_['RETURN_DATA']
    value_ = data_['value']
    # print(req.text)
    list_ = value_[0]['billDetailApiList']
    print(list_)
    return list_


if __name__ == '__main__':
    list_ = POST('http://uatocmsapi.crb.cn/crb-third-party-api-sec/invoiceApi/getFinishedWine', 'e342W3WeJH423rr')
    # 转excel
    json2Excel(list_)
