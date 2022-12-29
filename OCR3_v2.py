# encoding:utf-8
from pickle import APPEND
import requests
import base64
import os
import xlwt
import json
import pandas as pd
import re
import openpyxl

#百度OCR API
def API():
    client_id='B5DWHgZP7raTdgwW01vMtQfv'
    client_secret='sfPRn3UVf9gPTRiDdPfZuUxdna2pwCYu'
    url= f'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={client_id}&client_secret={client_secret}'
    response = requests.get(url)
    if response:
        access_token=(response.json()['access_token'])
    return access_token

def pics(path):
    print('正在生成图片路径')
    #生成一个空列表用于存放图片路径
    pics = []
    # 遍历文件夹，找到后缀为jpg和png的文件，整理之后加入列表
    for filename in os.listdir(path):
        if filename.endswith('') or filename.endswith('png'):
            pic = path + '/' + filename
            pics.append(pic)
    print('图片路径生成成功！')
    return pics

# 获取发票正文内容
def get_context(pic):
    # print('正在获取图片正文内容！')
    global df
    ListB = []
    data = {}
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice"
        # 二进制方式打开图片文件
    f = open(pic, 'rb')
    img = base64.b64encode(f.read())
    params = {"image":img}
    request_url = request_url + "?access_token=" + API()
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    response = requests.post(request_url, data=params, headers=headers)
    
    json1 = response.json()
    print (json1)
    number=json1['words_result']['CommodityNum']
    item_No = 0
    for i in number:
        item_No = item_No+1
        #print (json1['words_result']['SellerName'])
    for k in range(0,item_No):
        try:
            ListA=[json1['words_result']['InvoiceDate'],
            json1['words_result']['InvoiceNum'],
            json1['words_result']['SellerName'],
            json1['words_result']['PurchaserName'],
            json1['words_result']['CommodityName'][k]['word'],
            json1['words_result']['CommodityType'][k]['word'],
            json1['words_result']['CommodityUnit'][k]['word'],
            json1['words_result']['CommodityNum'][k]['word'],
            json1['words_result']['CommodityPrice'][k]['word'],
            json1['words_result']['CommodityAmount'][k]['word'],
            json1['words_result']['CommodityTaxRate'][k]['word'],
            json1['words_result']['CommodityTax'][k]['word']]
        except Exception as e:
            try:
                ListA=[json1['words_result']['InvoiceDate'],
                json1['words_result']['InvoiceNum'],
                json1['words_result']['SellerName'],
                json1['words_result']['PurchaserName'],"N/A",
                json1['words_result']['CommodityType'][k]['word'],
                json1['words_result']['CommodityUnit'][k]['word'],
                json1['words_result']['CommodityNum'][k]['word'],
                json1['words_result']['CommodityPrice'][k]['word'],
                json1['words_result']['CommodityAmount'][k]['word'],
                json1['words_result']['CommodityTaxRate'][k]['word'],
                json1['words_result']['CommodityTax'][k]['word']]
            except Exception as e:
                try:
                    ListA=[json1['words_result']['InvoiceDate'],
                    json1['words_result']['InvoiceNum'],
                    json1['words_result']['SellerName'],
                    json1['words_result']['PurchaserName'],"N/A",
                    json1['words_result']['CommodityType'][k]['word'],
                    json1['words_result']['CommodityUnit'][k]['word'],
                    json1['words_result']['CommodityNum'][k]['word'],
                    json1['words_result']['CommodityPrice'][k]['word'],
                    json1['words_result']['CommodityAmount'][k]['word'],
                    json1['words_result']['CommodityTaxRate'][k]['word'],
                    json1['words_result']['CommodityTax'][k]['word']]
                except Exception as e:
                        try:
                            ListA=[json1['words_result']['InvoiceDate'],
                            json1['words_result']['InvoiceNum'],
                            json1['words_result']['SellerName'],
                            json1['words_result']['PurchaserName'],
                            json1['words_result']['CommodityName'][k]['word'],"N/A",
                            json1['words_result']['CommodityUnit'][k]['word'],
                            json1['words_result']['CommodityNum'][k]['word'],
                            json1['words_result']['CommodityPrice'][k]['word'],
                            json1['words_result']['CommodityAmount'][k]['word'],
                            json1['words_result']['CommodityTaxRate'][k]['word'],
                            json1['words_result']['CommodityTax'][k]['word']]
                        except Exception as e:
                            try:
                                ListA=[json1['words_result']['InvoiceDate'],
                                json1['words_result']['InvoiceNum'],
                                json1['words_result']['SellerName'],
                                json1['words_result']['PurchaserName'],
                                json1['words_result']['CommodityName'][k]['word'],
                                json1['words_result']['CommodityType'][k]['word'],"N/A",
                                json1['words_result']['CommodityNum'][k]['word'],
                                json1['words_result']['CommodityPrice'][k]['word'],
                                json1['words_result']['CommodityAmount'][k]['word'],
                                json1['words_result']['CommodityTaxRate'][k]['word'],
                                json1['words_result']['CommodityTax'][k]['word']]
                            except Exception as e:
                                pass
        ListB.append(ListA)
        df = pd.DataFrame(ListB, columns=['发票日期', '发票号码', '销售方名称','购买方名称','产品名称','产品型号','产品单位','产品数量','单价','金额','税率','税额']) 



def datas(pics):
    newDF=pd.DataFrame(columns=['发票日期', '发票号码', '销售方名称','购买方名称','产品名称','产品型号','产品单位','产品数量','单价','金额','税率','税额'], index=[])
    for p in pics:
        print(p)
        get_context(p)
        data=df
        print (data)
        newDF = pd.concat([newDF,data])
    newDF.to_excel('输出结果.xlsx', sheet_name='Sheet1', index=False)
    return 0





def main():
    print('开始执行！！！')
    # 发票的存放地址
    df = pd.DataFrame()
    path = os.path.abspath(os.curdir)+'/上传发票/'
    Pics = pics(path)
    datas(Pics)
    print("输出完成")

if __name__ == '__main__':
    main()
