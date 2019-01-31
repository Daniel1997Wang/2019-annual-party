#!/usr/bin/env python
# -*- coding: utf-8 -*-

# 1月21日晚年会食用
# 注私有信息请到腾讯云注册食用，以下全部以*表示

import xlrd as xlrd
from qcloudsms_py import SmsSingleSender
from qcloudsms_py.httpclient import HTTPError
import time


appid = "****"
appkey = "****"
template_id = "****"
sms_sign = "AIOT汉字识别"


GIFT = ["明信片","明信片","胶带","胶带","英语单词本","英语单词本","便利贴","便利贴","硬笔书法本"
    , "硬笔书法本","硬笔书法本","硬笔书法本","硬笔书法本","硬笔书法本","硬笔书法本","硬笔书法本"
        ,"硬笔书法本","硬笔书法本","硬笔书法本","硬笔书法本","无","无","无","无"]


# 幸运码
key_value = ["3377","6435","9959","9343","7133","4383","9167","4293",
             "8510","7277","2646","4016","1140","6598","6349","1038",
             "2760","2814","6144","5020","9932","7596","4706","2185"]


# 数据读取
def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(r'data.xls')
    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_name('Sheet1')
    # 获取整行和整列的值（数组）
    Row = []
    name = sheet1.col_values(6)
    id_number = sheet1.col_values(7)
    phone_number = sheet1.col_values(8)
    Key = sheet1.col_values(9)
    lucky_number = sheet1.col_values(10)
    Row.append(name)
    Row.append(id_number)
    Row.append(phone_number)
    Row.append(lucky_number)
    Row.append(Key)
    return Row


# 模拟抽奖
def get_data():
    f = open("Pi.txt","r")
    lines = f.readlines()#读取全部内容
    data = lines[0]
    data = data.replace("\n", "")
    f.close()
    return data


#输入的数据为1000-9999
def search(number):
    Pi = get_data()
    if str(number) in Pi:
        index = Pi.index(str(number))
        return index


# 得到获奖名单
def get_result():
    new_data = []
    data = read_excel()
    #print(data)
    lucky_number = data[3]
    key = data[4]
    #print(key)
    for i in range(1,len(data[0])):
        if str(int(key[i])) in key_value:
            temp = []
            number = int(lucky_number[i])
            index = search(number)
            temp.append(data[0][i])
            temp.append(int(data[1][i]))
            temp.append(int(data[2][i]))
            temp.append(number)
            temp.append(index)
            new_data.append(temp)
        else:
            print(data[0][i],end="")
            print(",你输入的手写字帖编码出错")
    return new_data


#得到最终获奖名单
def get_finally_name():
    phone_number, Message = [],[]
    data = get_result()
    res = sorted(data,key=lambda x:(x[4],x[1]))
    ans = []
    for i in range(len(res)):
        name = str(res[i][0])
        lucky_number = str(res[i][3])
        gift = GIFT[i]
        message = "您好，" + name + \
                  "！感谢参与手写汉字数据集的采集工作。根据您提交的抽奖号码(" + lucky_number + \
                  ")，赠送给您一份小礼品——" + gift + "。"
        phone_number.append(res[i][2])
        Message.append(message)

    return phone_number,Message


def main():
    phone_numbers,Message = get_finally_name()
    time.sleep(60)
    # 需要发送短信的手机号码
    print(phone_numbers,end="")
    for i in range(len(Message)):
        message = Message[i]
        print(message)
        sms_type = 0  # Enum{0: 普通短信, 1: 营销短信}
        ssender = SmsSingleSender(appid, appkey)
        try:
            result = ssender.send(sms_type, 86, phone_numbers[i],
                                  message, extend="", ext="")
        except HTTPError as e:
            print(e)
        except Exception as e:
            print(e)

        print(i+1,result)


if __name__ == '__main__':
    main()