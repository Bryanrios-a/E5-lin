# -*- coding: UTF-8 -*-
import os
import requests as req
import json, sys, time, random

# 获取环境变量中的应用数量，默认为'1'
app_num = os.getenv('APP_NUM')
if app_num == '' or app_num is None:
    app_num = '1'

# 创建一个列表，存储access_token，初始值
access_token_list = ['placeholder'] * int(app_num)

###########################
# config选项说明
# 0：关闭  ， 1：开启  建议1
# api_rand：是否随机排序api （开启随机抽取12个，关闭默认初版10个）。默认1开启
# rounds: 轮数，即每次启动跑几轮。
# rounds_delay: 是否开启每轮之间的随机延时，后面两参数代表延时的区间。默认0关闭
# api_delay: 是否开启api之间的延时，默认0关闭
# app_delay: 是否开启账号之间的延时，默认0关闭
########################################
config = {
         'api_rand': 1,
         'rounds': 3,
         'rounds_delay': [1,60,120],
         'api_delay': [0,2,6],
         'app_delay': [0,30,60],
         }

# API列表（基础Graph API）
api_list = [
    r'https://graph.microsoft.com/v1.0/me/',
    r'https://graph.microsoft.com/v1.0/users',
    r'https://graph.microsoft.com/v1.0/me/people',
    r'https://graph.microsoft.com/v1.0/groups',
    r'https://graph.microsoft.com/v1.0/me/contacts',
    r'https://graph.microsoft.com/v1.0/me/drive/root',
    r'https://graph.microsoft.com/v1.0/me/drive/root/children',
    r'https://graph.microsoft.com/v1.0/drive/root',
    r'https://graph.microsoft.com/v1.0/me/drive',
    r'https://graph.microsoft.com/v1.0/me/drive/recent',
    r'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe',
    r'https://graph.microsoft.com/v1.0/me/calendars',
    r'https://graph.microsoft.com/v1.0/me/events',
    r'https://graph.microsoft.com/v1.0/sites/root',
    r'https://graph.microsoft.com/v1.0/sites/root/sites',
    r'https://graph.microsoft.com/v1.0/sites/root/drives',
    r'https://graph.microsoft.com/v1.0/sites/root/columns',
    r'https://graph.microsoft.com/v1.0/me/onenote/notebooks',
    r'https://graph.microsoft.com/v1.0/me/onenote/sections',
    r'https://graph.microsoft.com/v1.0/me/onenote/pages',
    r'https://graph.microsoft.com/v1.0/me/messages',
    r'https://graph.microsoft.com/v1.0/me/mailFolders',
    r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',
    r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    r"https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'high'",
    r'https://graph.microsoft.com/v1.0/me/messages?$search="hello world"',
    r'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top',
]

# Copilot 相关 API (beta)
copilot_api_list = [
    r'https://graph.microsoft.com/beta/copilot/completions',      # 通用 Copilot 补全
    r'https://graph.microsoft.com/beta/me/copilot/chat',          # 用户聊天
    r'https://graph.microsoft.com/beta/me/copilot/suggestions',   # Copilot 建议
]

# 合并 Copilot API
api_list.extend(copilot_api_list)

# 微软 refresh_token 获取
def getmstoken(ms_token, appnum):
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': ms_token,
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': 'http://localhost:53682/'
    }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',
                    data=data, headers=headers)
    jsontxt = json.loads(html.text)
    if 'refresh_token' in jsontxt:
        print(r'账号/应用 ' + str(appnum) + ' 的微软密钥获取成功')
    else:
        print(r'账号/应用 ' + str(appnum) + ' 的微软密钥获取失败' + '\n' +
              '请检查 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 是否正确')
    refresh_token = jsontxt.get('refresh_token')
    access_token = jsontxt.get('access_token')
    return access_token

# API 调用
def runapi(apilist, a):
    access_token = access_token_list[a-1]
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }
    for b in range(len(apilist)):
        url = api_list[apilist[b]]

        # 判断是否为 Copilot API
        if "copilot/completions" in url:
            payload = {
                "messages": [
                    {"role": "user", "content": "Hello Copilot, summarize Microsoft 365 briefly."}
                ]
            }
            res = req.post(url, headers=headers, json=payload)
        elif "copilot/chat" in url:
            payload = {"messages": [{"role": "user", "content": "Give me 3 productivity tips."}]}
            res = req.post(url, headers=headers, json=payload)
        elif "copilot/suggestions" in url:
            res = req.get(url, headers=headers)
        else:
            res = req.get(url, headers=headers)

        if res.status_code == 200:
            print('第' + str(apilist[b]) + "号api调用成功 -> " + url)
        else:
            print("调用失败 " + url + " 状态码: " + str(res.status_code))

        if config['api_delay'][0] == 1:
            time.sleep(random.randint(config['api_delay'][1], config['api_delay'][2]))

# 一次性获取 access_token
for a in range(1, int(app_num)+1):
    client_id = os.getenv('CLIENT_ID_' + str(a))
    client_secret = os.getenv('CLIENT_SECRET_' + str(a))
    ms_token = os.getenv('MS_TOKEN_' + str(a))
    access_token_list[a-1] = getmstoken(ms_token, a)

# 随机 api 序列
fixed_api = [0, 1, 5, 6, 20, 21]
ex_api = [2,3,4,7,8,9,10,22,23,24,25,26,27,13,14,15,16,17,18,19,11,12,28,29,30]  
# 增加了 Copilot 的 28,29,30
fixed_api.extend(random.sample(ex_api, 6))
random.shuffle(fixed_api)
final_list = fixed_api

# 实际运行
if int(app_num) > 1:
    print('多账户/应用模式下，日志里可能会出现一堆***，属于正常情况')
print("如果api数量少于规定值，可能是api赋权没弄好，或者onedrive未初始化。")
print('共 ' + str(app_num) + r' 账号/应用，每个账号/应用 ' + str(config['rounds']) + ' 轮')

for r in range(1, config['rounds']+1):
    if config['rounds_delay'][0] == 1:
        time.sleep(random.randint(config['rounds_delay'][1], config['rounds_delay'][2]))
    for a in range(1, int(app_num)+1):
        if config['app_delay'][0] == 1:
            time.sleep(random.randint(config['app_delay'][1], config['app_delay'][2]))
        print('\n应用/账号 ' + str(a) + ' 的第' + str(r) + '轮 ' +
              time.asctime(time.localtime(time.time())) + '\n')
        if config['api_rand'] == 1:
            print("已开启随机顺序,共十二个api,自己数")
            apilist = final_list
        else:
            print("原版顺序,共十个api,自己数")
            apilist = [5,9,8,1,20,24,23,6,21,22]
        runapi(apilist, a)
