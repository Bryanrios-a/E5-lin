# -*- coding: UTF-8 -*-
import os
import xlsxwriter
import requests as req
import json,sys,time,random

emailaddress=os.getenv('EMAIL')
app_num=os.getenv('APP_NUM')

config = {
         'allstart': 0,
         'rounds': 1,
         'rounds_delay': [0,0,5],
         'api_delay': [0,0,5],
         'app_delay': [0,0,5],
         }        
if app_num == '' or app_num is None:
    app_num = '1'

city=os.getenv('CITY')
if city == '' or city is None:
    city = 'Beijing'
access_token_list=['dummy']*int(app_num)

# 微软 refresh_token 获取
def getmstoken(ms_token,appnum):
    headers={'Content-Type':'application/x-www-form-urlencoded'}
    data={'grant_type': 'refresh_token',
        'refresh_token': ms_token,
        'client_id':client_id,
        'client_secret':client_secret,
        'redirect_uri':'http://localhost:53682/'
        }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    if 'refresh_token' in jsontxt:
        print(r'账号/应用 '+str(appnum)+' 的微软密钥获取成功')
    else:
        print(r'账号/应用 '+str(appnum)+' 的微软密钥获取失败')
    refresh_token = jsontxt.get('refresh_token','')
    access_token = jsontxt.get('access_token','')
    return access_token

# api 延时
def apiDelay():
    if config['api_delay'][0] == 1:
        time.sleep(random.randint(config['api_delay'][1],config['api_delay'][2]))
        
def apiReq(method,a,url,data='QAQ'):
    apiDelay()
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    if method == 'post':
        posttext=req.post(url,headers=headers,data=data)
    elif method == 'put':
        posttext=req.put(url,headers=headers,data=data)
    elif method == 'delete':
        posttext=req.delete(url,headers=headers)
    else:
        posttext=req.get(url,headers=headers)
    if posttext.status_code < 300:
        print('        操作成功')
    else:
        print('        操作失败: '+str(posttext.status_code))
    return posttext.text
          
# 上传文件到 onedrive
def UploadFile(a,filesname,f):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/content'
    apiReq('put',a,url,f)
    
# 发送邮件
def SendEmail(a,subject,content):
    url=r'https://graph.microsoft.com/v1.0/me/sendMail'
    mailmessage={'message': {'subject': subject,
                             'body': {'contentType': 'Text', 'content': content},
                             'toRecipients': [{'emailAddress': {'address': emailaddress}}]},
                 'saveToSentItems': 'true'}            
    apiReq('post',a,url,json.dumps(mailmessage))	

# excel 写入
def excelWrite(a,filesname,sheet):
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/add'
    data={"name": sheet}
    print('    添加工作表')
    apiReq('post',a,url,json.dumps(data))
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/worksheets/'+sheet+r'/tables/add'
    data={"address": "A1:D8","hasHeaders": False}
    print('    添加表格')
    jsontxt=json.loads(apiReq('post',a,url,json.dumps(data)))
    print('    添加行')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/AutoApi/App'+str(a)+r'/'+filesname+r':/workbook/tables/'+jsontxt['id']+r'/rows/add'
    rowsvalues=[[random.randint(1,1200) for _ in range(4)] for _ in range(2)]
    data={"values": rowsvalues}
    apiReq('post',a,url,json.dumps(data))

# To Do
def taskWrite(a,taskname):
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists'
    data={"displayName": taskname}
    print("    创建任务列表")
    listjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks'
    data={"title": taskname}
    print("    创建任务")
    taskjson=json.loads(apiReq('post',a,url,json.dumps(data)))
    # 删除任务 & 列表
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']+r'/tasks/'+taskjson['id']
    apiReq('delete',a,url)
    url=r'https://graph.microsoft.com/v1.0/me/todo/lists/'+listjson['id']
    apiReq('delete',a,url)    

# Teams
def teamWrite(a,channelname):
    url=r'https://graph.microsoft.com/v1.0/me/joinedTeams'
    print("    获取team")
    jsontxt = json.loads(apiReq('get',a,url))
    objectlist=jsontxt.get('value',[])
    if not objectlist: return
    print("    创建team频道")
    data={"displayName": channelname,"description": "demo","membershipType": "standard"}
    url=r'https://graph.microsoft.com/v1.0/teams/'+objectlist[0]['id']+r'/channels'
    jsontxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    url=r'https://graph.microsoft.com/v1.0/teams/'+objectlist[0]['id']+r'/channels/'+jsontxt['id']
    print("    删除team频道")
    apiReq('delete',a,url)

# OneNote
def onenoteWrite(a,notename):
    url=r'https://graph.microsoft.com/v1.0/me/onenote/notebooks'
    data={"displayName": notename}
    print('    创建笔记本')
    notetxt = json.loads(apiReq('post',a,url,json.dumps(data)))
    print('    删除笔记本')
    url=r'https://graph.microsoft.com/v1.0/me/drive/root:/Notebooks/'+notename
    apiReq('delete',a,url)

# 用户信息
def userInfo(a):
    url = r'https://graph.microsoft.com/v1.0/me'
    print("    获取用户信息")
    resp = apiReq('get', a, url)
    print(resp)

# 日历
def calendarWrite(a):
    url = r'https://graph.microsoft.com/v1.0/me/events'
    data = {"subject": "测试日历事件",
            "start": {"dateTime": "2025-10-01T09:00:00","timeZone": "UTC"},
            "end": {"dateTime": "2025-10-01T10:00:00","timeZone": "UTC"}}
    print("    创建日历事件")
    resp = apiReq('post', a, url, json.dumps(data))
    print(resp)

# 联系人
def contactWrite(a):
    url = r'https://graph.microsoft.com/v1.0/me/contacts'
    data = {"givenName": "Test","surname": "Contact","emailAddresses": [{"address": "test@example.com"}]}
    print("    创建联系人")
    resp = apiReq('post', a, url, json.dumps(data))
    print(resp)

# 群组
def groupList(a):
    url = r'https://graph.microsoft.com/v1.0/groups?$top=1'
    print("    获取群组")
    resp = apiReq('get', a, url)
    print(resp)

# 网站和列表
def siteList(a):
    url = r'https://graph.microsoft.com/v1.0/sites/root/sites'
    print("    获取网站列表")
    resp = apiReq('get', a, url)
    print(resp)

# Copilot Chat
def copilotChat(a, prompt):
    url = r'https://graph.microsoft.com/beta/copilot/chat'
    data = {"messages": [{"role": "user", "content": prompt}]}
    print("    调用Copilot对话")
    resp = apiReq('post', a, url, json.dumps(data))
    print(resp)

# 一次性获取 access_token
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)
print('')    

# 获取天气
headers={'Accept-Language': 'zh-CN'}
weather=req.get(r'http://wttr.in/'+city+r'?format=4&?m',headers=headers).text

# 实际运行：发邮件
for a in range(1, int(app_num)+1):
    print('账号 '+str(a))
    if emailaddress:
        SendEmail(a,'weather',weather)

# 定义 API 函数池
api_functions = [
    lambda a: excelWrite(a,'QVQ'+str(random.randint(1,600))+'.xlsx','Sheet'+str(random.randint(1,10))),
    lambda a: teamWrite(a,'Team'+str(random.randint(1,600))),
    lambda a: taskWrite(a,'Task'+str(random.randint(1,600))),
    lambda a: onenoteWrite(a,'Note'+str(random.randint(1,600))),
    lambda a: userInfo(a),
    lambda a: calendarWrite(a),
    lambda a: contactWrite(a),
    lambda a: groupList(a),
    lambda a: siteList(a),
    lambda a: copilotChat(a,"帮我写一首七言绝句")
]

# 随机执行 5 项
for _ in range(1,config['rounds']+1):
    if config['rounds_delay'][0] == 1:
        time.sleep(random.randint(config['rounds_delay'][1],config['rounds_delay'][2]))     
    print('第 '+str(_)+' 轮\n')        
    for a in range(1, int(app_num)+1):
        if config['app_delay'][0] == 1:
            time.sleep(random.randint(config['app_delay'][1],config['app_delay'][2]))        
        print('账号 '+str(a))    
        chosen = random.sample(api_functions, 5)
        for func in chosen:
            func(a)
        print('-')

