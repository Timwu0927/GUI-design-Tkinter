import win32com.client
import docx 
import tkinter as tk

def outlookrun(mail):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# 收件箱文件夹
    # inbox = outlook.Folders("edison-hu@nec.com.tw").Folders("收件匣")  
    inbox = outlook.Folders(mail).Folders("收件匣")    
    # 接收邮件，参数False代表不把收取进度显示出来，若要显示改为True即可
    outlook.SendAndReceive(True)
    # 获取收件箱文件夹里面的邮件对象(所有)
    messages = inbox.Items
    # 收件箱邮件总数
    count = len(messages)
    #print(messages)
    print(count)
    title=[]
    all_data=[]
    for test in messages:
        try:
            if '[PSA公告]' in test.Subject  and ('FW' not in test.Subject ) and ('RE:' not in test.Subject ):

                print('----------------------')
                title=test.Subject
                print(title)
                body=test.Body.split('Best Regard')
                print(body[0].strip())
                print('----------------------')
                all_data.append({'title':title,'content':body[0].strip()})
        except:
            print('失敗')
    return all_data