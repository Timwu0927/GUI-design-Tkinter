#-*- coding:utf-8 -*-
import datecaculate
import os
from ftplib import FTP #載入ftp模組
import codecs
import configparser

config = configparser.ConfigParser()
config.read('Information_FTP.ini',encoding = 'utf8')


# config['database']['server']

ftp=FTP() #設定變數
ftp.encoding='Big5'
ftp.set_debuglevel(2) #開啟除錯級別2，顯示詳細資訊
ftp.connect(config['FTP']['server'],int(config['FTP']['port'])) #連線的ftp sever和port
ftp.login(config['FTP']['username'],config['FTP']['password'])
print(ftp.getwelcome)
print('ftp連結成功')


def ChooseFtpPath(pathname):
	
	ftp.cwd(pathname)  # 设置FTP路径
	printlist = ftp.nlst()       # 获得目录列表
	for name in printlist:
	    print(name)   




def UploadToFtp(localfilepath,filename):
	
	
	# localfile = "C:\\For-MS-Crawler\\微軟安全性更新公告-{}\\安全更新.xlsx".format(today) #設定檔案位置 (local路徑\)
	localfile = localfilepath #設定檔案位置 (local路徑\)
	print(localfile)


	f = open( localfile, 'rb')        #開啟檔案     
	# ftp.storbinary('STOR %s'% 'MicroSoftSecurity-{}.xlsx'.format(today), f) #上傳檔案
	ftp.storbinary('STOR %s'% filename, f) #上傳檔案












