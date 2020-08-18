import pyodbc
import configparser

config = configparser.ConfigParser()
config.read('Information_DB.ini',encoding = 'utf8')#因為只有進入點那三隻需要用到這個PY 所以路徑要設 對那三隻來說的相對路徑
try:

   conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='\
      +config['database']['server']+';DATABASE='+config['database']['database']\
         +';UID='+config['database']['username']+';PWD='+ config['database']['password'])
   cursor = conn.cursor()
   print("資料庫連接成功")
except:
   print("資料庫連接失敗")


def InsertToSQL(ModuleID,CreatedDate,Title,MoreLink,Description):
	try:
		cursor = conn.execute("insert into dbo.Portal_Announcements(ModuleID,CreatedByUser,CreatedDate,\
		Title,MoreLink,MobileMoreLink,ExpireDate,Description) values(?,?,?,?,?,?,?,?)",ModuleID,\
		'AutoMationProgram',CreatedDate,Title,\
		MoreLink,' '\
		,'2099-12-31 00:00:00.000',Description)
		conn.commit()
		print('InsertToSQL 成功')
	except:
		print('InsertToSQL失敗!')

def LetArrayToSql(ModuleID,Array):
	try:
		for index in Array:
			print('----------------------')
			print('#標題:{} \n#內文：{} \n#發文日期：{}\n#連結：{}'.format(index['title'], index['content'],index['date'],index['link']))
			InsertToSQL(ModuleID,index['date'],index['title'],index['link'],index['content'])
			print('LetArrayToSql 成功')
			print('----------------------')
	except:
		print('LetArrayToSql and InsertToSQL Fail')

def SeleteArrayFromSQL(command):
	cursor.execute(command) 
	row = cursor.fetchall()
	conn.commit()
	return row









