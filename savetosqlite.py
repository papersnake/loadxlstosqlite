#-*- coding: UTF-8 -*-
import xlrd
import sqlite3
import pprint
import os,sys

'''
商品id 					itemid 				VARCHAR(12)
商品名称				itemname 			NVARCHAR(255)
商品主图				itempic 			VARCHAR(255)
商品详情页链接地址		itemdesclink 		VARCHAR(255)
商品一级类目			itemcats 			VARCHAR(100)
淘宝客链接				taobaokelink        VARCHAR(255)
商品价格(单位：元)		itemprice 			INTEGER /100
商品月销量				itemsellcount 		INTEGER
收入比率(%) 			srrate  			INTEGER
佣金					commission 			INTEGER
卖家旺旺				sellerww 			NVARCHAR(100)
卖家id 					sellerid 			NVARCHAR(100)
店铺名称				shopname 			NVARCHAR(100)
平台类型				platform 			CHAR(4)
优惠券id 				couponid 			CHAR(32)
优惠券总量				coupontotal 		INTEGER
优惠券剩余量			couponamount 		INTEGER
优惠券面额				coupondeno 			NVARCHAR(100)
优惠券开始时间			couponstarttime 	TEXT
优惠券结束时间			couponendtime 		TEXT
优惠券链接				couponlink 			VARCHAR(255)
商品优惠券推广链接		couponsharelink 	VARCHAR(255)
'''
def createdb(filename):
	conn = sqlite3.connect('test.db')
	print('ceate database successfully')
	cur = conn.cursor()  
	#sqlite if exists drop table
	cur.execute('DROP TABLE IF EXISTS tbkiteminfo;')
	cur.execute('''
		CREATE TABLE tbkiteminfo(
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			itemid 				VARCHAR(12),
			itemname 			NVARCHAR(255),
			itempic 			VARCHAR(255),
			itemdesclink 		VARCHAR(255),
			itemcats 			VARCHAR(100),
			taobaokelink        VARCHAR(255),
			itemprice 			INTEGER,
			itemsellcount 		INTEGER,
			srrate  			INTEGER,
			commission 			INTEGER,
			sellerww 			NVARCHAR(100),
			sellerid 			NVARCHAR(100),
			shopname 			NVARCHAR(100),
			platform 			CHAR(4),
			couponid 			CHAR(32),
			coupontotal 		INTEGER,
			couponamount 		INTEGER,
			coupondeno 			NVARCHAR(100),
			couponstarttime 	TEXT,
			couponendtime 		TEXT,
			couponlink 			VARCHAR(255),
			couponsharelink 	VARCHAR(255)
			);''')
	print 'create database table successfully'
	print 'start loading....'
	savetodb(filename,conn,cur);
	print 'successfully load excel file to sqlite db'
	if cur:
		cur.close()
	if conn:
		conn.close();

def savetodb(filename,conn,cur):
	file = xlrd.open_workbook(filename)

	sheet = file.sheet_by_index(0)

	nrows = sheet.nrows

	ncols = sheet.ncols

	col_names = []

	param=[];
	for i in xrange(1, nrows):
		cell=[]
		for j in xrange(0, ncols):
			cell.append(sheet.cell(i,j).value)
			#cell.append(str(j))
			#tmp=tuple(cell)
		param.append(cell)
	#print(param)
	try:
		sql='INSERT INTO tbkiteminfo(itemid,itemname,itempic,itemdesclink,itemcats,taobaokelink,itemprice,itemsellcount,srrate,commission,sellerww,sellerid,shopname,platform,couponid,coupontotal,couponamount,coupondeno,couponstarttime,couponendtime,couponlink,couponsharelink) '
		sql+='values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
		#sql='INSERT INTO tbkiteminfo(itemid,itemname) '
		#sql+='values(?,?)'
		#print param
		cur.executemany(sql,param)
		conn.commit()
	except Exception as e:
		print e
		#pprint.pprint(dump(cur))
		conn.rollback()

def listallfile():
	path = sys.path[0]
	xlsfiles=[]
	for filename in os.listdir(path):
		try:
			suff=filename.split('.')[1]
			suff=suff.strip().lower()
			if(suff=='xls'):
				xlsfiles.append(filename)
		except:
			continue

	#print xlsfiles
	return xlsfiles

def choosefile(filelist):
	filenum=len(filelist)
	print filelist[1]
	if filenum>1 :
		for i in xrange(0,filenum):
			print str(i) +':'+filelist[i]

		str_input=u"请选择要导入的文件标号:"
		str_input=str_input.encode('gbk')
		user_input=raw_input(str_input)
		try:
			#print filelist[int(user_input)]
			return filelist[int(user_input)]
		except Exception as e:
			print e
	else:
		print #filelist[0]
		return filelist[0]


if __name__=='__main__':
	filelist=listallfile()
	choosedfile=choosefile(filelist)

	print choosedfile

	print 'start load to database.'
	createdb(choosedfile)
	
	#createdb()