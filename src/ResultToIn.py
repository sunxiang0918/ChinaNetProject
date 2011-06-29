#-*- coding:cp936 -*-
import logging,sys
import xlrd
import xlwt


######################################
#该工具通过华为导出应用列表，自动筛选brew应用，生成添加真机excel和点brew功能审核列表
######################################
logger = logging.getLogger('Mylog')#初始化调试输出日志对象
formatter = logging.Formatter('[%(asctime)s][%(levelname)s] %(message)s')
runtimelog = logging.FileHandler("debug.log")
runtimelog.setFormatter(formatter)
logger.addHandler(runtimelog)

# 写屏功能，如不需要，则请注释下面三行
stdoutlog = logging.StreamHandler(sys.stdout)
stdoutlog.setFormatter(formatter)
logger.addHandler(stdoutlog)
logger.setLevel(logging.DEBUG) # DEBUG, INFO, WARNING, ERROR, CRITICAL ...etc

#########################写EXCEL
nol1=u'应用ID'
nol2=u'附件ID'
nol3=u'是否携带广告'
nol4=u'真机适配：1;手动理论适配：2'
nol5=u'真机适配机型（华为平台机型ID）'
nol55=u'操作类型'
nol6=u'测试结论'
nol7=u'不通过时跳转至（功能审核不通过：0,待版权确认：1,待内容审核：2)'
nol8=u'备注（不通过原因）'
nol9=u'评测得分'

importname = u'测试列表.xls'
wrname=u'测试结果导入.xls'
wtname=u'真机适配机型导入.xls'

ab = {}#key为附件ID，value是一个list
list = []#存放附件ID的list

attid = ''
appid = ''
result=''
comment=''
tid=''
grade=''
adv=''
testresult=u'不通过'


excel = xlrd.open_workbook(importname)#打开excel
sheet = excel.sheet_by_name('Sheet1')

for r in range(sheet.nrows-1):
	appid = sheet.row(r+1)[0].value
	attid = sheet.row(r+1)[1].value
	result = sheet.row(r+1)[2].value
	comment = sheet.row(r+1)[4].value
	tid = sheet.row(r+1)[7].value
	grade = sheet.row(r+1)[3].value
	adv = sheet.row(r+1)[6].value
	if len(attid)>7: attid = attid[5:]
	#list.append(attid)  是否需要单独保存key为list
	tlist = []#临时list，构建appid,attid,result,comment,tid,grade,adv
	maplist = []#临时list，构建多个【appid,attid,result,comment,tid,grade,adv】
	tmplist = []
	tlist.append(appid)
	tlist.append(attid)
	tlist.append(result)
	tlist.append(comment)
	tlist.append(tid)
	tlist.append(grade)
	tlist.append(adv)
	#logger.debug('~~~~~~~~~')	
	if ab.get(attid) != None : 
		tmplist = ab.get(attid)
		#logger.debug('aaaaaaaaaaa')
		#logger.debug(attid)
		#logger.debug(ab.get(attid))		
		#logger.debug('YES')
	tmplist.append(tlist)
	#logger.debug(tlist)
	ab[attid] = tmplist 
	#logger.debug(ab[attid])	
#print ab.get(list[0])[1].encode('cp936')
#logger.debug(ab)

file = xlwt.Workbook()#真机适配机型导入
table = file.add_sheet('Sheet1')
table.write(0,0,nol1)
table.write(0,1,nol2)
table.write(0,2,nol3)
table.write(0,3,nol4)
table.write(0,4,nol5)
table.write(0,5,nol55)

file1 = xlwt.Workbook()#测试结果导入
table1 = file1.add_sheet('Sheet1')
table1.write(0,0,nol1)
table1.write(0,1,nol2)
table1.write(0,2,nol6)
table1.write(0,3,nol7)
table1.write(0,4,nol8)
table1.write(0,5,nol9)

row = 1
row1 = 1

for key in ab:
	testresult = u'不通过'
	rdirect = '0'
	avg = 0
	tmp = ab[key]
	x = 1 #一个附件通过的任务数
	for t in range(len(tmp)):
		if tmp[t][2]==u'功能测试通过':			
			testresult = u'通过'
			rdirect = ''
			#logger.debug(u'通过')
			table.write(row,0,tmp[t][0])
			table.write(row,1,key)
			table.write(row,2,tmp[t][6])
			table.write(row,3,'1')
			table.write(row,4,tmp[t][4])
			table.write(row,5,'1')
			row = row + 1
			x = x + 1
			if tmp[t][5]=='':tmp[t][5]=0
			avg = tmp[t][5] + avg
		else :
			rdirect = ''
			table.write(row,0,tmp[t][0])
			table.write(row,1,key)
			table.write(row,2,tmp[t][6])
			table.write(row,3,'1')
			table.write(row,4,tmp[t][4])
			table.write(row,5,'2')
			row = row + 1
	table1.write(row1,0,tmp[t][0])
	table1.write(row1,1,key)
	table1.write(row1,2,testresult)
	table1.write(row1,3,rdirect)
	if testresult == u'不通过':
		table1.write(row1,4,tmp[0][3])
	#table1.write(row1,4,'')		
	table1.write(row1,5,avg/x)#求平均分
	row1 = row1 + 1

		
file.save(wtname)
file1.save(wrname)
print row1-1