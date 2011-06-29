#-*- coding:cp936 -*-
import logging,sys
import xlrd
import xlwt


######################################
#�ù���ͨ����Ϊ����Ӧ���б��Զ�ɸѡbrewӦ�ã�����������excel�͵�brew��������б�
######################################
logger = logging.getLogger('Mylog')#��ʼ�����������־����
formatter = logging.Formatter('[%(asctime)s][%(levelname)s] %(message)s')
runtimelog = logging.FileHandler("debug.log")
runtimelog.setFormatter(formatter)
logger.addHandler(runtimelog)

# д�����ܣ��粻��Ҫ������ע����������
stdoutlog = logging.StreamHandler(sys.stdout)
stdoutlog.setFormatter(formatter)
logger.addHandler(stdoutlog)
logger.setLevel(logging.DEBUG) # DEBUG, INFO, WARNING, ERROR, CRITICAL ...etc

#########################дEXCEL
nol1=u'Ӧ��ID'
nol2=u'����ID'
nol3=u'�Ƿ�Я�����'
nol4=u'������䣺1;�ֶ��������䣺2'
nol5=u'���������ͣ���Ϊƽ̨����ID��'
nol55=u'��������'
nol6=u'���Խ���'
nol7=u'��ͨ��ʱ��ת����������˲�ͨ����0,����Ȩȷ�ϣ�1,��������ˣ�2)'
nol8=u'��ע����ͨ��ԭ��'
nol9=u'����÷�'

importname = u'�����б�.xls'
wrname=u'���Խ������.xls'
wtname=u'���������͵���.xls'

ab = {}#keyΪ����ID��value��һ��list
list = []#��Ÿ���ID��list

attid = ''
appid = ''
result=''
comment=''
tid=''
grade=''
adv=''
testresult=u'��ͨ��'


excel = xlrd.open_workbook(importname)#��excel
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
	#list.append(attid)  �Ƿ���Ҫ��������keyΪlist
	tlist = []#��ʱlist������appid,attid,result,comment,tid,grade,adv
	maplist = []#��ʱlist�����������appid,attid,result,comment,tid,grade,adv��
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

file = xlwt.Workbook()#���������͵���
table = file.add_sheet('Sheet1')
table.write(0,0,nol1)
table.write(0,1,nol2)
table.write(0,2,nol3)
table.write(0,3,nol4)
table.write(0,4,nol5)
table.write(0,5,nol55)

file1 = xlwt.Workbook()#���Խ������
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
	testresult = u'��ͨ��'
	rdirect = '0'
	avg = 0
	tmp = ab[key]
	x = 1 #һ������ͨ����������
	for t in range(len(tmp)):
		if tmp[t][2]==u'���ܲ���ͨ��':			
			testresult = u'ͨ��'
			rdirect = ''
			#logger.debug(u'ͨ��')
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
	if testresult == u'��ͨ��':
		table1.write(row1,4,tmp[0][3])
	#table1.write(row1,4,'')		
	table1.write(row1,5,avg/x)#��ƽ����
	row1 = row1 + 1

		
file.save(wtname)
file1.save(wrname)
print row1-1