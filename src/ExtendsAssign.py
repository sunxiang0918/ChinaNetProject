#-*- coding:cp936 -*-
#import logging,sys
import xlrd
import xlwt
####################################
#
#�ù��߸����ص���ͽ�����չ��
#
####################################
#ϵ�л���
####################

####androidϵ��
series1 = ['4702']#4487 ��չ
series2 = ['4792']#4792 ����չ
series3 = ['4984']#4984 ����չ
series4 = ['4806']#4271 ��չ
series5 = ['5008','16757','16368','16389','5170','16488']#4973 ��չ
series6 = ['4874','16270']#4436 ��չ
series7 = ['16829','16288','5018','17168']#4962 ��չ
series8 = ['5150']#5282 ��չ
series9 = ['16668']#16788 ��չ
series10 = ['16669']#16669 ����չ
series11 = ['5132']#5132 ����չ
android = {'4487':series1,'4271':series4,'4973':series5,'4436':series6,'4962':series7,'5282':series8,'16788':series9}
####WMϵ��
ser1 = ['75']#79
ser2 = ['119']#'5113'
ser3 = ['4195','4177','284','146','360','4159']#4105
ser4 = ['111']#4147
ser5 = ['246']#246  ����չ
ser6 = ['4470']#4470 ����չ
ser7 = ['4475','4776','4784']#4752
ser8 = ['4362']#4780 
ser9 = ['4183']#4183 ����չ
wm = {'79':ser1,'5113':ser2,'4105':ser3,'4147':ser4,'4752':ser7,'4780':ser8}


excelname = '���������͵���.xls'
excel = xlrd.open_workbook(excelname)#�����Ѳ���ͨ���������ϵ
sheet = excel.sheet_by_name('Sheet1')
ab = {} #�ֵ䱣��excel��keyΪ����ID
for r in range(sheet.nrows-1):
	tmplist = []
	attchid = sheet.row(r+1)[1].value
	if ab.get(attchid) != None :
		tmplist = ab.get(attchid)
	tmplist.append(sheet.row(r+1)[4].value)
	ab[attchid] = tmplist
print ab
file = xlwt.Workbook()
table = file.add_sheet('Sheet1')

r = 0#дexcel����

for key in ab:
	tmp = ab[key]#����ͨ���Ļ���ID
	bl = False # ��ʶ�Ƿ���android�ҵ���չ����	
	for x in range(len(tmp)):#ѭ��ͨ���Ļ���id
		for a in android:#����android��չϵ��
			if int(tmp[x]) == int(a):
				bl = True#����androidϵ���ҵ�
				series = android.get(a)
				#print series
				for s in series:
					table.write(r,0,key)					
					table.write(r,1,s)
					r = r + 1					
				break
		if bl == False:
			for w in wm:#����WM
				if int(tmp[x]) == int(w):
					series = wm.get(w)
					for s in series:
						table.write(r,0,key)					
						table.write(r,1,s)
						r = r + 1					
					break			
file.save('ExtendsAssign.xls')