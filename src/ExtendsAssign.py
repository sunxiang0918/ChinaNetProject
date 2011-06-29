#-*- coding:UTF-8 -*-
#import logging,sys
import xlrd
import xlwt
####################################
#
#该工具根据重点机型进行拓展用
#
####################################
#系列划分
####################

####android系列
series1 = ['4702']#4487 扩展
series2 = ['4792']#4792 不扩展
series3 = ['4984']#4984 不扩展
series4 = ['4806']#4271 扩展
series5 = ['5008','16757','16368','16389','5170','16488']#4973 扩展
series6 = ['4874','16270']#4436 扩展
series7 = ['16829','16288','5018','17168']#4962 扩展
series8 = ['5150']#5282 扩展
series9 = ['16668']#16788 扩展
series10 = ['16669']#16669 不扩展
series11 = ['5132']#5132 不扩展
android = {'4487':series1,'4271':series4,'4973':series5,'4436':series6,'4962':series7,'5282':series8,'16788':series9}
####WM系列
ser1 = ['75']#79
ser2 = ['119']#'5113'
ser3 = ['4195','4177','284','146','360','4159']#4105
ser4 = ['111']#4147
ser5 = ['246']#246  不扩展
ser6 = ['4470']#4470 不扩展
ser7 = ['4475','4776','4784']#4752
ser8 = ['4362']#4780 
ser9 = ['4183']#4183 不扩展
wm = {'79':ser1,'5113':ser2,'4105':ser3,'4147':ser4,'4752':ser7,'4780':ser8}


excelname = '真机适配机型导入.xls'
excel = xlrd.open_workbook(excelname)#读入已测试通过的适配关系
sheet = excel.sheet_by_name('Sheet1')
ab = {} #字典保存excel，key为附件ID
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

r = 0#写excel行数

for key in ab:
	tmp = ab[key]#测试通过的机型ID
	bl = False # 标识是否在android找到拓展机型	
	for x in range(len(tmp)):#循环通过的机型id
		for a in android:#遍历android扩展系列
			if int(tmp[x]) == int(a):
				bl = True#已在android系列找到
				series = android.get(a)
				#print series
				for s in series:
					table.write(r,0,key)					
					table.write(r,1,s)
					r = r + 1					
				break
		if bl == False:
			for w in wm:#遍历WM
				if int(tmp[x]) == int(w):
					series = wm.get(w)
					for s in series:
						table.write(r,0,key)					
						table.write(r,1,s)
						r = r + 1					
					break			
file.save('ExtendsAssign.xls')