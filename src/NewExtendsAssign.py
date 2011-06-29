#-*- coding:UTF-8 -*- 
'''
Created on 2011-5-6

@author: TXB
'''
import copy
import xlwt
import xlrd
import sys

mobileExcelFilePath=u'/mobile.xls'   #机型扩展信息excel
realExcelFilePath=u'/真机适配机型导入'    #需要扩展的信息excel
DEFALUT_SUFFIX='.xls'       #默认的文件后缀名

#设置默认的编码格式
reload(sys)
sys.setdefaultencoding('utf-8') #@UndefinedVariable

'''
    从机型Excel里面读取成组的机型信息
    @param filename: excel文件名  unix下需要写绝对路径并且文件名不能为中文,windows下可以写相对路径
    @return: 成组的机型信息   {'android':{key:[id,id,id],key:[id,id,id]},'WM':{key:[id,id,id],key:[id,id,id]}}
'''
def loadAllMobiles(filename):
    file=xlrd.open_workbook(filename,formatting_info=True)
    sheet=file.sheet_by_index(0)
    allMobile={u'android':{},u'WM':{}}    #构造规则文件列表
    temp=[]
    lastColor=None      #上一行的单元格背景色
    for row in range(sheet.nrows):    #遍历excel所有的行
        #获取单元格的背景色
        color=getBgColor(file,sheet,row,4)
        #判断该行的背景色是否和上一行的一样,这样用于分组
        if lastColor==None or lastColor!=color:
            #如果上一行的单元格背景色不为空,并且和这一行的不一样.则新建temp用于存放机型
            if lastColor!=None:
                temp=[]
            #把该行的单元格背景色赋予lastColor,方便下一行对比
            lastColor=color
            
            value=sheet.row(row)[0].value.strip()   #获取类型变量
            #区别android 和WM的,并分别放入Map中
            if 'android' in value:
                allMobile['android'][value]=temp
            else:
                allMobile['WM'][value]=temp
        
        #获取机型信息
        mobile={}           #机型对象
        mobile['brand']=str(sheet.row(row)[2].value).strip().decode('utf-8')    #获取品牌
        mobile['model']=str(sheet.row(row)[3].value).strip().decode('utf-8')       #获取型号
        string=str(sheet.row(row)[4].value).strip()     #去除字符串两端的空格
        string=string.replace(u'\xa0', u'').replace(u'\xc2',u'')    #去除字符串中的特殊字符
        if string!='':      #如果不为空,添加到临时列表中
            string=str(float(string)*1.0)
            mobile['mobileId']=string
            temp.append(mobile)
    return allMobile

'''
    获取某一单元格的背景色
    @param book: exlce文件
    @param sheet: sheet
    @param row:    行
    @param col:    列
    @return: 某一单元格的背景色 (int)
'''
def getBgColor(book,sheet,row,col):
    xfx=sheet.cell_xf_index(row,col)    #获取某一单元格的xf
    xf = book.xf_list[xfx]      #从整个excel表中获取xf
    bgColor = xf.background.pattern_colour_index    #获取单元格样式,这里是背景色
    return bgColor

'''
    读取 真机适配机型导入.xls 对里面的真机进行处理
    @param filename: 文件名
    @return: 根据附件ID 分组了的机型信息   {附件ID:[{'MOID': ID, 'ATTID': ID, 'APPID': ID, 'OPERTYPE':ID},{'MOID': ID, 'ATTID': ID, 'APPID': ID, 'OPERTYPE':ID}]}
'''
def loadRealMobiles(filename):
    file=xlrd.open_workbook(filename)       #读取Excel
    sheet=file.sheet_by_index(0)
    ab = {} #字典保存excel，key为附件ID
    #循环所有的行
    for row in range(1,sheet.nrows):
        tmplist = []
        attchid = sheet.row(row)[1].value   #获取附件ID
        if ab.get(attchid) != None :
            tmplist = ab.get(attchid)
        #构造信息对象,然后放入结果中
        temp={}
        temp['APPID']=sheet.row(row)[0].value
        temp['ATTID']=sheet.row(row)[1].value
        temp['MOID']=sheet.row(row)[4].value
        temp['OPERTYPE']=sheet.row(row)[5].value    #操作类型
        tmplist.append(temp)
        ab[attchid] = tmplist
    return ab

'''
    对真机进行扩展
    @param allmobiles: 成组的机型信息
    @param realMobiles: 要处理的真机信息
    @return: 已经扩展了的机型信息 
'''
def handleMobiles(allMobiles,realMobiles):
    result=[]
    for key in realMobiles:     #遍历所有的需要扩展的机子
        tmpList = realMobiles[key]      #测试通过的机型ID
        for index in range(len(tmpList)):   #循环通过的机型id
            mobile=tmpList[index]       #获取一个机型
            kuozhanList=getFittedMobiles(mobile, allMobiles)        #获取扩展列表
            if kuozhanList !=None and len(kuozhanList)>0:   #如果找到了需要扩展的信息,就添加到结果中
                for i in range(len(kuozhanList)):
                    isExists=False
                    temp={}
                    temp['MOID']=kuozhanList[i]['mobileId']
                    temp['ATTID']=mobile['ATTID']
                    temp['APPID']=mobile['APPID']
                    temp['brand']=kuozhanList[i]['brand']
                    temp['model']=kuozhanList[i]['model']
                    temp['OPERTYPE']=mobile['OPERTYPE']
                    for j in range(len(result)):
                        res=result[j]
                        #避免重复,这里需要判断
                        if temp['MOID']==res['MOID'] and temp['ATTID']==res['ATTID'] and temp['APPID']==res['APPID']:
                            isExists=True
                            break
                    if not isExists:    #如果不为重复,添加到结果集中
                        result.append(temp)
    return result

'''
    获取该机型需要扩展的机型,如果没有 返回空
    @param mobile:需要扩展的机型
    @param allMobiles:所有的扩展机型信息
    @return: 如果找到了,就返回扩展机型ID,不包括自身  [ID,ID,ID]
'''  
def getFittedMobiles(mobile,allMobiles):
    for kind in allMobiles:
        aKindOfMobile=allMobiles[kind]  #获取某一类别的机器
        for key in aKindOfMobile:
            mobiles=aKindOfMobile[key]
            moid=float(mobile['MOID'])*1.0  #先转换为浮点数
            moid=str(moid)      #转换成字符串
            for mob in mobiles: #判断机器是否是在列表中,如果是,则把除了他的所有扩展机型都放入结果中
                if moid == mob['mobileId']:
                    if mobile['OPERTYPE']=='1':
                        temp=copy.copy(mobiles)
                        #temp.remove(mob)
                        return temp
                    else:
                        temp=[]
                        temp.append(mob)
                        return temp
    #如果没有适配的真机，因为需要把所有的都添加进去。所以这个随便生成一个
    temp=[];
    tempmobile={}
    moid=float(mobile['MOID'])*1.0  #先转换为浮点数
    moid=str(moid)
    tempmobile['mobileId']=moid
    tempmobile['brand']=''
    tempmobile['model']=''
    temp.append(tempmobile)
    return temp

'''
    把结果写回新的Excel中 (真机适配机型导入_new.xls)
    @param result: 真机适配扩展结果
    @return: None
'''
def writeKuoZhanExcel(result):
    successedFile = xlwt.Workbook()
    failedFile= xlwt.Workbook()
    
    successedTable = successedFile.add_sheet('Sheet1')
    successedTable.write(0,0,u'应用ID')
    successedTable.write(0,1,u'附件ID')
    successedTable.write(0,2,u'真机适配：1;手动理论适配：2')
    successedTable.write(0,3,u'真机适配机型（华为平台机型ID）')
    successedTable.write(0,4,u'适配品牌')
    successedTable.write(0,5,u'适配机型')
    successedTable.write(0,6,u'操作类型')
    
    failedTable = failedFile.add_sheet('Sheet1')
    failedTable.write(0,0,u'应用ID')
    failedTable.write(0,1,u'附件ID')
    failedTable.write(0,2,u'真机适配：1;手动理论适配：2')
    failedTable.write(0,3,u'真机适配机型（华为平台机型ID）')
    failedTable.write(0,4,u'适配品牌')
    failedTable.write(0,5,u'适配机型')
    failedTable.write(0,6,u'操作类型')
    
    successIndex=0
    failedIndex=0
    for index in range(len(result)):
        if result[index]['OPERTYPE']=='1':
            successedTable.write(successIndex+1,0,result[index]['APPID'])
            successedTable.write(successIndex+1,1,result[index]['ATTID'])
            successedTable.write(successIndex+1,2,'1')
            successedTable.write(successIndex+1,3,int(float(result[index]['MOID'])))
            successedTable.write(successIndex+1,4,result[index]['brand'])
            successedTable.write(successIndex+1,5,result[index]['model'])
            successedTable.write(successIndex+1,6,result[index]['OPERTYPE'])
            successIndex=successIndex+1
        else:
            failedTable.write(failedIndex+1,0,result[index]['APPID'])
            failedTable.write(failedIndex+1,1,result[index]['ATTID'])
            failedTable.write(failedIndex+1,2,'1')
            failedTable.write(failedIndex+1,3,int(float(result[index]['MOID'])))
            failedTable.write(failedIndex+1,4,result[index]['brand'])
            failedTable.write(failedIndex+1,5,result[index]['model'])
            failedTable.write(failedIndex+1,6,result[index]['OPERTYPE'])
            failedIndex=failedIndex+1
    successedFile.save(realExcelFilePath+'_successed'+DEFALUT_SUFFIX)
    failedFile.save(realExcelFilePath+'_failed'+DEFALUT_SUFFIX)
    print 'already to write file'

'''
    程序的入口
'''
if __name__ == "__main__":

    allMobiles=loadAllMobiles(mobileExcelFilePath)
    
    realMobiles=loadRealMobiles(realExcelFilePath+DEFALUT_SUFFIX)

    result=handleMobiles(allMobiles, realMobiles)

    writeKuoZhanExcel(result)    