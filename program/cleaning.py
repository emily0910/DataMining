# -*- coding: UTF-8 -*-
__author__ = 'jennyzhang'

import xlrd
import json

#打开excel
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

#获取数据
def init():
    #打开excel表格
    file= 'data.xlsx'
    data = open_excel(file)
    #进入表单“Sheet1”
    table = data.sheet_by_name('Sheet1')
    #获得Sheet1的行数与列数
    rows = table.nrows  #行数
    cols = table.ncols  #列数
    #把所有列的内容分别保存
    #获取标称变量的值（前三列）
    nominalData={}
    nominalData["season"]=table.col_values(0) #季节
    nominalData["riverSize"]=table.col_values(1)  #河流大小
    nominalData["riverSpeed"]=table.col_values(2)  #河水速度
    #获取水样化学参数（第4至第11列）
    chemicalParameters={}
    chemicalParameters["mxPH"]=table.col_values(3)
    chemicalParameters["mnO2"]=table.col_values(4)
    chemicalParameters["Cl"]=table.col_values(5)
    chemicalParameters["NO3"]=table.col_values(6)
    chemicalParameters["NH4"]=table.col_values(7)
    chemicalParameters["oPO4"]=table.col_values(8)
    chemicalParameters["PO4"]=table.col_values(9)
    chemicalParameters["Chla"]=table.col_values(10)
    #获得7种不同有害藻类在相应水样中的频率数目（第12至第18列）
    frequency={}
    frequency["a1"]=table.col_values(11)
    frequency["a2"]=table.col_values(12)
    frequency["a3"]=table.col_values(13)
    frequency["a4"]=table.col_values(14)
    frequency["a5"]=table.col_values(15)
    frequency["a6"]=table.col_values(16)
    frequency["a7"]=table.col_values(17)

    return nominalData,chemicalParameters,frequency





#数据清洗 将缺失部分剔除（“XXXXXX”删除）
def cleaning(chemicalParameters,frequency):
    weed_chemicalParameters=chemicalParameters
    weed_frequency=frequency
    cleaning_weed(weed_chemicalParameters)
    cleaning_weed(weed_frequency)
    return chemicalParameters,frequency
	# 利用高频数值来插值
    # library(DMwR)
    # Adata1 <- centralImputation(mydata)
    # write.table(Adata1,'D:/CentralImputationData.txt',col.names = F,row.names = F, quote = F)

    # 利用变量之间的相关关系来插值
    # symnum(cor(mydata[,1:23],use='complete.obs'))
	# lm(formula=rectal.temperature~surgery, data=mydata)
	# Adata1 <- mydata
	# fillPO4 <- function(oP){
		# if(is.na(oP))
			# return(NA)
		# else return (38.122193 + 0.00913 * oP)
	# }
	# Adata1[is.na(Adata1$rectal.temperature),'rectal.temperature'] <- sapply(Adata1[is.na(Adata1$rectal.temperature),'surgery'],fillPO4)
	# write.table(Adata1,'D:/linearDefaultData.txt',col.names = F,row.names = F, quote = F)

	# 利用案例之间的关系来插值
	# library(VIM)
	# a <- kNN(mydata,k = 3)
	# write.table(a[,1:28],'D:/knnImputationData.txt',col.names = F,row.names = F, quote = F)

	# 展示插值前后的变化
	# library(car)
	# par(mfrow=c(1,2))
	# hist(mydata$rectal.temperature,prob=T,xlab='', main='Histogram of rectal.temperature before handle', ylim=0:1)
	# hist(Adata1$rectal.temperature,prob=T,xlab='', main='Histogram of rectal.temperature after handle', ylim=0:1)
	# par(mfrow=c(1,1))


#将缺失部分剔除（“XXXXXX”删除）
def cleaning_weed(obj):
    for key in obj.keys():
        tmp=[]
        for i in range(0, len(obj[key])):
            if(isinstance(obj[key][i],(float,int))==True):
                tmp.append(obj[key][i])
        obj[key]=tmp

if __name__=="__main__":
    #打开excel获取数据
    (nominalData,chemicalParameters,frequency)=init()
    #只获取数值属性，保存为json
    result={}
    (chemicalParameters,frequency)=cleaning(chemicalParameters,frequency)
    result["chemicalParameters"]=chemicalParameters
    result["frequency"]=frequency
    #保存结果
    fileIn=open(r"D:\data.json",'w')
    data_save=json.dumps(result ,indent=1)
    #print json.dumps(result ,indent=1)
    fileIn.write(data_save)
    #对标称属性，给出每个可能取值的频数，


