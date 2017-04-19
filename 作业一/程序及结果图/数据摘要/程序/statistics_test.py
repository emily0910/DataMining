# -*- coding: UTF-8 -*-
__author__ = 'ZhifengXu'

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
    
    data = xlrd.open_workbook(r'C:\Users\xzf0724\Python\test.xlsx')
    #进入表单“Sheet1”
	
    table = data.sheet_by_index(0)
    #获得Sheet1的行数与列数
	# print table.sheet_names() # [u'sheet1', u'sheet2']
    rows = table.nrows  #行数
    cols = table.ncols  #列数
    #把所有列的内容分别保存
    #获取17个标称变量的值
    nominalData={}
    nominalData["surg"]=table.col_values(0) #手术
    nominalData["age"]=table.col_values(1)  #年龄
    nominalData["toe"]=table.col_values(6)  #四肢温度
    nominalData["pp"]=table.col_values(7)  #外设脉冲
    nominalData["mm"]=table.col_values(8)  #粘膜
    nominalData["crt"]=table.col_values(9)  #毛细血管再填充时间
    nominalData["peris"]=table.col_values(11)  #蠕动
    nominalData["ad"]=table.col_values(12)  #肿胀
    nominalData["nt"]=table.col_values(13)  #鼻胃管
    nominalData["nr"]=table.col_values(14)  #鼻胃反流
    nominalData["nrPH"]=table.col_values(15)  #鼻胃回流PH
    nominalData["ref"]=table.col_values(16)  #直肠检查 - 粪便
    nominalData["abd"]=table.col_values(17)  #腹部
    nominalData["aa"]=table.col_values(20)  #腹腔镜外观
    nominalData["outcome"]=table.col_values(22)  #结果
    nominalData["sl"]=table.col_values(23)  #手术病变
    nominalData["cpd"]=table.col_values(27)  #是否是病例数据
	
    #11个数值型变量的值
    parameters={}
    parameters["HN"]=table.col_values(2)  #医院编号
    parameters["rt"]=table.col_values(3)  #直肠温度
    parameters["pulse"]=table.col_values(4)  #脉冲
    parameters["rr"]=table.col_values(5)  #呼吸频率
    parameters["crt"]=table.col_values(10)  #疼痛
    parameters["pcv"]=table.col_values(18)  #填充细胞体积
    parameters["tp"]=table.col_values(19)  #总蛋白质
    parameters["atp"]=table.col_values(21)  #腹部总蛋白质
    parameters["tol1"]=table.col_values(24)  #tol1
    parameters["tol2"]=table.col_values(25)  #tol2
    parameters["tol3"]=table.col_values(26)  #tol3
	
    #获得11种不同数值型数据的相应频率数目
    frequency={}
    # frequency["a1"]=table.col_values(2)
    # frequency["a2"]=table.col_values(3)
    # frequency["a3"]=table.col_values(4)
    # frequency["a4"]=table.col_values(5)
    # frequency["a5"]=table.col_values(10)
    # frequency["a6"]=table.col_values(18)
    # frequency["a7"]=table.col_values(19)
    # frequency["a8"]=table.col_values(21)
    # frequency["a9"]=table.col_values(24)
    # frequency["a10"]=table.col_values(25)
    # frequency["a11"]=table.col_values(26)

    return nominalData,parameters,frequency



#求出标称量的频数
def nominalDataFrequency(nominalData):
    #手术
    nominalDataSurg={}
    for item in nominalData["surg"]:
        if item not in nominalDataSurg.keys():
            nominalDataSurg[item]=1
        else:
            nominalDataSurg[item]+=1

    #年龄
    nominalDataAge={}
    for item in nominalData["age"]:
        if item not in nominalDataAge.keys():
            nominalDataAge[item]=1
        else:
            nominalDataAge[item]+=1

    #四肢温度
    nominalDataToe={}
    for item in nominalData["toe"]:
        if item not in nominalDataToe.keys():
            nominalDataToe[item]=1
        else:
            nominalDataToe[item]+=1
    #外设脉冲
    nominalDataPp={}
    for item in nominalData["pp"]:
        if item not in nominalDataPp.keys():
            nominalDataPp[item]=1
        else:
            nominalDataPp[item]+=1
	#粘膜
    nominalDataMm={}
    for item in nominalData["mm"]:
        if item not in nominalDataMm.keys():
            nominalDataMm[item]=1
        else:
            nominalDataMm[item]+=1
	#毛细血管再填充时间
    nominalDataCrt={}
    for item in nominalData["crt"]:
        if item not in nominalDataCrt.keys():
            nominalDataCrt[item]=1
        else:
            nominalDataCrt[item]+=1
	#蠕动
    nominalDataPeris={}
    for item in nominalData["peris"]:
        if item not in nominalDataPeris.keys():
            nominalDataPeris[item]=1
        else:
            nominalDataPeris[item]+=1
	#肿胀
    nominalDataAd={}
    for item in nominalData["ad"]:
        if item not in nominalDataAd.keys():
            nominalDataAd[item]=1
        else:
            nominalDataAd[item]+=1
	#鼻胃管
    nominalDataNt={}
    for item in nominalData["nt"]:
        if item not in nominalDataNt.keys():
            nominalDataNt[item]=1
        else:
            nominalDataNt[item]+=1
	#鼻胃反流
	nominalDataNr={}
    for item in nominalData["nr"]:
        if item not in nominalDataNr.keys():
            nominalDataNr[item]=1
        else:
            nominalDataNr[item]+=1
	#鼻胃回流PH
	nominalDataNrPH={}
    for item in nominalData["nrPH"]:
        if item not in nominalDataNrPH.keys():
            nominalDataNrPH[item]=1
        else:
            nominalDataNrPH[item]+=1
	#直肠检查 - 粪便
	nominalDataRef={}
    for item in nominalData["ref"]:
        if item not in nominalDataRef.keys():
            nominalDataRef[item]=1
        else:
            nominalDataRef[item]+=1
	#腹部
	nominalDataAbd={}
    for item in nominalData["abd"]:
        if item not in nominalDataAbd.keys():
            nominalDataAbd[item]=1
        else:
            nominalDataAbd[item]+=1
	#腹腔镜外观
	nominalDataAa={}
    for item in nominalData["aa"]:
        if item not in nominalDataAa.keys():
            nominalDataAa[item]=1
        else:
            nominalDataAa[item]+=1
    #结果
	nominalDataOutcome={}
    for item in nominalData["outcome"]:
        if item not in nominalDataOutcome.keys():
            nominalDataOutcome[item]=1
        else:
            nominalDataOutcome[item]+=1
	#手术病变
	nominalDataSl={}
    for item in nominalData["sl"]:
        if item not in nominalDataSl.keys():
            nominalDataSl[item]=1
        else:
            nominalDataSl[item]+=1
    #是否是病例数据
	nominalDataCpd={}
    for item in nominalData["cpd"]:
        if item not in nominalDataCpd.keys():
            nominalDataCpd[item]=1
        else:
            nominalDataCpd[item]+=1

    #标称变量的频数统计结果整体保存为json对象
    nominalDataFrequency={}
    nominalDataFrequency["surg"]=nominalDataSurg
    nominalDataFrequency["age"]=nominalDataAge
    nominalDataFrequency["toe"]=nominalDataToe
    nominalDataFrequency["pp"]=nominalDataPp
    nominalDataFrequency["mm"]=nominalDataMm
    nominalDataFrequency["crt"]=nominalDataCrt
    nominalDataFrequency["peris"]=nominalDataPeris
    nominalDataFrequency["ad"]=nominalDataAd
    nominalDataFrequency["nt"]=nominalDataNt
    nominalDataFrequency["nr"]=nominalDataNr
	
    nominalDataFrequency["nrPH"]=nominalDataNrPH
    nominalDataFrequency["ref"]=nominalDataRef
    nominalDataFrequency["abd"]=nominalDataAbd
    nominalDataFrequency["aa"]=nominalDataAa
    nominalDataFrequency["outcome"]=nominalDataOutcome
    nominalDataFrequency["s1"]=nominalDataSl
    nominalDataFrequency["cpd"]=nominalDataCpd
    #print json.dumps(nominalDataFrequency,indent=1)
    #保存结果
    fileIn=open(r"D:\nominalDataFrequency_test.json",'w')
    data_save=json.dumps(nominalDataFrequency,indent=1)
    fileIn.write(data_save)

#数值属性，给出最大、最小、均值、中位数、四分位数及缺失值的个数
def statistic(parameters,frequency):
    (parameters,frequency)=cleaning(parameters,frequency)
    result={}
    #数值型变量的属性
    for key in parameters:
        result[key]={}
        result[key]["max"]=max(parameters[key])
        result[key]["min"]=min(parameters[key])
        result[key]["mean"]=sum(parameters[key])/len(parameters[key])
        result[key]["midian"]=midian(parameters[key])
        result[key]["quartiles"]=quartiles(parameters[key])
        result[key]["miss_num"]=68-len(parameters[key])
    #数值型变量的频率
    # for key in frequency:
        # result[key]={}
        # result[key]["max"]=max(frequency[key])
        # result[key]["min"]=min(frequency[key])
        # result[key]["mean"]=sum(frequency[key])/len(frequency[key])
        # result[key]["midian"]=midian(frequency[key])
        # result[key]["quartiles"]=quartiles(frequency[key])
        # result[key]["miss_num"]=368-len(frequency[key])
    #print json.dumps(result ,indent=1)
    #保存结果
    fileIn=open(r"D:\statistic_max_min_etc_test.json",'w')
    data_save=json.dumps(result ,indent=1)
    fileIn.write(data_save)


#求中位数
def midian(arr):
    arr.sort()
    if(len(arr)%2==0):
        return (arr[len(arr)/2]+arr[len(arr)/2-1])/2.0
    else:
        return arr[len(arr)/2]

#求四分位数
def quartiles(arr):
    arr.sort()
    Q=[]
    Q1=arr[(len(arr)+1)/4-1]
    Q2=arr[((len(arr)+1)/4)*2-1]
    Q3=arr[((len(arr)+1)/4)*3-1]
    Q.append(Q1)
    Q.append(Q2)
    Q.append(Q3)
    return Q

#数据清洗 将缺失部分剔除（“XXXXXX”删除）
def cleaning(parameters,frequency):
    weed_parameters=parameters
    weed_frequency=frequency
    cleaning_weed(weed_parameters)
    cleaning_weed(weed_frequency)
    return parameters,frequency


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
    (nominalData,parameters,frequency)=init()
    #对标称属性，给出每个可能取值的频数，
    nominalDataFrequency(nominalData)
    #数值属性，给出最大、最小、均值、中位数、四分位数及缺失值的个数
    statistic(parameters,frequency)

