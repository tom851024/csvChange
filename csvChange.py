import xml.etree.ElementTree as ET
from openpyxl import Workbook
import pandas as pd
from datetime import datetime

majorNumber = {'110': '中國文學系', '117': '中國文學系碩士在職專班', '120': '外國語文學系', '121': '外國語文學系碩士班英語教學組', '122': '外國語文學系碩士班英美文學組', '130': '歷史學系',
'140': '華語文教學國際碩士學位學程', '150': '日本語言文化學系', '180': '宗教研究所', '190': '哲學系', '197': '哲學系碩士在職專班', '210': '應用物理學系', '211': '應用物理學系材料及奈米科技組',
'212': '應用物理學系光電組', '220': '化學系', '221': '化學系化學組', '222': '化學系化學生物組', '230': '生命科學系', '231': '生命科學系生物醫學組', '232': '生命科學系生態暨生物多樣性組',
'240': '應用數學系', '250': '生醫暨材料科學國際博士學位學程', '260': '生物多樣性國際研究生博士學位學程', '310': '化學工程與材料工程學系', '330': '工業工程與經營資訊學系',
'337': '工業工程與經營資訊學系高階醫務工程與管理碩士在職專班', '340': '環境科學與工程學系', '350': '資訊工程學系', '351': '資訊工程學系資電工程組', '352': '資訊工程學系數位創意組',
'353': '資訊工程學系軟體工程組', '357': '資訊工程學系碩士在職專班', '358': '資訊工程學系碩士在職專班大數據物聯網應用組', '359': '資訊工程學系碩士在職專班高階資訊經營與創業組',
'360': '電機工程學系', '361': '電機工程學系 IC 設計與無線通訊組', '362': '電機工程學系奈米電子與能源技術組', '370': '數位創新碩士學位學程', '410': '企業管理學系',
'417': '企業管理學系高階企業經營碩士在職專班', '420': '國際經營與貿易學系', '430': '會計學系', '437': '會計學系碩士在職專班', '440': '財務金融學系', '447': '財務金融學系碩士在職專班',
'457': '高階經營管理碩士在職專班', '460': '國際企業管理碩士學位學程', '470': '統計學系', '490': '資訊管理學系', '520': '經濟學系', '521': '經濟學系一般經濟組', '522': '經濟學系產業經濟組',
'530': '政治學系', '531': '政治學系政治理論組', '532': '政治學系國際關係組', '533': '政治學系地方政治組', '540': '行政管理暨政策學系', '547': '行政管理暨政策學系第三部門碩士在職專班',
'550': '社會學系', '560': '社會工作學系', '570': '教育研究所', '577': '教育研究所碩士在職專班', '587': '公共事務碩士在職專班', '610': '畜產與生物科技學系', '620': '食品科學系',
'621': '食品科學系食品科技組', '622': '食品科學系食品工業管理組', '660': '餐旅管理學系', '667': '餐旅管理學系碩士在職專班', '670': '運動休閒與健康管理學位學程',
'680': '高齡健康與運動科學學士學位學程', '710': '美術學系', '717': '美術學系碩士在職專班', '720': '音樂學系', '730': '建築學系', '740': '工業設計學系', '747': '工業設計學系碩士在職專班',
'750': '景觀學系', '757': '景觀學系碩士在職專班', '760': '表演藝術與創作碩士學位學程', '810': '法律學系', '910': '國際經營管理學位學程', '920': '永續科學與管理學士學位學程',
'930': '國際學院不分系英語學士'} #科系索引

insertData = []

def date_key(key_value): #排序日期
    return datetime.strptime(key_value[1], "%Y/%m/%d")

def sortScore(dictData, topic_id, number_score, finish_date, stdId):
     sameData = [] # 紀錄原本已存在 dictData 中 同學號以及同測驗項目的 list key 值
     tmpDict = {} # 暫存此次測驗日期(val)及時間順序(key)
     needSort = False #判斷同樣學號的資料是否有同樣的測驗項目需排序
     for n in range(0, len(dictData[stdId]), 1):
        if dictData[stdId][n][0] == topic_id:
            sameData.append(n)
            tmpDict[dictData[stdId][n][3]] = dictData[stdId][n][2]
            needSort = True

     if needSort == True:
        tmpDict['0'] = finish_date
        sorted_dates = sorted(tmpDict.items(), key=date_key)
        sorted_dict = {key: value for key, value in sorted_dates}
        key_value_pairs = list(sorted_dict.items())
        index_of_0 = [pair[0] for pair in key_value_pairs].index('0')
        dictData[stdId].append([topic_id, number_score, finish_date, (index_of_0+1)]) #排序後並將這次迴圈跑到的資料加入 dictData
        
        for index in range(0, len(sameData), 1):
            # print(dictData[stdId][sameData[index]][3])
            index_of_n = [pair[0] for pair in key_value_pairs].index(str(dictData[stdId][sameData[index]][3]))
            dictData[stdId][index][3] = index_of_n + 1
            # print(index, dictData[stdId][index])
            # print(dictData)

        # print(dictData)

     return dictData

def getInsertData(stdId, stdData, dictData):
    tmpData = [None] * 40
    #取學制
    if stdId[0] == 'S':
        tmpData[0] = '學士'
    elif stdId[0] == 'G':
        tmpData[0] = '碩士'

    #取系所
    if stdId[3:6] in majorNumber:
        tmpData[1] = majorNumber[stdId[3:6]]
    else:
        tmpData[1] = ''

    tmpData[2] = stdId #學號
    tmpData[3] = 1 #診斷次數
    
    for n in dictData[stdId]: # 將分數依據類型放到暫存陣列對應位置中
        if n[0] == '11':
            tmpData[4] = n[1]
        elif n[0] == '12':
            tmpData[5] = n[1]
        elif n[0] == '13':
            tmpData[6] = n[1]
        elif n[0] == '14':
            tmpData[7] = n[1]
        elif n[0] == '15':
            tmpData[8] = n[1]
        elif n[0] == '16':
            tmpData[9] = n[1]
        elif n[0] == '17':
            tmpData[10] = n[1]
        elif n[0] == '18':
            tmpData[11] = n[1]

    tmpData[12] = dictData[stdId][0][2]
    # if stdId in dictData:
    #     tmpData = sortScore(tmpData, dictData)
    
    insertData.append(tmpData)

    return 0


file_object = open("8.xml") #讀 xml
ori_xml = file_object.read()
file_object.close()
pro_xml = ori_xml.replace("utf-8", "gb2313")
# print(pro_xml)

root = ET.fromstring(pro_xml)
dictData = {}

#測試資料
dictData["S07190055"] = []
dictData["S07190055"].append(["11", "5.87", "2023/11/05", "1"])
dictData["S07190055"].append(["11", "5.87", "2024/10/15", "2"])

for main_data in root.findall('.//commOcuppationMainData'):
    student_id = main_data.get('StudentID') #取學號
    if student_id in dictData: #判斷是否有二次以上的診斷
        needSort = True
        # print("True")
    else:
        needSort = False
        dictData[student_id] = []
        # print("False")
    
    # print(student_id)
    
    for detail_data in main_data.findall('.//commOcuppationDetailData'):
        topic_id = detail_data.get('Topic_ID') #取診斷項目 ID
        number_score = detail_data.get('Number_Score') #取診斷分數
        finish_date = detail_data.get('Finished_Date') #取診斷結束日期
        if needSort == True:
            dictData = sortScore(dictData, topic_id, number_score, finish_date, student_id)
            # print("--")
        else:
            dictData[student_id].append([topic_id, number_score, finish_date, 1])
        # print(dictData[student_id])

    #test
    # print(student_id)
    # if student_id == 'S07190055':
    #     print(dictData['S07190055'][1][1])


columnTitle = ["學制", "系所", "學號", "診斷次數", "溝通表達第1次", "持續學習第1次", "人際互動第1次", "團隊合作第1次", "問題解決第1次", "創新第1次", "工作責任及紀律第1次", "資訊科技應用第1次", "第1次診斷完成時間", 
"溝通表達第2次", "持續學習第2次", "人際互動第2次", "團隊合作第2次", "問題解決第2次", "創新第2次", "工作責任及紀律第2次", "資訊科技應用第2次", "第2次診斷完成時間",
"溝通表達第3次", "持續學習第3次", "人際互動第3次", "團隊合作第3次", "問題解決第3次", "創新第3次", "工作責任及紀律第3次", "資訊科技應用第3次", "第3次診斷完成時間",
"溝通表達第4次", "持續學習第4次", "人際互動第4次", "團隊合作第4次", "問題解決第4次", "創新第4次", "工作責任及紀律第4次", "資訊科技應用第4次", "第4次診斷完成時間"]

insertData.append(columnTitle)

for std_id, std_data in dictData.items():
    getInsertData(std_id, std_data, dictData)

df = pd.DataFrame(insertData)
# df.to_excel("score.xlsx", index=False, header=False)

