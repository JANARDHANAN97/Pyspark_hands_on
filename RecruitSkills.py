import pandas as pd
import numpy as np
import openpyxl
import collections, re


def CurateSkills(skillStr):
    UniqueSkillSet = []
    for subSkillStr in skillStr:
        subSkillList = subSkillStr.split(',')
        noNASkillList = [ re.sub(r'\(.*?\)', '', eachSubSkill) for eachSubSkill in subSkillList if (eachSubSkill != 'N/A' and eachSubSkill != 'nan')]
        noNASpaceSkillList = [re.sub(r'\s+', '', eachSubSkill) for eachSubSkill in noNASkillList]
        UniqueSkillSet.extend([eachSkill for eachSkill in noNASpaceSkillList if eachSkill not in UniqueSkillSet])
    return UniqueSkillSet
    
def GenerateThisWeekOIR(weekFolder):
    dfAll = pd.read_csv("D:\\CompetencyI&ES\\WeeklyMeeting\\"+weekFolder+"\\New_Open_Indent_Report.CSV")#ER&D OIR-PQ 23-2-21.CSV")
    df = dfAll[(dfAll.PRACTICE_CODE_TEXT == "CDP.AI")]
    newDUPractice = ["DU" if eachDu in ["IES", "ERA1", "ERA2", "ERAP", "EREU"] else "LTDU" for eachDu in df["EXECUTION_HUB"]]
    df["DU_LTDU"] = newDUPractice
    #print ("Modified DataFrame:",df)


    newdf = pd.pivot_table(df,index=["CUSTOMER_NAME"], columns=["DU_LTDU","INDENT_STATUS"], values=["OPEN_POS"],aggfunc=[np.sum],fill_value=0)
    #print ("New DataFrame is :", newdf)

    SkillDict = collections.OrderedDict()
    RecruitSkillDict = collections.OrderedDict()
    nOpenPosDict = collections.OrderedDict()
    
    for ind in df.index: 
        customerName = df['CUSTOMER_NAME'][ind]
        nOpenPos  = int(df['OPEN_POS'][ind])
        print "nOpenPos:", nOpenPos
        skill = str(df['ADDITIONAL_SKILLS'][ind]) + ',' + str(df['ALT_MAND_SKILL'][ind])
        if (customerName not in SkillDict.keys()):
            SkillDict[customerName]=[]
            RecruitSkillDict[customerName]=[]
            nOpenPosDict[customerName]=0
        if (skill not in SkillDict[customerName]):
            SkillDict[customerName].append(skill)
            isRecruit = df['INDENT_STATUS'][ind]
            if ( isRecruit == "RECRUIT" ):
                RecruitSkillDict[customerName].append(skill)
        nOpenPosDict[customerName]+= nOpenPos
    print ("nOpenPosDict is ", nOpenPosDict)
    pivotCustomerNamesRowIndex=newdf.index.values
    OrderedSkillList = []
    OrderedRecruitSkillList = []
    OrderedOpenPosn = []
    for CustomerName in pivotCustomerNamesRowIndex:
        print ("CustomerName = ", CustomerName, " and skill = ", SkillDict[CustomerName], "and OpenPos=", nOpenPosDict[CustomerName])
        curatedSkills = CurateSkills(SkillDict[CustomerName])
        RecruitCuratedSkills = CurateSkills(RecruitSkillDict[CustomerName])
        
        OrderedSkillList.append(curatedSkills) #SkillDict[CustomerName])
        OrderedRecruitSkillList.append(RecruitCuratedSkills)
        OrderedOpenPosn.append(nOpenPosDict[CustomerName])
    print ("OrderedOpenPosn =", OrderedOpenPosn)
    newdf['TotalPositions'] = OrderedOpenPosn
    newdf['ConsolidatedSkill']=OrderedSkillList
    newdf['RecruitSkill']=OrderedRecruitSkillList

    print (newdf)
    newdf.to_excel("D:\\CompetencyI&ES\\WeeklyMeeting\\"+weekFolder+"\\OIR2Recruit_CustomerSkillAnalysis.xlsx", sheet_name="OIR_CustomerSkillAnalysis")  


#MoveToLastWeekSheet( )
GenerateThisWeekOIR("Mar2_2021")
#skillString = ['Hadoop Admin (L3),N/A', 'Hadoop Admin (L2),N/A', 'N/A,N/A', 'Hadoop (L1),N/A', 'IBM zTPF - Transaction Processing Facility (L2),C++ (L2)', 'BDD - Behavioral Driven Development Testing (L2),DevOps-Cucumber (L2)', 'Apache Kafka (L2),N/A', 'IBM zTPF - Transaction Processing Facility (L1),SOAP UI Testing (L3)', 'IBM zTPF - Transaction Processing Facility (L2),SOAP UI Testing (L2)', 'IBM zTPF - Transaction Processing Facility (L1),C++ (L1)', 'Google Go Programming (L2),N/A', 'Microfocus LoadRunner (L2),Apache Cassandra database (L1)']
#res = CurateSkills(skillString)
#print "Result is :", res

#CompareWeeklyOIR( )
#df1['new column that will contain the comparison results'] = np.where(condition,'value if true','value if false')
#df1['priceDiff?'] = np.where(df1['Price1'] == df2['Price2'], 0, df1['Price1'] - df2['Price2']) #create new column in df1 for price diff 
#print (df1)
#Outer_Join = pd.merge(df1, df2, how='outer', on=['Client_ID', 'Client_ID'])
#print(Outer_Join)
#horizontal_concat = pd.concat([df3, df4], axis=1) 
  
#display(vertical_concat, horizontal_concat)

def MoveToLastWeekSheet( ) :
    ss = openpyxl.load_workbook("output.xlsx")
    ss_sheet = ss.get_sheet_by_name('Sheet1')
    ss_sheet.title = 'LastWeek'
    ss.save("output.xlsx")
def CompareWeeklyOIR( ):
    #inputFile1 = "test1.xls"
    #inputFile2 = "test2.xls"
    
    inputFile1 = "try1.xlsx"
    inputFile2 = "try2.xlsx"

    dfLastWk = pd.read_excel(inputFile1)
   
    print ("dfLastWk = ", dfLastWk)
    dfCurrWk = pd.read_excel(inputFile2)
    print ("dfCurrWk = ", dfCurrWk)
    print dfLastWk.compare(dfCurrWk)
