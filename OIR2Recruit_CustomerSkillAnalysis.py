#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import numpy as np
import openpyxl
import collections, re

import os 
os.chdir(r"C:\Users\janar\Downloads\Report files-2")
def CurateSkills(skillStr):
    UniqueSkillSet = []
    for subSkillStr in skillStr:
        subSkillList = subSkillStr.split(',')
        noNASkillList = [ re.sub(r'\(.*?\)', '', eachSubSkill) for eachSubSkill in subSkillList if (eachSubSkill != 'N/A' and eachSubSkill != 'nan')]
        noNASpaceSkillList = [re.sub(r'\s+', '', eachSubSkill) for eachSubSkill in noNASkillList]
        UniqueSkillSet.extend([eachSkill for eachSkill in noNASpaceSkillList if eachSkill not in UniqueSkillSet])
    return UniqueSkillSet
    
def GenerateThisWeekOIR(weekFolder):
    dfAll = pd.read_csv("New_Open_Indent_Report.CSV")#ER&D OIR-PQ 23-2-21.CSV")
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
        #print ("nOpenPos:", nOpenPos)
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
    #print ("nOpenPosDict is ", nOpenPosDict)
    
    pivotCustomerNamesRowIndex=newdf.index.values
    OrderedSkillList = []
    OrderedRecruitSkillList = []
    OrderedOpenPosn = []
    
    for CustomerName in pivotCustomerNamesRowIndex:
        #print ("CustomerName = ", CustomerName, " and skill = ", SkillDict[CustomerName], "and OpenPos=", nOpenPosDict[CustomerName])
        curatedSkills = CurateSkills(SkillDict[CustomerName])
        RecruitCuratedSkills = CurateSkills(RecruitSkillDict[CustomerName])
        
        OrderedSkillList.append(curatedSkills) #SkillDict[CustomerName])
        OrderedRecruitSkillList.append(RecruitCuratedSkills)
        OrderedOpenPosn.append(nOpenPosDict[CustomerName])
    #print ("OrderedOpenPosn =", OrderedOpenPosn)
    newdf['TotalPositions'] = OrderedOpenPosn
    newdf['ConsolidatedSkill']=OrderedSkillList
    newdf['RecruitSkill']=OrderedRecruitSkillList
    
    
    onsite = {}
    offshore = {}
    for i in set(df["CUSTOMER_NAME"]):
        #print(i, len(df[(df["CUSTOMER_NAME"]==i) & (df["AREA"]=="ONSITE")]))
        onsite[i] = len(df[(df["CUSTOMER_NAME"]==i) & (df["AREA"]=="ONSITE")])
        #print(i, len(df[(df["CUSTOMER_NAME"]==i) & (df["AREA"]=="OFFSHORE")]))
        offshore[i] = len(df[(df["CUSTOMER_NAME"]==i) & (df["AREA"]=="OFFSHORE")])
    #print(onsite)
    #print(offshore)
    newdf["onsite"] = onsite.values()
    newdf["offshore"] = offshore.values()
    
    df2 = pd.read_csv("TA_Report.csv")
    clms = df2.columns
    g1 = df2.groupby("CUSTOMER_ACCOUNT")
    for i in clms[2:]:
        #print(i,dict(g1[i].agg(np.sum)))
        newdf[i] = dict(g1[i].agg(np.sum)).values()
    
    #print (newdf)
    newdf.to_excel("OIR2Recruit_CustomerSkillAnalysis.xlsx", sheet_name="OIR_CustomerSkillAnalysis")  


GenerateThisWeekOIR("Mar5_2021")







