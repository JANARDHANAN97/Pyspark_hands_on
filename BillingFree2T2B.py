import pandas as pd
import numpy as np
import openpyxl
import collections, re
# encoding=utf8  
import sys  

reload(sys)  
sys.setdefaultencoding('utf8')

def addInfo(dataFrame, colName, nIndex, valueList) :
    try:
        if ( nIndex != -1):
            valueList.append(dataFrame[colName].iloc[nIndex])
        else:
            valueList.append(dataFrame[colName].astype('str').str.cat(sep='\n'))
    except:
        valueList.append('')
 
def GenerateAEDReport(weekFolder):
    dfCDPEmp = pd.read_csv("D:\\CompetencyI&ES\\WeeklyMeeting\\CDPFresherTracker.csv")# , encoding="utf-8")
    dfERD = pd.read_csv("D:\\CompetencyI&ES\\WeeklyMeeting\\"+weekFolder+"\\ERDEBD.csv")#, encoding='utf8')
    dfCHM = pd.read_csv("D:\\CompetencyI&ES\\WeeklyMeeting\\"+weekFolder+"\\DailyCHM.csv")#, encoding="utf-8")
	#dfTrngTracker = pd.read_csv("D:\\CompetencyI&ES\\WeeklyMeeting\\"+weekFolder+"\\TrainingTracker.csv")#, encoding="utf-8")
	
    LOCATION_LIST = []
    EMPLOYEE_EMAIL_ID_LIST=[]
    HOME_ORGUNIT_DESC_LIST = []
    PRAC_CC_DESC_LIST = []
    PRIMARY_SUPERVISOR_LIST = []
    BILLABILITY_STATUS_LIST = []
    BILLABLE_CATEGORY_LIST = []
    EXPERIENCE_LIST = []
    PROJECT_ACQUIRED_SKILL_LIST = []
    CERTIFIED_SKILL_LIST = []
    EBD_LIST = []
    ASSIGN_END_LIST = []
    CUSTOMER_LIST = []
    PROJECT_SUPERVISOR_LIST = []
    for ind in dfCDPEmp.index:
      #try:
        EMP_NO = dfCDPEmp['EMP NO'][ind] #EMP ID
        print ("EMP_NO =", EMP_NO)
        
        dfCHMEmp = dfCHM[(dfCHM.EMP_CODE== EMP_NO)]
        dfERDEmp = dfERD[(dfERD['EMP_NO'] == EMP_NO)]

        addInfo(dfCHMEmp, 'LOCATION',           -1, LOCATION_LIST)
        addInfo(dfCHMEmp, 'EMPLOYEE_EMAIL_ID',  -1, EMPLOYEE_EMAIL_ID_LIST)
        addInfo(dfCHMEmp, 'HOME_ORGUNIT_DESC',  -1, HOME_ORGUNIT_DESC_LIST)
        addInfo(dfCHMEmp, 'PRAC_CC_DESC',       -1, PRAC_CC_DESC_LIST)
        addInfo(dfCHMEmp, 'PRIMARY_SUPERVISOR', -1, PRIMARY_SUPERVISOR_LIST)
        addInfo(dfCHMEmp, 'PROJECT_SUPERVISOR', -1, PROJECT_SUPERVISOR_LIST)
        addInfo(dfCHMEmp, 'BILLABILITY_STATUS', -1, BILLABILITY_STATUS_LIST)
        addInfo(dfCHMEmp, 'BILLABLE_CATEGORY',  -1, BILLABLE_CATEGORY_LIST)
        addInfo(dfCHMEmp, 'EXPERIENCE',  -1, EXPERIENCE_LIST)
        addInfo(dfCHMEmp, 'PROJECT_ACQUIRED_SKILL',  -1, PROJECT_ACQUIRED_SKILL_LIST)
        addInfo(dfCHMEmp, 'CERTIFIED_SKILL',  -1, CERTIFIED_SKILL_LIST)
        
        dfERDEmp['ALLOCATION_DATE'] = pd.to_datetime(dfERDEmp['ALLOCATION_DATE'])
        
        dfERDEmpRecentAllocation = dfERDEmp.sort_values('ALLOCATION_DATE').tail(1)#.groupby(['ASSIGN_END','EBD','CUSTOMER_NAME'])
        #dfERDEmpRecentAllocation = dfERDEmp.sort_values('ALLOCATION_DATE').groupby(['ASSIGN_END','EBD','CUSTOMER_NAME']).tail(1)
        print "dfERDEmpRecentAllocation is :", dfERDEmpRecentAllocation
        print "Allocation Date : dfERDEmpRecentAllocation is :", dfERDEmpRecentAllocation['ALLOCATION_DATE']
        addInfo(dfERDEmpRecentAllocation, 'ASSIGN_END', 0, ASSIGN_END_LIST)
        addInfo(dfERDEmpRecentAllocation, 'EBD', 0, EBD_LIST)
        addInfo(dfERDEmpRecentAllocation, 'CUSTOMER_NAME', 0, CUSTOMER_LIST)
      #except:
      # print ("Skip this record")
       
    dfCDPEmp['LOCATION'] =	LOCATION_LIST
    dfCDPEmp['EMPLOYEE_EMAIL_ID'] = EMPLOYEE_EMAIL_ID_LIST	
    dfCDPEmp['HOME_ORGUNIT_DESC'] =	HOME_ORGUNIT_DESC_LIST
    dfCDPEmp['PRAC_CC_DESC'] =	PRAC_CC_DESC_LIST
    dfCDPEmp['PRIMARY_SUPERVISOR'] = PRIMARY_SUPERVISOR_LIST
    dfCDPEmp['PROJECT_SUPERVISOR'] = PROJECT_SUPERVISOR_LIST
    dfCDPEmp['BILLABILITY_STATUS'] =	BILLABILITY_STATUS_LIST
    dfCDPEmp['BILLABLE_CATEGORY'] =	BILLABLE_CATEGORY_LIST
    dfCDPEmp['EXPERIENCE'] =	EXPERIENCE_LIST
    dfCDPEmp['PROJECT_ACQUIRED_SKILL'] =	PROJECT_ACQUIRED_SKILL_LIST
    dfCDPEmp['CERTIFIED_SKILL'] =	CERTIFIED_SKILL_LIST    
    
    dfCDPEmp['EBD'] = EBD_LIST
    dfCDPEmp['ASSIGN_END'] =	ASSIGN_END_LIST  
    dfCDPEmp['CUSTOMER'] = CUSTOMER_LIST
    
    dfCDPEmp['ASSIGN_END_FMT'] = pd.to_datetime(dfCDPEmp['ASSIGN_END'], infer_datetime_format=True)
    
    curr_time = pd.to_datetime("now")
    dfCDPEmp['NBILLABLEDAYS_LEFT'] = (dfCDPEmp['ASSIGN_END_FMT']-curr_time).dt.days
    
    
    #dfCDPEmp['Remarks'] = dfCDPEmp.apply(lambda row: row.BILLABILITY_STATUS != 'B', axis = 1) 
    print (dfCDPEmp.columns.values.tolist())
    print (type(dfCDPEmp['CDP_NONCDP']))
    print (type(dfCDPEmp['BILLABILITY_STATUS']))
    
    
    
    #print (dfCDPEmp['CDP_NONCDP'] == 'CDP' & dfCDPEmp['BILLABILITY_STATUS'] != 'B')
    
    conditions = [(dfCDPEmp['CDP_NONCDP'] == 'CDP') & (dfCDPEmp['BILLABILITY_STATUS'] == 'F'),
                  (dfCDPEmp['CDP_NONCDP'] == 'CDP') & (dfCDPEmp['PRAC_CC_DESC'] != "ER&D-CDP.ai"),
                  (dfCDPEmp['CDP_NONCDP'] == 'NOCDP') & (dfCDPEmp['FRESHER_TYPE'] != "ELITE") & (dfCDPEmp['BILLABILITY_STATUS'] == 'F' ),
                  (dfCDPEmp['CDP_NONCDP'] == 'CDP') & (dfCDPEmp['NBILLABLEDAYS_LEFT'] <30),
                  (dfCDPEmp['HOME_ORGUNIT_DESC'] == "IES-CDP.ai")]

    # create a list of the values we want to assign for each condition
    values = ['CDP Free Resource - Check WMG. Allocate in Academy', 
              'Check PRAC and Supervisor - Email Neeloofar', 
              'Check Availability - STAR_TURBO_WASE is FREE from NonCDP',
              'Billing End Date Nearing - Contact Commit holder',
              'Training Engagement - Explore Allocation in projects']

    # create a new column and use np.select to assign values to it using our lists as arguments
    dfCDPEmp['Remarks'] = np.select(conditions, values,default="No Remarks")
    

    print (dfCDPEmp)
    dfCDPEmp.to_csv("CDPFresherBillability"+weekFolder+".csv", index = False)#, sheet_name="AEDAnalysis", encoding="utf-8")



GenerateAEDReport("Feb18_2021")