import xlrd
import openpyxl
import os
import sys
import xml.etree.cElementTree as ET

#country = 'RD_EVERYTHING'#change
file = open('KIRANKT.xml', 'a')
     #Creating a dummy job
    #======================
jobString1= '<JOB PARENT_FOLDER="RDWZZ_D1_MIGRATION" CYCLIC_TYPE="C" CYCLIC_TOLERANCE="0" VERSION_HOST="L507B9D7FE50F" VERSION_SERIAL="6" IS_CURRENT_VERSION="Y" VERSION_OPCODE="N" USE_INSTREAM_JCL="N" MULTY_AGENT="N" APPL_TYPE="OS" RULE_BASED_CALENDAR_RELATIONSHIP="O" CHANGE_TIME="160300" CHANGE_DATE="20190417" CHANGE_USERID="chattea3" CREATION_TIME="143732" CREATION_DATE="20190707" CREATION_USER="chattea3" IND_CYCLIC="S" SYSDB="0" SHIFTNUM="+00" SHIFT="Ignore Job" DAYS_AND_OR="A" DEC="1" NOV="1" OCT="1" SEP="1" AUG="1" JUL="1" JUN="1" MAY="1" APR="1" MAR="1" FEB="1" JAN="1" MAXRUNS="0" MAXDAYS="0" AUTOARCH="0" MAXRERUN="0" MAXWAIT="0" RETRO="0" DAYSCAL="CalenderName" CONFIRM="0" CMDLINE="command_var" MEMLIB="\\10.52.116.242\" INTERVAL="00001M" NODEID="APPAU101MEL5317.globaltest.anz.com" CYCLIC="0" TASKTYPE="Command" CRITICAL="0" RUN_AS="AUCTMTERADATADSA" CREATED_BY="chattea3" DESCRIPTION="description" JOBNAME="jobname_var" MEMNAME="memname_var" SUB_APPLICATION="subapp_load" APPLICATION="app_var" JOBISN="63">'
jobString2='<INCOND NAME="in_a" ODATE="ODAT" AND_OR="A"/>'
jobString3='<OUTCOND NAME="out_p" ODATE="ODAT" SIGN="+"/>'
#jobString6='<OUTCOND NAME="out_p1" ODATE="ODAT" SIGN="+"/>'
jobString4='<OUTCOND NAME="out_n" ODATE="ODAT" SIGN="-"/>'
jobString5='</JOB>'

base_path=os.path.dirname(os.path.realpath(__file__))

wb = xlrd.open_workbook('RD_EVERYTHING.xlsx')
wb.sheet_names()
sheet1 = wb.sheet_by_name('RD_EVERYTHING')#change
#sheet2=  wb.sheet_by_name(Sheet2)

for i in range(1,1283):#change

 #To append the job and command
    #==============================
    jobName=str(sheet1.cell_value(i,1))
    prejobName=str(sheet1.cell_value(i-1,1))
    sucjobName=str(sheet1.cell_value(i+1,1))

    
    i_Condition=prejobName+'-TO-'+jobName
    o_Condition=jobName+'-TO-'+sucjobName
    command=str(sheet1.cell_value(i,2))
    category=str(sheet1.cell_value(i,0))

    #inCond_toPP=str(sheet2.cell_value(i,0))

    

    addCMD = jobString1.replace("command_var", command)
            
    addJobname = addCMD.replace("jobname_var", jobName)
    addjob_mem_cmd_Name=addJobname.replace("memname_var", jobName)

    #To add the in condition & To remove the in condition
    #===================================================
    #Start_job
    if (sheet1.cell_value(i,0)!=sheet1.cell_value(i-1,0)) and (sheet1.cell_value(i,0)==sheet1.cell_value(i+1,0)):
        in_cond='RDWRG_'+str("INCONDITION_NOT_REQ")
        #inCond_toPP=str(sheet2.cell_value(j,0))
        addIncond=jobString2.replace("in_a",in_cond)
        addOutcond=jobString3.replace("out_p",o_Condition)
        #addOutcond_pp=jobString6.replace("out_p1",in_cond)
        remOutcond=jobString4.replace("out_n",in_cond)
        
        job_Def=addjob_mem_cmd_Name+'\n'+addIncond+'\n'+addOutcond+'\n'+remOutcond+'\n'+jobString5
        file.write(job_Def+'\n')
            
    #Middle_job
    elif (sheet1.cell_value(i,0)==sheet1.cell_value(i-1,0)) and (sheet1.cell_value(i,0)==sheet1.cell_value(i+1,0)):
        in_cond='RDWRG_'+str("INCONDITION_NOT_REQ")
        addIncond=jobString2.replace("in_a",i_Condition)
        addOutcond=jobString3.replace("out_p",o_Condition)
        #addOutcond_pp=jobString6.replace("out_p1",in_cond)
        remOutcond=jobString4.replace("out_n",i_Condition)

        job_Def=addjob_mem_cmd_Name+'\n'+addIncond+'\n'+addOutcond+'\n'+remOutcond+'\n'+jobString5
        file.write(job_Def+'\n')

        
    #End_job
    elif (sheet1.cell_value(i,0)==sheet1.cell_value(i-1,0)) and (sheet1.cell_value(i,0)!=sheet1.cell_value(i+1,0))and ("_DUMMY" not in str(sheet1.cell_value(i,1))):
        
        in_cond='RDWRG_'+str("INCONDITION_NOT_REQ")
        if("_BTQ" in jobName):
            dummyJob_name=jobName.replace("_BTQ","_DUMMY")
        elif("_FLD" in jobName):
            dummyJob_name=jobName.replace("_FLD","_DUMMY")


        outDummy_cond=jobName+"-TO-"+ dummyJob_name

        addIncond=jobString2.replace("in_a",i_Condition)
        remOutcond=jobString4.replace("out_n",i_Condition)
        addOutcond=jobString3.replace("out_p",outDummy_cond)

        
        job_Def=addjob_mem_cmd_Name+'\n'+addIncond+'\n'+addOutcond+'\n'+remOutcond+'\n'+jobString5
        file.write(job_Def+'\n')

        
        addIncond_dummy=jobString2.replace("in_a",outDummy_cond)
        remOutcond_dummy=jobString4.replace("out_n",outDummy_cond)
        addOutcond_dummy=jobString3.replace("out_p",in_cond)

            
        addJobname = addCMD.replace("jobname_var", dummyJob_name)
        addjob_mem_cmd_Name_dummy=addJobname.replace("memname_var", dummyJob_name)
        
        job_Def=addjob_mem_cmd_Name_dummy+'\n'+addIncond_dummy+'\n'+addOutcond_dummy+'\n'+remOutcond_dummy+'\n'+jobString5
        file.write(job_Def+'\n')

        

    elif (sheet1.cell_value(i,0)==sheet1.cell_value(i-1,0)) and (sheet1.cell_value(i,0)!=sheet1.cell_value(i+1,0)) and ("_DUMMY" in str(sheet1.cell_value(i,1))):

        in_cond='RDWRG_'+str("INCONDITION_NOT_REQ")
        dummyJob_name=jobName.replace("_BTQ","_DUMMY")
        outDummy_cond=jobName+"-TO-"+dummyJob_name

        addIncond=jobString2.replace("in_a",i_Condition)
        remOutcond=jobString4.replace("out_n",i_Condition)
        addOutcond=jobString3.replace("out_p",in_cond)

        job_Def=addjob_mem_cmd_Name+'\n'+addIncond+'\n'+addOutcond+'\n'+remOutcond+'\n'+jobString5
        file.write(job_Def+'\n')

    else:
        print(jobName)


file.close()
    
        

  



