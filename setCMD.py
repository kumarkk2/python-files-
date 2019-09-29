import xlrd,xlwt
import openpyxl
import os
import sys
import xml.etree.cElementTree as ET



wb = xlrd.open_workbook('Calendars .xlsx')
wb.sheet_names()
sheet1 = wb.sheet_by_name('Sheet1')


base_path=os.path.dirname(os.path.realpath(__file__))

xml_file = "Finally.xml"

tree=ET.parse(xml_file)

root=tree.getroot()
for child1 in root.findall('./FOLDER/JOB'):
      
   jobName_xml = child1.attrib['JOBNAME']

##   if jobName_xml in sheet1.cell_value
##
##   if '_DUMMY' in jobName_xml:
##        dummyInd = jobName_xml.index('_DUMMY')
##        commonPattern = jobName_xml[:dummyInd]
##        corespondPP = commonPattern+'_PP'
##        jobName_xml = corespondPP
##   else:
##        jobName_xml = child1.attrib['JOBNAME']
   

   for i in range(2098):
       jobName_xl = str(sheet1.cell_value(i,0))
       if jobName_xl in jobName_xml:

          DAYS_CAL=str(sheet1.cell_value(i,3))
          child1.attrib["DAYSCAL"] = DAYS_CAL
           #cmdline = str(sheet1.cell_value(i,5))
           #child1.attrib["CMDLINE"] = cmdline
           #print(child1.attrib['JOBNAME'])
           #print(cmdline)
##          print(jobName_xml)
##          print(cmdline)
      
##      elif '_PP' in jobName_xml:
##          dummyJob = jobName_xml.replace('_PP','_DUMMY')
          
        
print('Done')
tree.write('Calendar_added.xml')


      
##      for incond in child.findall('./INCOND'):
##         print(incond.attrib['NAME'])



