import xlrd
import openpyxl
import os
import sys
import xml.etree.cElementTree as ET

wb = xlrd.open_workbook('PARTY_ID_JOBS.xlsx')
wb.sheet_names()
sheet1 = wb.sheet_by_name('Sheet4')

base_path=os.path.dirname(os.path.realpath(__file__))
xml_file = "afterparty_linkless.xml"
tree=ET.parse(xml_file)
root=tree.getroot()

for child1 in root.findall('./FOLDER/JOB'):

    jobName_xml = child1.attrib['JOBNAME']

    for i in range(125):
      pre_jobName_xl = str(sheet1.cell_value(i,0))
      act_jobName_xl = str(sheet1.cell_value(i,2))
      post_jobName_xl= str(sheet1.cell_value(i,4))
      outCond_xl=  str(sheet1.cell_value(i,1))
      inCond_xl=  str(sheet1.cell_value(i,3))

      if(pre_jobName_xl==jobName_xml):
            #print(jobName_xml)
            out_cond_p = ET.Element('OUTCOND', NAME=outCond_xl, ODATE="ODAT", SIGN="+")
            child1.append(out_cond_p)

      if(post_jobName_xl==jobName_xml):
            #print(jobName_xml)
            in_cond = ET.Element('INCOND', NAME=inCond_xl, ODATE="ODAT", AND_OR="A")
            child1.append(in_cond)
            out_cond_n = ET.Element('OUTCOND', NAME=inCond_xl, ODATE="ODAT", SIGN="-")
            child1.append(out_cond_n)
            
tree.write('afterparyidlinking3.xml')
    

