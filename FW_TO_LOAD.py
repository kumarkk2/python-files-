import xlrd
import openpyxl
import os
import sys
import xml.etree.cElementTree as ET



wb = xlrd.open_workbook('RG_RF_1.xlsx')
wb.sheet_names()
sheet1 = wb.sheet_by_name('Sheet1')


base_path=os.path.dirname(os.path.realpath(__file__))

xml_file = "RG_CLM_PP_DUM.xml"

tree=ET.parse(xml_file)

root=tree.getroot()
for child1 in root.findall('./FOLDER/JOB'):
      
   jobName_xml = child1.attrib['JOBNAME']


   for i in range(55):
      jobName_xl = str(sheet1.cell_value(i,4))
      outCond_xl=  str(sheet1.cell_value(i,5))
      #catName= sheet1.cell_value(i,0)
      #outcond1= sheet2.cell_value(i,9)
      
         
      if "_PP" in jobName_xml:
      
         if(jobName_xl==jobName_xml):
            #print(jobName_xml)
            in_cond = ET.Element('INCOND', NAME=outCond_xl, ODATE="ODAT", AND_OR="A")
            child1.append(in_cond)
            print(outCond_xl)
            #print(in_cond)
         
            out_cond = ET.Element('OUTCOND', NAME=outCond_xl, ODATE="ODAT", SIGN="-")
            child1.append(out_cond)
            #print(out_cond)
   
tree.write('RG_RF_LOAD2.xml')


      
##      for incond in child.findall('./INCOND'):
##         print(incond.attrib['NAME'])



