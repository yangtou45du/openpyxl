#!/usr/bin/python
# -*- coding: cp936 -*-
# -*- coding: encoding -*-
from openpyxl import load_workbook
import openpyxl
#��ȡ�ļ�
wb=load_workbook("testData1.xlsx")#Ĭ�ϴ��ļ����ɶ�д

#��ȡ������
sheetNames=wb.sheetnames#h��ȡ����sheet����
#print(sheetNames)
a_sheet=wb['Sheet1']#����sheet����ȡsheet
#print a_sheet
title=a_sheet.title#��ȡtitle��
#print(title)
sheet=wb.active#��ȡ��ǰ������ʾ��sheet
#print sheet


#��ȡ��Ԫ��ֵ
b4=sheet['B4']
print b4.column#��ȡ��
print(b4.row)#��ȡ��
print b4.value#��ȡ��Ԫ��ֵ
b4_too=sheet.cell(row=4,column=2)
print(b4_too.value)#��ȡ��Ԫ��ֻ
print(sheet.max_row)#��ȡ�����
print(sheet.max_column)#��ȡ�����
#sheet.rowsÿһ�е�����,ÿһ����Ԫ�棬sheet.columnsÿһ�У�ÿ��Ԫ����ÿһ�еĵ�Ԫ��
for row in sheet.rows:#��Ϊ���У����Է���A1, B1, C1������˳��
     for cell in row:
        print cell.value
for column in sheet.columns:# A1, A2, A3������˳��
    for cell in column:
        print cell.value
#��ȡĳ�е�����
for cell in list(sheet.rows)[2]:
    print cell.value
#��ȡ��������ĵ�Ԫ��,�������A1Ϊ���Ͻǣ�B3Ϊ���½Ǿ�����������е�Ԫ��
for i in range(1,4):
    for j in range(1,3):
        print(sheet.cell(row=i,column=j).value)
#��Ƭ
for row_cell in sheet['A1':'B3']:
   for cell in row_cell:
       print cell.value
for cell in sheet['A1':'B3']:
    print(cell)

#������ĸ����кţ������кŷ�����ĸ
from openpyxl.utils import get_column_letter, column_index_from_string
#�����е����ַ�����ĸ
print get_column_letter((2))#B
print(column_index_from_string('D'))#4

#������д��excel
from openpyxl import Workbook
wb=Workbook()#�½���������û����
sheet=wb.active#�������
sheet.title="test"#ֱ�Ӹ�ֵ�Ϳ��ԸĹ����������
print(wb.sheetnames)#��ȡ����
wb.create_sheet('data',index=1)# ������s���ڶ���������index=0���ǵ�һ��λ��
print(wb.sheetnames)#��ȡ����
#wb.remove(sheet)# ɾ��ĳ��������
#del wb[sheet]# ɾ��ĳ��������


