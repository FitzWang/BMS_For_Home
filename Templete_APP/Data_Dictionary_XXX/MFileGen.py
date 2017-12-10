# -*- coding:utf-8 -*-  
import os
import xlrd

Filepath = os.getcwd()
Filename = Filepath.split("\\")[-1]
ExcelFile = xlrd.open_workbook(Filename + '.xlsx')
Excelsheet = ExcelFile.sheets()[0]
ExcelRows = Excelsheet.nrows
print(ExcelRows)

#Creat MFile
MFileName = Filename + '.m'
MFile = open(MFileName,'w')
for row in range(1,ExcelRows):
    RowValues = Excelsheet.row_values(row)
    MFile.write(RowValues[0]+' = fitz.'+RowValues[1]+';\n')
    MFile.write(RowValues[0]+'.CoderInfo.StorageClass = \'Custom\';\n')
    if RowValues[1] == 'Constant':
        MFile.write(RowValues[0]+'.CoderInfo.CustomStorageClass =\'Const\';\n')
        if RowValues[2] == 'float':
            InitialValue = RowValues[0]+'.Value = '+str(RowValues[3])+';\n'
        else:
            InitialValue = RowValues[0]+'.Value = '+str(int(RowValues[3]))+';\n'   
        Dimensions = ''
    elif RowValues[1] == 'Variable':
        MFile.write(RowValues[0]+'.CoderInfo.CustomStorageClass =\'Volatile\';\n')
        if RowValues[2] == 'float':
            InitialValue = RowValues[0]+'.InitialValue = \''+str(RowValues[3])+'\';\n'
        else:
            InitialValue = RowValues[0]+'.InitialValue = \''+str(int(RowValues[3]))+'\';\n'
        Dimensions = RowValues[0]+'.Dimensions = '+str(int(RowValues[6]))+';\n'
    MFile.write(RowValues[0]+'.CoderInfo.CustomAttributes.HeaderFile = \''+Filename+'.h\';\n')
    MFile.write(RowValues[0]+'.CoderInfo.CustomAttributes.DefinitionFile = \''+Filename+'.c\';\n')
    MFile.write(RowValues[0]+'.DataType = \''+RowValues[2]+'\';\n')
    MFile.write(RowValues[0]+'.Min = ['+RowValues[4]+'];\n')
    MFile.write(RowValues[0]+'.Max = ['+RowValues[5]+'];\n')
    MFile.write(Dimensions)
    MFile.write(InitialValue+'\n\n')
    