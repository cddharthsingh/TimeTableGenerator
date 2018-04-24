# -*- coding: utf-8 -*-
"""
Created on Thu Mar 29 02:04:38 2018

@author: Siddharth Singh
"""
import random
import xlrd
import xlsxwriter

workbook = xlrd.open_workbook('ttinput.xlsx')
worksheet = workbook.sheet_by_index(0)

#Reading Number of Sections
NumberOfSections=int(worksheet.cell(2, 2).value)
#print (NumberOfSections)

#Reading Number of Subjects
NumberOfSubjects=int(worksheet.cell(2, 7).value)
#print (NumberOfSubjects)

#Reading Number of Labs
NumberOfLabs=int(worksheet.cell(2, 12).value)
#print (NumberOfLabs)

#Reading Number of Batches per section
NumberOfBatches=int(worksheet.cell(4, 12).value)
#print (NumberOfBatches)

#Creating a dictionary that will hold Subject code as key and list of that subject's teachers as value
TeaDict = {}
for i in range (6,6+NumberOfSubjects):
    j=4
    temp = []
    while (worksheet.cell(i, j).value!=0):
        temp.append(worksheet.cell(i, j).value)
        j+=1
    TeaDict[worksheet.cell(i, 2).value] = temp
#print('Teachers For Each Subject:\n',TeaDict)

#Creating a dictionary that will store Subject code as key and subject credit as value
SubCredit = {worksheet.cell(i, 2).value:int(worksheet.cell(i, 3).value) for i in range (6,6+NumberOfSubjects)}
#print(SubCredit)

#Creating a list of subject Codes
SubList=[]
for i in range (0,NumberOfSubjects): SubList.append(worksheet.cell(6+i, 2).value)
#print(SubList)

#Creating a dictionary that will store Subject code as key and teacher as value for one section
ThisSection = {worksheet.cell(i, 2).value:'NULL' for i in range (6,6+NumberOfSubjects)}
#print(ThisSection)

#creating timetable table and initialising value as zeros except halfday
print('\n-----Initial-TIME-TABLE-----')
day = 6
hour = 7
table = [[0] * hour for i in range(day)]
table[5][4]=table[5][5]=table[5][6]='X' #halfday
for i in range (0,6): print(table[i])
print('-----------------------------')

def updateTTSecB(sheetno):
    #for section B doing same thing again but with ability to check if lab-timing is clashing or not
    
    worksheetSectionA = workbook.sheet_by_index(sheetno)
    #updating TimeTable
        
    #First setting lab timings. Each lab is of 3 hrs. each section has three lab day
    #choosing random days from week to set lab timing. number of lab day is equal to NumberOfBatches
    labDayFixed=[]
    labHours=[1,4]
    for i in range (0,NumberOfBatches):
        labSet=0
        while(labSet==0):
            labDay=random.randint(0,5)
            if labDay not in labDayFixed:
                j=random.choice(labHours)
                if labDay==5:
                        j=1             #on saturday lab can only be in first half
                if (worksheetSectionA.cell(labDay, labHours).value!='LAB'):
                    labDayFixed.append(labDay)
                    table[labDay][j]=table[labDay][j+1]=table[labDay][j+2]='LAB'
                    labSet=1
        
    
    #Using two loops to traverse through the time table
    for i in range (0,6):
        Sub = 'Blank'
        for j in range (0,7):
            if table[i][j]==0:
                
                check=0
                Secondpass=0
                Firstpass=1
                while(Secondpass==0):
                    #choosing a random subject from SubList that is not same as previous hour
                    if(j==0):
                        a=random.randint(0,NumberOfSubjects-1)
                        Sub = SubList[a]
                    else:
                        while(table[i][j-1]==Sub):
                            a=random.randint(0,NumberOfSubjects-1)
                            Sub = SubList[a]
                
                    #checking if the sub class is already taken in previous hours
                    for k in range (0,j):
                        if(Sub==table[i][k]):
                            Firstpass=0
                                            
                    #if passes first checkpost: check if subject's credit are remaining        
                    if (Firstpass==1):
                        if (SubCredit[Sub]!=0):
                            Secondpass=1
                            SubCredit.update({Sub:SubCredit[Sub]-1})
                        else:
                            Secondpass=0
                            
                    #Checking if number of trials are not more than number of subjects
                    check+=1
                    if(check==NumberOfSubjects):
                        Sub='Blank'
                        Secondpass=1
                
                    #if passes second check post
                    if (Secondpass==1):
                        table[i][j]=Sub
                        #selecting teacher for a section
                        if(Sub!='Blank'):
                            if(ThisSection[Sub]=='NULL'):
                                Teacher=TeaDict[Sub][random.randint(0,2)]
                                ThisSection[Sub]=Teacher
                                print(Sub, ':', ThisSection[Sub])




def updateTT():
    #updating TimeTable
    
    #First setting lab timings. Each lab is of 3 hrs. each section has three lab day
    #choosing random days from week to set lab timing. number of lab day is equal to NumberOfBatches
    labDayFixed=[]
    labHours=[1,4]
    for i in range (0,NumberOfBatches):
        labSet=0
        while(labSet==0):
            labDay=random.randint(0,5)
            if labDay not in labDayFixed:
                labDayFixed.append(labDay)
                j=random.choice(labHours)
                if labDay==5:
                    j=1             #on saturday lab can only be in first half
                table[labDay][j]=table[labDay][j+1]=table[labDay][j+2]='LAB'
                labSet=1
        
    
    #Using two loops to traverse through the time table
    for i in range (0,6):
        Sub = 'Blank'
        for j in range (0,7):
            if table[i][j]==0:
                
                check=0
                Secondpass=0
                Firstpass=1
                while(Secondpass==0):
                    #choosing a random subject from SubList that is not same as previous hour
                    if(j==0):
                        a=random.randint(0,NumberOfSubjects-1)
                        Sub = SubList[a]
                    else:
                        while(table[i][j-1]==Sub):
                            a=random.randint(0,NumberOfSubjects-1)
                            Sub = SubList[a]
                
                    #checking if the sub class is already taken in previous hours
                    for k in range (0,j):
                        if(Sub==table[i][k]):
                            Firstpass=0
                                            
                    #if passes first checkpost: check if subject's credit are remaining        
                    if (Firstpass==1):
                        if (SubCredit[Sub]!=0):
                            Secondpass=1
                            SubCredit.update({Sub:SubCredit[Sub]-1})
                        else:
                            Secondpass=0
                            
                    #Checking if number of trials are not more than number of subjects
                    check+=1
                    if(check==NumberOfSubjects):
                        Sub='Blank'
                        Secondpass=1
                
                    #if passes second check post
                    if (Secondpass==1):
                        table[i][j]=Sub
                        #selecting teacher for a section
                        if(Sub!='Blank'):
                            if(ThisSection[Sub]=='NULL'):
                                Teacher=TeaDict[Sub][random.randint(0,2)]
                                ThisSection[Sub]=Teacher
                                print(Sub, ':', ThisSection[Sub])
                            
def printTable():
    print('\n--------------------------TIME-TABLE--------------------------')
    for i in range (0,6): print(table[i])
    print('--------------------------------------------------------------\n')
    print('Remaining Subject Credits:\n',SubCredit)

updateTTSecB(0)
#printTable()
#if Subject credits remains, find a blank spot such that subject is not in that day and replace it with that subject
for i in range (0,6):
    day=0
    while (SubCredit[SubList[i]]!=0 and day<6):
        if SubList[i] not in table[day]:
            hr=0
            done=0
            while (done==0 and hr<7):
                if table[day][hr]=='Blank':
                    table[day][hr]=SubList[i]
                    SubCredit.update({SubList[i]:SubCredit[SubList[i]]-1})
                    done=1
                hr+=1
        else: day+=1        
   
printTable()


#creating a workbook named workbookOutput and creating worksheet1 named SectionA
workbookOutput = xlsxwriter.Workbook('Timetable.xlsx')
worksheet1 = workbookOutput.add_worksheet('SectionA')
cell_format = workbookOutput.add_format({'bold': True, 'font_color': '#ff1814', 'bg_color':'#f4f6f9'})
cell_format2 = workbookOutput.add_format({'valign':'center'})
#creating table format
col = 1
for u in range(0,7):
    hour= 'Hour ' + str(u+1)
    worksheet1.write(0, col, hour, cell_format)
    col += 1
worksheet1.write(1, 0, 'MONDAY', cell_format)
worksheet1.write(2, 0, 'TUESDAY', cell_format)
worksheet1.write(3, 0, 'WEDNESDAY', cell_format)
worksheet1.write(4, 0, 'THURSDAY', cell_format)
worksheet1.write(5, 0, 'FRIDAY', cell_format)
worksheet1.write(6, 0, 'SATURDAY', cell_format)
#writing content in table
row = 1
col = 1
for day in range(0,6):
    for hour in range(0,7):
        worksheet1.write(row, col, table[day][hour], cell_format2)
        col+=1
    row += 1
    col=1
    
updateTTSecB(0)
#if Subject credits remains, find a blank spot such that subject is not in that day and replace it with that subject
for i in range (0,6):
    day=0
    while (SubCredit[SubList[i]]!=0 and day<6):
        if SubList[i] not in table[day]:
            hr=0
            done=0
            while (done==0 and hr<7):
                if table[day][hr]=='Blank':
                    table[day][hr]=SubList[i]
                    SubCredit.update({SubList[i]:SubCredit[SubList[i]]-1})
                    done=1
                hr+=1
        else: day+=1        
   
printTable()

worksheet2 = workbookOutput.add_worksheet('SectionB')
cell_format = workbookOutput.add_format({'bold': True, 'font_color': '#ff1814', 'bg_color':'#f4f6f9'})
cell_format2 = workbookOutput.add_format({'valign':'center'})
#creating table format
col = 1
for u in range(0,7):
    hour= 'Hour ' + str(u+1)
    worksheet2.write(0, col, hour, cell_format)
    col += 1
worksheet2.write(1, 0, 'MONDAY', cell_format)
worksheet2.write(2, 0, 'TUESDAY', cell_format)
worksheet2.write(3, 0, 'WEDNESDAY', cell_format)
worksheet2.write(4, 0, 'THURSDAY', cell_format)
worksheet2.write(5, 0, 'FRIDAY', cell_format)
worksheet2.write(6, 0, 'SATURDAY', cell_format)
#writing content in table
row = 1
col = 1
for day in range(0,6):
    for hour in range(0,7):
        worksheet2.write(row, col, table[day][hour], cell_format2)
        col+=1
    row += 1
    col=1




workbookOutput.close()
