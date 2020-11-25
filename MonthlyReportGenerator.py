# -*- coding: utf-8 -*- 
import os
import sys
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(BASE_DIR)
from configparser import ConfigParser
import shutil
import easyword
import easyexcel
import datetime
from dateutil.relativedelta import relativedelta
import calendar
import pandas


def getpreviousmonthfullname(span):
    previousmonth = datetime.date.today() - relativedelta(months=+span)
    return calendar.month_name[previousmonth.month]

def getpreviousyear(span):
    previousmonth = datetime.date.today() - relativedelta(months=+span)        
    return str(previousmonth.year)

def getpreviousday(span):
    previousmonth = datetime.date.today() - relativedelta(months=+span)        
    return str(previousmonth.day).rjust(2, '0')

def getlastMonthlastDay():
    currentmonth = datetime.date.today()
    lastMonthlastDay = datetime.datetime(currentmonth.year, currentmonth.month, 1) - relativedelta(days=+1)        
    return str(lastMonthlastDay.day).rjust(2, '0')+'-'+getpreviousMonthSub(1)+'-'+getpreviousyear(1)

def getlastMonthlastDayInExcel():
    currentmonth = datetime.date.today()
    lastMonthlastDay = datetime.datetime(currentmonth.year, currentmonth.month, 1) - relativedelta(days=+1)        
    return str(lastMonthlastDay.day)+'/'+str(lastMonthlastDay.month)+'/'+str(lastMonthlastDay.year)

def getPreviousMonthFullInExcel(span):
    previousmonth = datetime.date.today() - relativedelta(months=+span)
    return str(previousmonth.year)+'-'+str(previousmonth.month).rjust(2, '0')+'-01'

def getpreviousMonthSub(span):
    months = "JanFebMarAprMayJunJulAugSepOctNovDec"
    previousmonth = datetime.date.today() - relativedelta(months=+span)
    pos = (previousmonth.month - 1) * 3
    return months[pos:pos+3]

def generateReport():
    
    print('Prepare the template start')
    work_dir = os.path.dirname(os.path.abspath(__file__))
    filepath = os.path.join(work_dir,'config.ini')
    print(filepath)
    config_raw = ConfigParser()
    config_raw.read(filepath)
    reportname = config_raw.get('general', 'reportname').strip()
    reporttemplate = config_raw.get('general', 'reporttemplate').strip()
    reporttemplatepath = os.path.join(work_dir,reporttemplate)
        
    reportpath = os.path.join(work_dir,reportname)
        
    if not os.path.exists(reporttemplatepath):
        print('The template file is not existed! Please check and rerun the program')
        return ''
        
    shutil.copy(reporttemplatepath,reportpath)
        
    print('Prepare the template end')
    
    print('Check the resource files start')
    resourceFile1 = config_raw.get('general', 'resourceFile1').strip()
    resourceFile2 = config_raw.get('general', 'resourceFile2').strip()
    resourceFile3 = config_raw.get('general', 'resourceFile3').strip()
    resourcePath1 = os.path.join(work_dir,resourceFile1)
    resourcePath2 = os.path.join(work_dir,resourceFile2)
    resourcePath3 = os.path.join(work_dir,resourceFile3)
    
    if not os.path.exists(resourcePath1):
        print('The resource file '+ resourceFile1 +' is not existed! Please check and rerun the program')
        return ''
    
    if not os.path.exists(resourcePath2):
        print('The resource file '+ resourceFile2 +' is not existed! Please check and rerun the program')
        return ''
    
    if not os.path.exists(resourcePath3):
        print('The resource file '+ resourceFile3 +' is not existed! Please check and rerun the program')
        return ''
    
    print('Check the resource files end')
    
    print('Update the monthly word report start')
    
    try:
      
        word = easyword.EasyWord(reportpath)
        word.replace_doc('{Month_Year}',getpreviousmonthfullname(1)+' '+getpreviousyear(1))
        word.replace_footer('{Month_Year}',getpreviousmonthfullname(1)+' '+getpreviousyear(1))
        word.replace_doc('{Current_Date}',getpreviousday(0)+' '+getpreviousmonthfullname(0)+' '+getpreviousyear(0))
        word.replace_doc('{LastMonth_LastDay}',getlastMonthlastDay())
        word.replace_doc('{Month}',getpreviousmonthfullname(1))
        

# 
#         
        df1 = pandas.read_excel(resourcePath2, 'Sheet1', dtype=object,skiprows=2)
#         print(df1.columns)
        
        EnergyProducedTotalNumber = 0
        GrossPR = 0
        NetPR = 0
        Availability = 0
        DownTimeDays = 0
        AdjustedEnergyNumber = 0
        AllInvertersNumber = 0
        rownumber = 0
        currentrownumber = 3
        for i in df1.index:
            currentrownumber =  currentrownumber + 1
            if str(df1['Date'].at[i]) == 'TOT':
                EnergyProducedTotalNumber = df1[df1.columns[1]].at[i]
            if str(df1['Date'].at[i]) == 'AVE':
                GrossPR = df1[df1.columns[3]].at[i] * 100
                NetPR = df1[df1.columns[4]].at[i] * 100
                Availability = df1[df1.columns[5]].at[i] * 100
                rownumber = currentrownumber
            if 'down' in str(df1[df1.columns[7]].at[i]).lower():
                DownTimeDays = DownTimeDays + 1
            if 'Production adjusted' in str(df1[df1.columns[6]].at[i]):
                AdjustedEnergyNumber = df1[df1.columns[7]].at[i]
            if str(df1['Date.1'].at[i]) == 'AVE':
                AllInvertersNumber = df1['All inverters'].at[i]
        
        word.replace_doc('{EnergyProducedTotalNumber}',format(EnergyProducedTotalNumber,',.0f'))
        
        word.replace_doc('{GrossPR}',format(GrossPR,'.1f'))
        word.replace_doc('{NetPR}',format(NetPR,'.1f'))
        word.replace_doc('{Availability}',format(Availability,'.1f'))
        
        word.replace_doc('{DownTimeDays}',format(DownTimeDays,'.0f'))
        word.replace_doc('{AdjustedEnergyNumber}',format(AdjustedEnergyNumber,',.0f'))
        word.replace_doc('{AllInvertersNumber}',format(AllInvertersNumber,'.1f'))
        
           
        
        df2 = pandas.read_excel(resourcePath2, 'Sheet2', dtype=object,skiprows=5)
#         print(df2)
        
        VarNumber = 0
        belowOrAbove = 'below'
        NetPRVarNumber = 0
        belowOrAbove2 = 'below'
        
        rownumber2 = 0
        currentrownumber2 = 6
        for i in df2.index:
            currentrownumber2 = currentrownumber2 + 1
            if getPreviousMonthFullInExcel(1) in str(df2[df2.columns[0]].at[i]):
                VarNumber = df2[df2.columns[4]].at[i]
                NetPRVarNumber = df2[df2.columns[7]].at[i] - df2[df2.columns[8]].at[i]
                rownumber2 = currentrownumber2
        

        if VarNumber > 0:
            belowOrAbove = 'above'
        
        if NetPRVarNumber > 0:
            belowOrAbove = 'above'
        
        VarNumberForWord = round(abs(VarNumber) * 100)
        NetPRVarNumberForWord = abs(NetPRVarNumber) * 100
        
        word.replace_doc('{VarNumber}',format(VarNumberForWord,'.0f')) 
        word.replace_doc('{belowOrAbove}',belowOrAbove)
        
        word.replace_doc('{NetPRVarNumber}',format(NetPRVarNumberForWord,'.1f')) 
        word.replace_doc('{belowOrAbove2}',belowOrAbove)
        
         
        excel = easyexcel.EasyExcel()
        excel.open(resourcePath2)
        
        sheet1 = excel.getSheet(1)
        sheet2 = excel.getSheet(2)
        
        word.copyChartFromExcelToWord(sheet1,1,'chart_1',-1,-1,0)
        
        word.copyChartFromExcelToWord(sheet2,1,'chart_2',-1,-1,0) 
        word.copyChartFromExcelToWord(sheet2,2,'chart_3',-1,-1,0) 
        word.copyChartFromExcelToWord(sheet2,3,'chart_4',-1,-1,0) 
        word.copyChartFromExcelToWord(sheet2,4,'chart_5',-1,-1,0)
        word.copyChartFromExcelToWord(sheet1,7,'chart_6',-1,-1,0)
        word.copyChartFromExcelToWord(sheet1,2,'chart_7',-1,-1,0)
        word.copyChartFromExcelToWord(sheet1,3,'chart_8',-1,-1,0)
        word.copyChartFromExcelToWord(sheet1,4,'chart_9',-1,-1,0)
        word.copyChartFromExcelToWord(sheet1,6,'chart_ten',-1,-1,0)
        word.copyChartFromExcelToWord(sheet1,5,'chart_eleven',-1,-1,0)
        
        excelrange = excel.getRange(1, 3, 1, rownumber, 8)
        
        word.copyTableFromExcelToWordWithFormat(excelrange, 'Table_1')
        
        excelrange = excel.getRange(2, 4, 1, rownumber2, 16)
        
        word.copyTableFromExcelToWordAsPicture(excelrange, 'Table_2')           
        
        df3 = pandas.read_csv(resourcePath3, dtype=object,sep=',',error_bad_lines=False,encoding='gb18030')
        
#         print(df3)
        
        

        for i in df3.index:
            if getlastMonthlastDayInExcel() in str(df3[df3.columns[0]].at[i]):
                print(df3[df3.columns[1]].at[i])
                print(df3[df3.columns[3]].at[i])
                TotalEnergyNumber = int(df3[df3.columns[1]].at[i]) + int(df3[df3.columns[3]].at[i])
                
        
        print(TotalEnergyNumber)
                
        word.replace_doc('{TotalEnergyNumber}',format(TotalEnergyNumber,',.0f'))  

        
    except Exception as e:
       print("general function error happen")
    finally:
       word.close()
       excel.close(0)
       

        
    
    print('Update the monthly word report end')


if __name__ == '__main__':
    
    generateReport()
