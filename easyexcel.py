import time
import os.path
import win32com.client 

class EasyExcel(object):
    '''class of easy to deal with excel'''

    def __init__(self):
        '''initial excel application'''
        self.m_filename = ''
        self.m_exists = False
        self.m_excel = win32com.client.DispatchEx('Excel.Application')
        self.m_excel.DisplayAlerts = False

    def open(self,filename=''):
        '''open excel file'''
        if getattr(self,'m_book',False):
            self.m_book.Close()
        self.m_filename=self.dealPath(filename) or ''
        self.m_exists = os.path.isfile(self.m_filename)
        if not self.m_filename or not self.m_exists:
            self.m_book = self.m_excel.Workbooks.Add()
        else:
            self.m_book=self.m_excel.Workbooks.Open(self.m_filename)

    def reset(self):
        '''reset'''
        self.m_excel = None
        self.m_book = None
        self.m_filename = ''

    def clearclip(self):
        self.m_excel.Application.CutCopyMode = False

    def save(self,newfile=''):
        '''save the excel content'''
        assert type(newfile) is str, 'filename must be type string'
        newfile = self.dealPath(newfile) or self.m_filename
        if not newfile or (self.m_exists and newfile == self.m_filename):
            self.m_book.Save()
            return
        pathname = os.path.dirname(newfile)
        if not os.path.isdir(pathname):
            os.makedirs(pathname)
        self.m_filename = newfile
        self.m_book.SaveAs(newfile)

    def close(self,flag_save=1):
        '''close the application, save default true(1)'''
        try:
            self.m_book.Close(SaveChanges=flag_save)
            self.m_excel.Quit()
        except Exception as e:
            print("Close excel exception!")
            print(e)
        time.sleep(2)
        self.reset()

    def dealPath(self,pathname=''):
        '''deal with windows file path'''
        if pathname:
            pathname = pathname.strip()    #remove white space
        if pathname:
            pathname = r'%s'%pathname  #backslashes are treated as escape charaters without 'r'
            pathname = pathname.replace(r'/','\\')
    
            pathname = os.path.abspath(pathname)
            if pathname.find(":\\") == -1:
                pathname = os.path.join(os.getcwd(),pathname)
        return pathname
    
    def getM_excel(self):
        return self.m_excel
    
    def addSheet(self,sheetname=None):
        '''add new sheet, the name of sheet can be modify, but the workbook can't'''
        sht = self.m_book.Worksheets.Add()
        sht.Name = sheetname if sheetname else sht.Name
        return sht

    def addSheetToEnd(self,sheetname=None):
        '''add new sheet, the name of sheet can be modify, but the workbook can't'''
        sht = self.m_book.Worksheets.Add(Before = None, After=self.m_book.Worksheets(self.m_book.Worksheets.count))
        sht.Name = sheetname if sheetname else sht.Name
        return sht

    def getSheet(self,sheet=1):
        '''get the sheet object by the sheet index'''
        #assert sheet>0,'the sheet index must bigger than 0'
        return self.m_book.Worksheets(sheet)

    def getSheetCount(self):
        '''get the number of sheet'''
        return self.m_book.Worksheets.Count
    
    def getSheetNames(self):
        '''get the number of sheet'''
        sheetnames = []
        for sheet in self.m_book.Worksheets:
            sheetnames.append(sheet.name)
        return sheetnames

    def getMaxRow(self,sheet):
        '''get the max row number, not the count of used row number'''
        return self.getSheet(sheet).Rows.Count

    def getUsedRowNumber(self,sheet):
        return self.getSheet(sheet).UsedRange.Rows.Count

    def getMaxCol(self,sheet):
        '''get the max col number, not the count of used col number'''
        return self.getSheet(sheet).Columns.Count

    def getRange(self,sheet,row1,col1,row2,col2):
        '''get the range object'''
        sht = self.getSheet(sheet)
        return sht.Range(self.getCell(sheet,row1,col1),self.getCell(sheet,row2,col2))

    def getRangeValue(self,sheet,row1,col1,row2,col2):
        '''return a tuples of  value'''
        return self.getRange(sheet,row1,col1,row2,col2).Value

    def getCell(self,sheet=1,row=1,col=1):
        '''get the cell object'''
        assert row>0 and col>0, 'the row and column index must bigger than 0'
        return self.getSheet(sheet).Cells(row,col)

    def getRow(self,sheet=1,row=1):
        '''get the row object'''
        assert row>0,'the row index must bigger than 0'
        return self.getSheet(sheet).Rows(row)

    def getCol(self,sheet=1,col=1):
        '''get the col object'''
        assert col>0,'the col index must bigger than 0'
        return self.getSheet(sheet).Columns(col)

    def getRowValue(self,sheet,row):
        '''get the row values'''
        return self.getRow(sheet,row).Value

    def getColValue(self,sheet,col):
        '''get the column values'''
        return self.getCol(sheet,col).value

    def deleteRows(self,sheet,fromRow,count=1):
        '''delete count rows of the sheet'''
        maxRow = self.getMaxRow(sheet)
        maxCol = self.getMaxCol(sheet)
        endRow = fromRow+count-1
        if fromRow > maxRow or endRow <1:
            return
        self.getRange(sheet,fromRow,1,endRow,maxCol).Delete()
        
    def deleteSheet(self, sheet):
        '''delete the whole sheet'''
        self.getSheet(sheet).Delete()
        
    def getCellValue(self,sheet,row,col):
        '''Get value of one cell'''
        return self.getCell(sheet,row,col).Value
    
    def runMacro(self,module_name,macro_name):
        '''run macro'''
        self.m_excel.Application.Run(module_name + "." + macro_name)
        
    def refreshAll(self):
        '''run macro'''
        self.m_book.RefreshAll()
        self.m_excel.CalculateUntilAsyncQueriesDone()

    def getTextboxValue(self, sheet, textbox_name):
        content = ''
        find = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name == textbox_name:
                content = shape.TextFrame2.TextRange.Text
                find = True
                break
        
        return (find,content)

    def getTextboxRange(self, sheet, textbox_name):
        find = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name == textbox_name:
                text_range = shape.TextFrame2.TextRange
                find = True
                break

        return (find,text_range)

    def setTextboxValue(self, sheet, textbox_name, content):
        find = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name == textbox_name:
                shape.TextFrame2.TextRange.Text = content
                find = True
                break
                
        return find