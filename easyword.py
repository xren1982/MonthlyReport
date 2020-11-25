import os
import win32com.client
import easyexcel


class EasyWord:
    def __init__(self, filename=None):
        self.xlApp=win32com.client.Dispatch('Word.Application')
        self.xlApp.Visible=0
        self.xlApp.DisplayAlerts=0
        if filename:
            self.filename=filename
            if os.path.exists(self.filename):
                self.doc=self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()    
                self.doc.SaveAs(filename)
        else:
            self.doc=self.xlApp.Documents.Add()
            self.filename=''
    def get_doc(self):
        return self.doc
    def get_xlApp(self):
        return self.xlApp
    def add_doc_end(self, string):
        rangee = self.doc.Range()
        rangee.InsertAfter('\n'+string)
    def add_doc_start(self, string):
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string+'\n')
    def insert_doc(self, insertPos, string):
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n'+string)
    def replace_doc(self,string,new_string):
        self.doc.Range().Find.ClearFormatting()
        self.doc.Range().Find.Replacement.ClearFormatting()
        self.doc.Range().Find.Execute(string, False, False, False, False, False, True, 1, True, new_string, 2)
    def copyTableFromExcelToWord(self,sheet,copyrange,replaceinword):
        sheet.Range(copyrange).Copy()
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.PasteExcelTable(False,False,True)
    def copyTableFromExcelToWordNew(self,excelrange,replaceinword):
        excelrange.Copy()
        rng = self.doc.Range()
        is_exists = rng.Find.Execute(FindText= replaceinword)
        if is_exists:
            rng.PasteExcelTable(False,False,True)
    def copyTableFromExcelToWordWithFormat(self,excelrange,replaceinword):
        '''Need to copy the format info like cells merge '''
        excelrange.Copy()
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.PasteExcelTable(False,False,False)
    def copyTableFromExcelToWordAsPicture(self,excelrange,replaceinword):
        excelrange.CopyPicture()
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.Paste()
    def findKeyword(self,keyword,maxreturn):
        rangelist = list()
        for i in range(0,maxreturn):
            newrange = self.doc.Range()
            finded = newrange.Find.Execute(FindText=keyword)
            if not finded:
                break
            self.doc.Range().Find.Execute(keyword, False, False, False, False, False, True, 1, True, 'TempFindKeyWordReplace', 1)
            rangelist.append(newrange)
        self.doc.Range().Find.Execute('TempFindKeyWordReplace',False, False, False, False, False, True, 1, True, keyword, 2)
        return rangelist
    
    def insertPic(self,picpath,replaceinword):
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.InlineShapes.AddPicture(picpath)
        self.replace_doc(replaceinword, '')
    def copyPicFromExcelToWord(self,sheet,picnumber,replaceinword,width=-1,height=-1):
        position = 0
        success = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name.startswith("Picture"):
                position = position + 1
                if(position == picnumber):
                    if(width >= 0):
                        shape.Width = width
                    if(height >= 0):
                        shape.Height = height
                    shape.Copy()
                    rng = self.doc.Range()
                    rng.Find.Execute(FindText= replaceinword)
                    rng.Paste()
                    success = True
        return success
    def copyChartFromExcelToWord(self,sheet,chartnumber,replaceinword,width=-1,height=-1,copyaspicture=0):
        position = 0
        success = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name.startswith("Chart"):
                position = position + 1
                if(position == chartnumber):
                    if(width >= 0):
                        shape.Width = width
                    if(height >= 0):
                        shape.Height = height
                    if copyaspicture == 1:
                        shape.CopyPicture()
                    else:
                        shape.Copy()
                    rng = self.doc.Range()
                    rng.Find.Execute(FindText= replaceinword)
                    rng.Paste()
                    success = True
        return success
    
    def copyGroupFromExcelToWord(self,sheet,chartnumber,replaceinword,width=-1,height=-1,copyaspicture=0):
        position = 0
        success = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name.startswith("Group"):
                position = position + 1
                if(position == chartnumber):
                    if(width >= 0):
                        shape.Width = width
                    if(height >= 0):
                        shape.Height = height
                    if copyaspicture == 1:
                        shape.CopyPicture()
                    else:
                        shape.Copy()
                    rng = self.doc.Range()
                    rng.Find.Execute(FindText= replaceinword)
                    rng.Paste()
                    success = True
        return success

    def copyChartFromExcelToWordwithChartname(self,sheet,chartname,replaceinword,width=-1,height=-1,copyaspicture=0):
        success = False
        for n, shape in enumerate(sheet.Shapes):
            if shape.Name == chartname:
                if(width >= 0):
                    shape.Width = width
                if(height >= 0):
                    shape.Height = height
                if copyaspicture == 1:
                    shape.CopyPicture()
                else:
                    shape.Copy()
                rng = self.doc.Range()
                rng.Find.Execute(FindText= replaceinword)
                rng.Paste()
                success = True
        return success

    def copyPicFromWordToWord(self,resourcexlApp,resourcedoc,resourcedockeyword,replaceinword,width=-1,height=-1):
        position = 0
        success = False
        keywordrange = resourcedoc.Range()
        keywordrange.Find.Execute(FindText= resourcedockeyword)
        for shape in resourcedoc.Range(keywordrange.Start).InlineShapes:
            #print(shape)
            position = position + 1
            if(position == 1):
                if(width >= 0):
                    shape.Width = width
                if(height >= 0):
                    shape.Height = height
                shape.Select()
                resourcexlApp.Selection.Copy()
                rng = self.doc.Range()
                rng.Find.Execute(FindText= replaceinword)
                rng.Paste()
                success = True
        return success
    
    def copyRangeFromWordToWord(self,resourcedoc,startkeyword,endkeyword,replaceinword):
        fromstart = False
        toend = False
        if startkeyword is not None and startkeyword != '':
            rangestart = resourcedoc.Range()
            rangestart.Find.Execute(FindText= startkeyword)
            fromstart = True
        if endkeyword is not None and endkeyword != '':   
            rangeend = resourcedoc.Range()
            rangeend.Find.Execute(FindText= endkeyword)
            toend = True
        if fromstart and toend:
            copyrange = resourcedoc.Range(rangestart.Start,rangeend.End)
        if fromstart and not toend:
            copyrange = resourcedoc.Range(rangestart.Start)
        if not fromstart and toend:
            copyrange = resourcedoc.Range(End=rangeend.End)
        if not fromstart and not toend:
            copyrange = resourcedoc.Range()
        copyrange.Copy()
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.Paste()
    
    def copyRangeFromWordToWordByRangePara(self,copyrange,replaceinword):
        copyrange.Copy()
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.Paste()
        
    def copyRangeFromWordToWord_2(self,resourcedoc,startkeyword,endkeyword,replaceinword):
        fromstart = False
        toend = False
        if startkeyword is not None and startkeyword != '':
            rangestart = resourcedoc.Range()
            rangestart.Find.Execute(FindText= startkeyword)
            fromstart = True
        if endkeyword is not None and endkeyword != '':   
            rangeend = resourcedoc.Range()
            rangeend.Find.Execute(FindText= endkeyword)
            toend = True
        if fromstart and toend:
            copyrange = resourcedoc.Range(rangestart.End+1,rangeend.End)
        if fromstart and not toend:
            copyrange = resourcedoc.Range(rangestart.End+1)
        if not fromstart and toend:
            copyrange = resourcedoc.Range(End=rangeend.End)
        if not fromstart and not toend:
            copyrange = resourcedoc.Range()
        copyrange.Copy()
        rng = self.doc.Range()
        rng.Find.Execute(FindText= replaceinword)
        rng.Paste()
         
    def save(self):
        self.doc.Save()
    def save_as(self, filename):
        self.doc.SaveAs(filename)
    def close(self):
        '''close the application, save default true(1)'''
        try:
            self.save()
            self.doc.Close()
            if(self.xlApp.Documents.Count == 0):
                self.xlApp.Quit()
            else:
                print("[info]Other word file is still be opened. It will not stop the word.exe!")

        except Exception as e:
            print("Close Word exception!")
            print(e)

    def replace_shape_text(self, string, new_string):
        shapes = self.xlApp.ActiveDocument.shapes
        
        for shape in shapes:
            try:
                for item in shape.GroupItems:
                    if item.TextFrame.HasText:
                        if item.TextFrame.TextRange.Text != '\r':
                            item.TextFrame.TextRange.Find.ClearFormatting()
                            item.TextFrame.TextRange.Find.Replacement.ClearFormatting()
                            item.TextFrame.TextRange.Find.Execute(string, False, False, False, False, False, True, 1, False, new_string, 2)
            except Exception as e:
                err = "ignore"

            try:
                if shape.TextFrame.HasText:
                    if shape.TextFrame.TextRange.Text != '\r':
                        shape.TextFrame.TextRange.Find.ClearFormatting()
                        shape.TextFrame.TextRange.Find.Replacement.ClearFormatting()
                        shape.TextFrame.TextRange.Find.Execute(string, False, False, False, False, False, True, 1, False, new_string, 2)
            except Exception as e:
                err = "ignore"


    def replace_header(self, string, new_string):
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
        self.xlApp.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(string, False, False, False, False, False, True, 1, False, new_string, 2)

    def replace_footer(self, string, new_string):
        self.xlApp.ActiveDocument.Sections[0].Footers[0].Range.Find.ClearFormatting()
        self.xlApp.ActiveDocument.Sections[0].Footers[0].Range.Find.Replacement.ClearFormatting()
        self.xlApp.ActiveDocument.Sections[0].Footers[0].Range.Find.Execute(string, False, False, False, False, False, True, 1, False, new_string, 2)