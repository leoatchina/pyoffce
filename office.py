# -*- coding: utf-8 -*-
# @Author: leoatchina
# @Date:   2016-06-21 10:02:12
# @Last Modified by:   leoatchina
# @Last Modified time: 2016-11-04 12:50:39



import sys,os,time
import win32com.client as win32
from win32com.client import constants
from tools import alertbox



class Excel(object):
    def __del__(self):
        self.xlApp.Quit()
        del self.xlApp

    def __getattr__(self, name):
        try:
            return getattr(self.xlApp, name)
        except AttributeError:
            raise AttributeError(
                "'%s' object has no attribute '%s'" % (type(self).__name__, name))
    def __init__(self,wkbook="template.xltx",visible=True):
        self.xlApp=win32.gencache.EnsureDispatch('Excel.Application')
        self.xlApp.DisplayAlerts=False
        self.path=sys.path[0]# 当前目录
        filePath=r"%s\%s"%(self.path,wkbook)
        for each in self.xlApp.Workbooks: #如果已经打开这个文件的话直接返回。
            if each.FullName==filePath:
                self.wkBook=each
                return None
        # 打开这个文件，或者新建一个文件
        if os.path.isfile(filePath):
            self.wkBook=self.xlApp.Workbooks.Open(filePath)
            self.xlApp.Visible=visible
        else:
            self.wkBook=self.xlApp.Workbooks.Add()
            self.xlApp.Visible=visible
        return None

    def copyRangeToSheet(self,myRange,sheet,row=1,column=1):
        rowCount=myRange.Rows.Count
        columnCount=myRange.Columns.Count
        myRange.Copy()
        sheet.Activate()
        sheet.Cells(row,column).Select()
        sheet.Paste()
        myRange=sheet.Range(sheet.Cells(row,column),sheet.Cells(row+rowCount-1,column+columnCount-1))
        return myRange

    def setColumnStyle(self,myRange,*styleList):
        try:
            length=len(styleList)
            assert not length % 2,"输入个数不成对匹配"
            sheet=myRange.Parent
            row=myRange.Row
            column=myRange.Column
            rowCount=myRange.Rows.Count
            columnCount=myRange.Columns.Count
            for i in range(0,length,2):
                columnName=styleList[i]
                style=styleList[i+1]
                col=self.findInRangeColumn(myRange,columnName)
                if col:#找到了对应的列号
                    myRange.Range(myRange.Cells(2,col),myRange.Cells(rowCount,col)).NumberFormatLocal = style

        except Exception as e:
            print e
        finally:
            return myRange

    def filterRange(self,myRange,*filterList):
        try:
            length=len(filterList)
            assert not length % 2,"输入个数不成对匹配"

            for i in range(0,length,2):
                columnName=filterList[i]
                criteria1=filterList[i+1]
                col=self.findInRangeColumn(myRange,columnName)
                if col:#找到了对应的列号
                    if r"," in criteria1:
                        ls=criteria1.split(r",")
                        ls=[str(i) for i in ls]
                        myRange.AutoFilter(Field=col,Criteria1=ls,Operator=constants.xlFilterValues)
                    else:
                        myRange.AutoFilter(Field=col,Criteria1=criteria1,Operator=constants.xlAnd)
        except Exception as e:
            print e
        finally:
            return myRange




    def mergeRange(self,myRange,columnName):
        try:
            col=self.findInRangeColumn(myRange,columnName)
            if col:
                rowCount=myRange.Rows.Count
                columnCount=myRange.Columns.Count
                startRow=2
                for i in range(2,rowCount):
                    if myRange.Cells(i,col).Value!=myRange.Cells(i+1,col).Value and i!=startRow:
                        myRange.Range(myRange.Cells(startRow,col),myRange.Cells(i,col)).Merge()
                        startRow=i+1
                return True
            else:
                raise Exception("canot find the %s"%columnName)
        except Exception as e:
            print e
            return False



    def seperateRange(self,myRange,columnName):
        try:
            returnList = None
            col=self.findInRangeColumn(myRange,columnName)
            if col:
                row=myRange.Row
                column=myRange.Column
                headValue=myRange.Rows(1).Value
                rowCount=myRange.Rows.Count
                columnCount=myRange.Columns.Count
                returnList=[rowCount+1]
                for i in range(rowCount-1,1,-1):
                    cellValue=myRange.Cells(i,col).Value.encode('utf-8')
                    cmpCellValue=myRange.Cells(i+1,col).Value.encode('utf-8')
                    if  cellValue!=cmpCellValue:
                        returnList.append(i+1) #这个i+1是要插入的位置
                for eachRow in returnList[1:]: #rowCount+1外其他都要插入行
                    myRange.Rows(eachRow).Insert(Shift=constants.xlDown)
                    myRange.Range(myRange.Cells(eachRow,1),myRange.Cells(eachRow,columnCount)).Value=headValue
                addList=range(0,len(returnList))#每个位置插入的行数
                returnList.reverse()#反转returnList
                returnList=[x+y for (x,y) in zip(addList,returnList)]
                returnList.insert(0,1)# 补充第一行
                returnList=[i+row-1 for i in returnList]
        except Exception as e:
            print e
            returnList = None
        finally:
            return returnList


    def sortRange(self,myRange,*orderList):
        try:
            sheet=myRange.Parent
            length=len(orderList)
            #由于不能直接输入列第一行作搜索，现在只能用下面的assert方法能判断匹配
            assert not length % 2,"输入个数不成对匹配"
            bSort=False
            rowCount=myRange.Rows.Count
            columnCount=myRange.Columns.Count
            row=myRange.Row
            column=myRange.Column
            sheet.Sort.SortFields.Clear()
            for i in range(0,length,2):
                columnName=orderList[i]
                customOrder=orderList[i+1]
                col=self.findInRangeColumn(myRange,columnName)
                if col:#找到了对应的列号
                    bSort=True
                    keyRange=sheet.Range(sheet.Cells(row+1,column+col-1),sheet.Cells(row+rowCount-1,column+col-1))
                    if isinstance(customOrder,str):
                        customOrder=customOrder.decode("utf-8").encode("gbk")
                        sheet.Sort.SortFields.Add(
                            Key=keyRange,
                            SortOn=constants.xlSortOnValues,
                            Order=constants.xlAscending,
                            CustomOrder=customOrder,
                        )
                    else:
                        sheet.Sort.SortFields.Add(
                            Key=keyRange,
                            SortOn=constants.xlSortOnValues,
                            Order=customOrder,
                        )
            if bSort:
                sheet.Sort.Header=constants.xlYes
                sheet.Sort.MatchCase=False
                sheet.Sort.Orientation=constants.xlTopToBottom
                sheet.Sort.SetRange(myRange)
                sheet.Sort.Apply()
        except Exception as e:
            raise(e)
        else:
            return myRange



    def getChart(self,sheet,chartName):
        chartName=chartName.decode("utf-8").encode("gbk")
        sheet.Activate()
        return sheet.ChartObjects(chartName).Chart

    def getCell(self,sheetName,row,col):
        sht=self.getSheet(sheetName)
        sht.Cells(row,col).Select()
        return sht.Cells(row,col)

    def getCellValue(self,sheetName,row,col):
        return getCell(self,sheetName,row,col).Value

    def setCellValue(self,sheetName,row,col,value):
        cell=getCell(self,sheetName,row,col)
        cell.Value=value.decode("utf-8").encode("gbk")


    def getRange(self,sheet,row,column, height,width):
        return sheet.Range(
                    sheet.Cells(row,column),
                    sheet.Cells(row+height-1,column+width-1)
            )



    def findInRangeColumn(self,Range,columnName):
        if Range:
            if columnName:
                row=Range.Rows(1)
                valueList=[i.encode("utf-8") for i in row.Value[0]]
                count=0
                for eachValue in valueList:
                    count=count+1
                    if eachValue==columnName:
                        return count
                    else:
                        continue
                return 0
            else:
                return 0
        else:
            return 0

    def activateSheet(self,sheetName):
        wkSheet=self.getSheet(sheetName)
        if wkSheet:
            self.wkBook.Worksheets(sheetName).Activate()
            return True
        else:
            return False


    def getSheet(self,sheetName):
        try:
            wkSheet=self.wkBook.Worksheets(sheetName)
        except Exception as e:
            return None
        else:
            return wkSheet


    def addSheet(self,sheetName="Test"):
        sheet=self.getSheet(sheetName)
        if sheet :
            return sheet
        else:
            sheet=self.wkBook.Sheets.Add()
            sheet.Name=sheetName
            return sheet


    def close(self):
        if self.wkBook:
            self.wkBook.Close(SaveChanges=0)
        self.wkBook=None
        self.xlApp.DisplayAlerts=True
        self.xlApp.Quit()

    def checkOpen(self,name):
        for each in self.xlApp.Workbooks:
            if each.FullName==name:
                return True
        return False

    def importCsv(self,filename):
        sheetName=filename.split("/")[-1].split(".")[0]
        print sheetName
        bWorkSheet=True
        for eachSheet in self.wkBook.Sheets:
            if eachSheet.Name==sheetName:#存在该sheet，清除之，为后面作准备
                wkSheet=eachSheet
                wkSheet.Select()
                wkSheet.UsedRange.Select()
                self.xlApp.Selection.ClearContents()
                bWorkSheet=False
                break
        if  bWorkSheet:
            wkSheet=self.wkBook.Sheets.Add()
            wkSheet.Name=sheetName
        qt=wkSheet.QueryTables.Add("TEXT;"+filename,wkSheet.Cells(1,1))
        qt.FieldNames = True
        qt.RowNumbers = False
        qt.FillAdjacentFormulas = False
        qt.PreserveFormatting = True
        qt.RefreshOnFileOpen = False
        #qt.RefreshStyle = xlInsertDeleteCells
        qt.SavePassword = False
        qt.SaveData = True
        qt.AdjustColumnWidth = False
        qt.RefreshPeriod = 0
        qt.TextFilePromptOnRefresh = False
        qt.TextFilePlatform = 65001
        qt.TextFileStartRow = 1
        # .TextFileParseType = xlDelimited
        # .TextFileTextQualifier = xlTextQualifierDoubleQuote
        qt.TextFileConsecutiveDelimiter = False
        qt.TextFileTabDelimiter = True
        qt.TextFileSemicolonDelimiter = False
        qt.TextFileCommaDelimiter = True
        qt.TextFileSpaceDelimiter = False
        qt.TextFileColumnDataTypes = [1, 1, 1, 1]
        qt.TextFileTrailingMinusNumbers = True
        qt.Refresh()
        return wkSheet





    def saveAs(self,filename,OverWrite=False): #
        if self.wkBook is None:
            alertbox("wkbook没有提供")
            return False
        if len(filename)>0:
            filename=filename.replace("\\\\",'\\').replace("/","\\")
            fileList=filename.split("/")
            wkBookname=fileList[-1]
            filepath="\\".join(fileList[:-1])
            if filepath:
                if os.path.isdir(filepath):
                    self.path=filepath
                else:
                    alertbox("对应目录不存在")
                    return False
            if OverWrite:
                pass
            else:
                if self.checkOpen(filename):
                    alertbox(filename+"已经打开")
                    return False
            tempDisplayAlerts=self.xlApp.DisplayAlerts
            self.xlApp.DisplayAlerts=False
            self.wkBook.SaveAs(filename)
            self.xlApp.DisplayAlerts=tempDisplayAlerts
            return filename
        else:
            alertbox("文件名没有提供")
            return False





class Word(object):
    def __init__(self,templateName=r'template.dotx',visible=True):
        try:
            self.wdApp= win32.gencache.EnsureDispatch('Word.Application')
            self.wdApp.Visible=visible
            tempPath=os.path.join(sys.path[0],templateName)
            if os.path.isfile(tempPath):
                self.wdDoc=self.wdApp.Documents.Add(tempPath)
            else:
                print  r"not template file"
                self.wdDoc=None
                exit(0)
        except Exception,e:
            print "exception in word init:",e


    def __getattr__(self, name):
        try:
            return getattr(self.wdApp, name)
        except AttributeError:
            raise AttributeError(
                "'%s' object has no attribute '%s'" % (type(self).__name__, name))


    def close(self):
        if self.wdDoc:
            self.wdDoc.Close(SaveChanges=0)
        self.wdDoc=None
        self.wdApp.Quit()


    def paste(self):
        self.wdDoc.Activate()
        return self.wdApp.Selection.Paste()


    def selectPage(self,page):
        self.wdDoc.Activate()
        if page:
            pass
        else:
            self.wdApp.Selection.EndKey(Unit=constants.wdStory)




    # @try_except
    def insertTable(self,myRange,bkMark=None,merge=False,style=None):

        myRange.Copy()
        self.wdDoc.Activate()
        time.sleep(1) #不加这个，很容易出错
        if bkMark:
            bkMark=self.getBookmark(bkMark)
            bkMark.Range.Paste()
        else:
            self.wdApp.Selection.EndKey(Unit=constants.wdStory)
            self.wdApp.Selection.Paste()
        tbl=self.wdDoc.Tables(self.wdDoc.Tables.Count)
        tbl.Rows(1).HeadingFormat=True
        tbl.Rows(1).Select

        if style:
            tbl.Style=style.decode("utf-8").encode("gbk")
        if merge:
            tbl.Select()
            cmpValue=tbl.Cell(2,1).Range.Text.encode("utf-8")[:-2]
            startRow=2
            endRow=2
            for row in range(3,tbl.Rows.Count+1):
                tbl.Select()
                cellValue=tbl.Cell(row,1).Range.Text.encode("utf-8")[:-2]
                if cellValue==cmpValue or cellValue=="":
                    endRow=row
                else:
                    if endRow>startRow:
                        for i in range(startRow+1,endRow+1):
                            tbl.Cell(i,1).Range.Text=""
                        rng=tbl.Cell(startRow,1).Range
                        rng.End=tbl.Cell(endRow,1).Range.End
                        rng.Select()
                        tbl.Cell(startRow,1).Merge(MergeTo=tbl.Cell(endRow,1))
                    startRow=row
                    endRow=row
                    cmpValue=cellValue
        self.wdApp.Selection.EndKey(Unit=constants.wdStory)

        # self.wdApp.Selection.GoTo(What=constants.wdGoToPage,Which=constants.wdGoToAbsolute,Count=startPage)
        # startPage=self.wdApp.Application.Selection.Information(constants.wdActiveEndPageNumber)
        return tbl

    def insertBreak(self,pgbreak=7):
        self.wdDoc.Activate()
        self.wdApp.Selection.EndKey(Extend=0)
        self.wdApp.Selection.InsertBreak(Type=pgbreak)



    def insertTxt(self,txt,enter=False,style=None):
        #I donot know why constants not works here
        #constants.wdMove=0,constants.wdCharacter=1 , constants.wdExtend=1
        #constants.wdPasteHTML=10
        # 导入一个从网上找到的操作htmlClipbord的python库
        from HtmlClipbord import PutHtml,GetHtml
        PutHtml(txt)
        self.wdDoc.Activate()

        self.wdApp.Selection.EndKey(Extend=0)
        start = self.wdDoc.Characters.Count
        self.wdApp.Selection.PasteSpecial(DataType=10)
        end = self.wdDoc.Characters.Count
        length=end-start
        if enter:
            self.wdDoc.Content.InsertAfter("\n")
            if style:
                self.wdApp.Selection.MoveLeft(Unit = constants.wdCharacter,Count = length,Extend = constants.wdExtend)
                self.wdApp.Selection.Style = style.decode("utf8").encode("gbk")
            self.wdApp.Selection.MoveRight(Unit = constants.wdCharacter,Count=2,Extend=constants.wdMove)
        else:
            if style:
                self.wdApp.Selection.MoveLeft(Unit = constants.wdCharacter,Count = length,Extend = constants.wdMove)
                self.wdApp.Selection.Style = style.decode("utf8").encode("gbk")

        return


    def saveAs(self,filename,OverWrite = False):
        if self.wdDoc is None:
            alertbox("文件名没有提供")
            return False
        if len(filename)>0:
            filename=filename.replace("\\\\",'\\').replace("/","\\")
            fileList=filename.split("/")
            wdDocname=fileList[-1]
            filepath="\\".join(fileList[:-1])
            if filepath:
                if os.path.isdir(filepath):
                    self.path=filepath
                else:
                    alertbox("对应目录不存在")
                    return False
            if OverWrite:
                pass
            else:
                if self.checkOpen(filename):
                    alertbox(filename+"已经打开")
                    return False
            tempDisplayAlerts=self.wdApp.DisplayAlerts
            self.wdApp.DisplayAlerts=False
            self.wdDoc.SaveAs(filename)
            self.wdApp.DisplayAlerts=tempDisplayAlerts
            return filename
        else:
            alertbox("文件名没有提供")
            return False



    def checkOpen(self,name):
        for each in self.wdApp.Documents:
            if each.FullName==name:
                return True
        return False


    def getLastTable(self):
        try:
            self.wdDoc.Activate()
            selection=self.wdApp.Selection
            tbl=selection.Tables(selection.Tables.Count)
        except Exception as e:
            return None
        else:
            return tbl



    def getBookmark(self,bookmark=None):
        if  bookmark:
            return self.wdDoc.Bookmarks(bookmark.decode("utf-8").encode("gbk"))



    def findBookmark(self,bookmark=None):
        if bookmark:
            returnList=[]
            for eachBkmark in self.wdDoc.Bookmarks:
                pass
