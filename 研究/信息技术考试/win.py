# -*- coding: utf-8 -*-
__author__ = 'tqs'
from win32com.client import Dispatch  
import win32com.client 
import time
import os
import re
import win32api
'''
windows操作部分说明：
考试波及知识点：
1.删除文件及文件夹
2.复制文件及文件夹
3.移动文件及文件夹
4.文件及文件夹改名
5.文件属性
考试样例：
1、在“蕨类植物”文件夹中，新建一个子文件夹“薄囊蕨类”。
2、将文件“淡水藻.ddd”移动到“藻类植物”文件夹中。
3、设置“螺旋藻.aaa”文件属性为“只读”。
4、在桌面上建立“绿色植物”的快捷方式。
'''
class WinOperation:
    def __init__(self):
            self.soucePath = ''
            self.destPath = ''
            self.destFilename = ''
            self.sourceFilename = ''
    def dele(self,destFilename):#删除文件及文件夹
        print('删除文件',destFilename)
        pass
    def rename(self,sourceFilename,destFilename):#文件改名
        print(sourceFilename,'文件改名为',destFilename)
        pass
    def mov(self,sourceFilename,destFilename):#移动文件
        print(sourceFilename,'移动文件为',destFilename)
        pass
    def copy(self,sourceFilename,destFilename):#复制文件
        print(sourceFilename,'移动文件为',destFilename)
        pass
    def prop(self,destFilename):#文件属性
        print('文件属性',destFilename)
        pass
    def realSourceFilename(self,soucePath,sourceFilename):
        return sourceFilename
    def realdestFilename(self,destPath,destFilename):
        return destFilename
    def judgeNew(self,OperStr):#从文本中判断新建文件或文件夹
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
        pass
    def judgeDele(self,OperStr):#从文本中判断删除文件
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        pass
    def judgeRename(self,OperStr):#从文本中判断重命名文件
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
        pass
    def judgeMov(self,OperStr):#从文本中判断移动文件
        #形如将文件“淡水藻.ddd”移动到“藻类植物”文件夹中。这种结构的解析
        #解析为源文件，目标文件
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
        #需要再取得完整路径，需要查找
        sourceFilename=self.realSourceFilename("d:\zrexam\windows",source)
        destFilename=self.realdestFilename("d:\zrexam\windows",dest)
        self.mov(sourceFilename,destFilename)
    def judgeCopy(self,OperStr):
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
        pass
    def judgeProp(self,OperStr):
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
##        win32api.SetFileAttributes(fileName,win32con.FILE_ATTRIBUTE_HIDDEN)
##        win32api.SetFileAttributes(fileName,win32con.FILE_ATTRIBUTE_NORMAL)
        pass
    def judgeOperFromList(self,OperStrList):#根据各小题选择对应的操作
        for item in OperStrList:
            pass
    def getOperStrListFromFile(self,filename):#从文件中将各小题放入列表        
        pass
    def judgeOperFromStr(self,OperStr):#根据小题文本选择对应的操作
        if OperStr.find("新建") !=-1:
            print("进入新建操作")
            self.judgeNew(OperStr)
            print("结束新建操作")
        if OperStr.find("删除") !=-1:
            print("进入删除操作")
            self.judgeDele(OperStr)
            print("结束删除操作")
        if OperStr.find("复制") !=-1:
            print("进入复制操作")
            self.judgeCopy(OperStr)
            print("结束复制操作")
        if OperStr.find("移动") !=-1:
            print("进入移动操作")
            self.judgeMov(OperStr)
            print("结束移动操作")
        if OperStr.find("改名") !=-1:
            print("进入改名操作")
            self.judgeRename(OperStr)
            print("结束改名操作")
        if OperStr.find("属性") !=-1:
            print("进入属性操作")
            self.judgeProp(OperStr)
            print("结束属性操作")
            
'''
word操作部分说明：
考试波及知识点：
1.字体
2.段落
3.查找替换
4.插入 表格，艺术字，图片
5.页边距，分栏

1. 将标题“师恩难忘”设置为黑体，居中对齐。
2．将文中第二段（这个小学设在一座庙内……）设置为首行缩进2字符。
3．将文中所有的“田老师”替换为“田先生”。
4. 设置页边距为上下各2.5厘米（应用于整篇文档）。
5. 在正文下面的空白处插入艺术字，内容为“师恩难忘”（样式任选）。
考试样例：
'''
class WordOperation:
    def __init__(self, filename=None):  #打开文件或者新建文件（如果不存在的话）
          self.wordApp = win32com.client.Dispatch('Word.Application')  
          if filename:  
              self.filename = filename
          else:
              self.filename = ''
    def save(self, newfilename=None):  #保存文件
          if newfilename:  
              self.filename = newfilename
          else:
              pass    
    def close(self):  #关闭文件
          del self.wordApp  
    def fontOper(self):        
        pass
    def replaceOper(self,source,dest):
        pass
    def insertOper(self,style):
        pass
    def pageOper(self):
        pass
    def paragraphOper(self):
        pass
    def judgePage(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeFont(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeReplace(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeInsert(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeParagraph(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeOperFromStr(self,OperStr):#根据小题文本选择对应的操作
        if OperStr.find("标题") !=-1 or OperStr.find("黑体") !=-1 or OperStr.find("居中对齐") !=-1:
            print("进入字体操作")
            self.judgeFont(OperStr)
            print("结束字体")
        elif OperStr.find("首行缩进") !=-1 or OperStr.find("行距") !=-1:
            print("进入段落操作")
            self.judgeParagraph(OperStr)            
            print("结束段落操作")
        elif OperStr.find("插入") !=-1:
            print("进入插入操作")
            self.judgeInsert(OperStr)
            print("结束插入操作")
        elif OperStr.find("页边距") !=-1:
            print("进入页边距操作")
            self.judgePage(OperStr)
            print("结束页边距操作")
        elif OperStr.find("分栏") !=-1:
            print("进入分栏操作")
            self.judgeFont(OperStr)
            print("结束分栏操作")
        elif OperStr.find("替换") !=-1:
            print("进入替换操作")
            self.judgeReplace(OperStr)
            print("结束替换操作")
            
'''
Excel操作部分说明：
考试波及知识点：
1.行高列宽
2.格式相关
3.公式函数
4.排序
5.插入图表

考试样例：
1.将A2所在行的行高设置为30（40像素）。
2.根据工作表中提供的公式，计算各班级的“3D社团参与比例”，并将结果填写在F3:F7单元格内。
3.给A2:F8单元格区域加所有框线。
4.按“无人机社团人数”由高到低排序。
5.选定A2:B7单元格区域，制作“三维折线图”，并插入到Sheet1工作表中。

'''
class ExcelOperation:
    def __init__(self, filename=None):  #打开文件或者新建文件（如果不存在的话）
          self.xlApp = win32com.client.Dispatch('Excel.Application')  
          if filename:  
              self.filename = filename  
              self.xlBook = self.xlApp.Workbooks.Open(filename)  
          else:  
              self.xlBook = self.xlApp.Workbooks.Add()  
              self.filename = ''
    def save(self, newfilename=None):  #保存文件
          if newfilename:  
              self.filename = newfilename  
              self.xlBook.SaveAs(newfilename)  
          else:  
              self.xlBook.Save()      
    def close(self):  #关闭文件
          self.xlBook.Close(SaveChanges=0)  
          del self.xlApp  
    def getCell(self, sheet, row, col):  #获取单元格的数据
          "Get value of one cell"  
          sht = self.xlBook.Worksheets(sheet)  
          return sht.Cells(row, col).Value  
    def setCell(self, sheet, row, col, value):  #设置单元格的数据
          "set value of one cell"  
          sht = self.xlBook.Worksheets(sheet)  
          sht.Cells(row, col).Value = value
    def setCellformat(self, sheet, row, col):  #设置单元格的数据
          "set value of one cell"  
          sht = self.xlBook.Worksheets(sheet)  
          sht.Cells(row, col).Font.Size = 15#字体大小
          sht.Cells(row, col).Font.Bold = True#是否黑体
          sht.Cells(row, col).Font.Name = "Arial"#字体类型
          sht.Cells(row, col).Interior.ColorIndex = 3#表格背景
          #sht.Range("A1").Borders.LineStyle = xlDouble
          sht.Cells(row, col).BorderAround(1,4)#表格边框
          sht.Rows(3).RowHeight = 30#行高
          sht.Cells(row, col).HorizontalAlignment = -4131 #水平居中xlCenter
          sht.Cells(row, col).VerticalAlignment = -4160 #
    def rowHeightOper(self,sheet,row,height):
        sht = self.xlBook.Worksheets(sheet)
        sht.Rows(row).RowHeight = height                
    def deleteRow(self, sheet, row):
          sht = self.xlBook.Worksheets(sheet)
          sht.Rows(row).Delete()#删除行
          sht.Columns(row).Delete()#删除列
    def getRange(self, sheet, row1, col1, row2, col2):  #获得一块区域的数据，返回为一个二维元组
          "return a 2d array (i.e. tuple of tuples)"  
          sht = self.xlBook.Worksheets(sheet)
          return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value  
    def addPicture(self, sheet, pictureName, Left, Top, Width, Height):  #插入图片
          "Insert a picture in sheet"  
          sht = self.xlBook.Worksheets(sheet)  
          sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)
    def cpSheet(self, before):  #复制工作表
          "copy sheet"  
          shts = self.xlBook.Worksheets  
          shts(1).Copy(None,shts(1))
    def judgeRowHeight(self,OperStr):#行高操作
        print('正在完成要求',OperStr)
    def judgeColWidth(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeFormula(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeFunction(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeSort(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeChart(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeBoxLine(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeOperFromStr(self,OperStr):#根据小题文本选择对应的操作
        if OperStr.find("行高") !=-1:
            print("进入行高操作")
            self.judgeRowHeight(OperStr)
            print("结束行高操作")
        if OperStr.find("列宽") !=-1:
            print("进入列宽操作")
            self.judgeColWidth(OperStr)
            print("结束列宽操作")
        if OperStr.find("公式") !=-1:
            print("进入公式操作")
            self.judgeFormula(OperStr)
            print("结束公式操作")
        if OperStr.find("函数") !=-1:
            print("进入函数操作")
            self.judgeFunction(OperStr)
            print("结束函数操作")
        if OperStr.find("所有框线") !=-1:
            print("进入所有框线操作")
            self.judgeBoxLine(OperStr)
            print("结束所有框线操作")
        if OperStr.find("排序") !=-1:
            print("进入排序操作")
            self.judgeSort(OperStr)
            print("结束排序操作")
        if OperStr.find("图表") !=-1:
            print("进入图表操作")
            self.judgeChart(OperStr)
            print("结束图表操作")
        pass
    
'''
PPT操作部分说明：
1.动画效果
2.切换效果
3.超级链接
4.背景
5.插入，图片，声音，视频

考试样例：
1.在第四张幻灯片的上方插入横排文本框，在文本框中输入“吃月饼”，字体黑体，字号32。
2.将第三张幻灯片的背景填充效果设置为纹理填充，纹理为“鱼类化石”。
3.设置第三张幻灯片的切换效果为“推进”，声音为“鼓掌”。
4.给第四张幻灯片右侧的图片设置进入中的“劈裂”动画效果，效果选项为“中央向上下展开”。
5.给第三张幻灯片中的文字“赏桂花”添加超链接，使其链接到第五张幻灯片。
'''

class PptOperation:
    def __init__(self):
        pass
    def AnimationOper(self):
        pass
    def SwitchOper(self):
        pass
    def InsertOper(self,style):
        pass
    def BackgroundOper(self):
        pass
    def HyperlinkOper(self):
        pass
    def judgeAnimation(self,OperStr):
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
    def judgeSwitch(self,OperStr):
        print('正在完成要求',OperStr)
        pattern = re.compile('“(.*)”')
        print (pattern.findall(OperStr))
        strFile=str(pattern.findall(OperStr))
        file1=strFile.split("”")
        source=file1[0][2:]#获得源文件
        print(source)
        file2=strFile.split("“")
        dest=file2[1][0:-2]#获得目标文件
        print(dest)
    def judgeInsert(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeBackground(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeHyperlink(self,OperStr):
        print('正在完成要求',OperStr)
    def judgeOperFromStr(self,OperStr):#根据小题文本选择对应的操作
        
        if OperStr.find("动画") !=-1:
            print("进入动画操作")
            self.judgeAnimation(OperStr)
            print("结束动画操作")
        if OperStr.find("切换") !=-1:
            print("进入切换操作")
            self.judgeSwitch(OperStr)
            print("结束切换操作")
        if OperStr.find("超级链接") !=-1:
            print("进入超级链接操作")
            self.judgeHyperlink(OperStr)
            print("结束超级链接操作")
        if OperStr.find("背景") !=-1:
            print("进入背景操作")
            self.judgeBackground(OperStr)
            print("结束背景操作")
        if OperStr.find("插入") !=-1:
            print("进入插入操作")
            self.judgeInsert(OperStr)
            print("结束插入操作")
            
'''
Input文字录入操作部分说明：
考试波及知识点：
com对象的调用演示：
class InputOperation:
'''
class OperationTypeJudge:
    def __init__(self):
        pass
    def getType(self,OperStr):
        if OperStr.find("替换") !=-1 or OperStr.find("首行缩进") !=-1:
            print('这是word题要求')
            print('已转word题处理')
        elif OperStr.find("公式") !=-1 or OperStr.find("函数") !=-1:
            print('这是excel题要求')
            print('已转excel题处理')
        elif OperStr.find("切换") !=-1 or OperStr.find("动画") !=-1:
            print('这是ppt题要求')
            print('已转ppt题处理')
        pass
    def getOperaPath(self):
        pass
    def getOperaFileName(self):
        pass
'''
选择题部分说明：
'''    
class SelectOperation:    
    def __init__(self):
        pass    
    def getQusetionTxt(self,item):
        pass
    def getQusetionPic(self,item):
        pass
    def getAnswer(self,item):
        pass
    def getCorrectAnswer(self,item):
        pass
    
'''
判断题部分说明：
'''     
class JudgeOperation:    
    def __init__(self):
        pass    
    def getQusetionTxt(self,item):
        pass
    def getQusetionPic(self,item):
        pass
    def getAnswer(self,item):
        pass
    def getCorrectAnswer(self,item):
        pass    
if __name__ == "__main__":
      win=WinOperation()
      win.judgeOperFromStr('1、在“蕨类植物”文件夹中，新建一个子文件夹“薄囊蕨类”。')
      win.judgeOperFromStr('2、将文件“淡水藻.ddd”移动到“藻类植物”文件夹中。')
      win.judgeOperFromStr('3、设置“螺旋藻.aaa”文件属性为“只读”。')
      win.judgeOperFromStr('4、在桌面上建立“绿色植物”的快捷方式。')

      word=WordOperation()
      word.judgeOperFromStr('1. 将标题“师恩难忘”设置为黑体，居中对齐。')
      word.judgeOperFromStr('2．将文中第二段（这个小学设在一座庙内……）设置为首行缩进2字符。')
      word.judgeOperFromStr('3．将文中所有的“田老师”替换为“田先生”。')
      word.judgeOperFromStr('4. 设置页边距为上下各2.5厘米（应用于整篇文档）。')
      word.judgeOperFromStr('5. 在正文下面的空白处插入艺术字，内容为“师恩难忘”（样式任选）。')

      excel=ExcelOperation(r'c:/test.xls')
      excel.judgeOperFromStr('1.将A2所在行的行高设置为30（40像素）。')
      excel.judgeOperFromStr('2.根据工作表中提供的公式，计算各班级的“3D社团参与比例”，并将结果填写在F3:F7单元格内。')
      excel.judgeOperFromStr('3.给A2:F8单元格区域加所有框线。')
      excel.judgeOperFromStr('4.按“无人机社团人数”由高到低排序。')
      excel.judgeOperFromStr('5.选定A2:B7单元格区域，制作“三维折线图”，并插入到Sheet1工作表中。')

      ppt=PptOperation()
      ppt.judgeOperFromStr('1.在第四张幻灯片的上方插入横排文本框，在文本框中输入“吃月饼”，字体黑体，字号32。')
      ppt.judgeOperFromStr('2.将第三张幻灯片的背景填充效果设置为纹理填充，纹理为“鱼类化石”。')
      ppt.judgeOperFromStr('3.设置第三张幻灯片的切换效果为“推进”，声音为“鼓掌”。')
      ppt.judgeOperFromStr('4.给第四张幻灯片右侧的图片设置进入中的“劈裂”动画效果，效果选项为“中央向上下展开”。')
      ppt.judgeOperFromStr('5.给第三张幻灯片中的文字“赏桂花”添加超链接，使其链接到第五张幻灯片。')
