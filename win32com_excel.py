# -*- coding: utf-8 -*-
# @Author: Administrator
# @Date:   2016-09-01 09:38:51
# @Last Modified by:   Administrator
# @Last Modified time: 2016-09-01 10:20:27
# http://sourceforge.net/projects/pywin32/files/pywin32/
import win32com.client
import os
import time
from os import listdir
from os.path import isfile, join
from win32com.client import constants
import datetime


class RemoteExcel():
    """对excel表格的操作

    """

    def __init__(self, filename=None):
        """初始化函数

        Args:
            filename: 要进行操作的文件名，如果存在此文件则打开，不存在则新建
                        此文件名一般包含路径

        """
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = 0
        self.xlApp.DisplayAlerts = 0  # 后台运行，不显示，不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.xlBook = self.xlApp.Workbooks.Open(filename)
            else:
                self.xlBook = self.xlApp.Workbooks.Add()  # 创建新的Excel文件
                self.xlBook.SaveAs(self.filename)
        else:
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''

    def get_cell(self, row, col, sheet=None):
        """读取单元格的内容

        Args:
            row: 行
            col: 列
            sheet: 表格名（不是文件名）

        """
        if sheet:
            sht = self.xlBook.Worksheets(sheet)
        else:
            sht = self.xlApp.ActiveSheet
        return sht.Cells(row, col).Value

    def set_cell(self, sheet, row, col, value, comment=""):
        """向表格单元格写入

        Args:
            sheet: 表格名（不是文件名）
            row: 行
            col: 列
            value: 定入内容
        """
        try:
            sht = self.xlBook.Worksheets(sheet)
        except:
            self.new_sheet(sheet)
            sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Value = value
        if comment:
            sht.Cells(row, col).AddComment()
            sht.Cells(row, col).Comment.Text(comment)

    def set_font(self, sheet, row, col, font, size):
        sht = self.xlBook.Worksheets(sheet)
        sht.Cells(row, col).Font.Name = font
        sht.Cells(row, col).Font.Size = size

    def save(self, newfilename=None):
        """保存表格"""
        if newfilename:
            self.filename = newfilename
            self.xlBook.SaveAs(self.filename)
        else:
            self.xlBook.Save()

    def close(self):
        """保存表格、关闭表格，结束操作"""
        self.save()
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def new_sheet(self, newSheetName):
        """新建一个新表格"""
        sheet = self.xlBook.Worksheets.Add()
        sheet.Name = newSheetName
        sheet.Activate()

    def active_sheet(self):
        return self.xlApp.ActiveSheet


class RemoteWord():

    def __init__(self, filename=None):
        self.xlApp = win32com.client.DispatchEx('Word.Application')
        self.xlApp.Visible = 0
        self.xlApp.DisplayAlerts = 0  # 后台运行，不显示，不警告
        if filename:
            self.filename = filename
            if os.path.exists(self.filename):
                self.doc = self.xlApp.Documents.Open(filename)
            else:
                self.doc = self.xlApp.Documents.Add()  # 创建新的文档
                self.doc.SaveAs(filename)
        else:
            self.doc = self.xlApp.Documents.Add()
            self.filename = ''

    def add_doc_end(self, string):
        '''在文档末尾添加内容'''
        rangee = self.doc.Range()
        rangee.InsertAfter('\n' + string)

    def add_doc_start(self, string):
        '''在文档开头添加内容'''
        rangee = self.doc.Range(0, 0)
        rangee.InsertBefore(string + '\n')

    def insert_doc(self, insertPos, string):
        '''在文档insertPos位置添加内容'''
        rangee = self.doc.Range(0, insertPos)
        if (insertPos == 0):
            rangee.InsertAfter(string)
        else:
            rangee.InsertAfter('\n' + string)

    def save(self):
        '''保存文档'''
        self.doc.Save()

    def save_as(self, filename):
        '''文档另存为'''
        self.doc.SaveAs(filename)

    def close(self):
        '''保存文件、关闭文件'''
        self.save()
        self.xlApp.Documents.Close()
        self.xlApp.Quit()


def auto_build_excel(from_dir=None, dest_file=None, month=None, sheet_name=None, password=None):
    # 自动整理每个人的考核表，并整合到一张表中
    if month:
        from_dir = "{}月".format(month)
        if len(str(month)) == 1:
            month = "0" + str(month)
        dest_file = "郑州电商{year}年{month}月英文排名团队绩效考核表.xlsx".format(
            year=datetime.date.today().year, month=month)
        sheet_name = "{year}年{month}月英文排名团队绩效考核表".format(
            year=datetime.date.today().year, month=month)
    onlyfiles = [f for f in listdir(from_dir) if isfile(
        join(from_dir, f)) and not f.startswith("~")]
    for exce_file in onlyfiles:
        excel = RemoteExcel(os.path.abspath(os.path.join(from_dir, exce_file)))
        name = excel.get_cell(2, 2)
        date = excel.get_cell(2, 5)

        score = []
        for i in range(6, 17):
            score.append(int(excel.get_cell(i, 8)))
        # 打印信息
        print(exce_file, name, date, score, sum(score))
        if password:
            excel.xlBook.Unprotect(password)
        # 复制当前sheets到另一文件
        try:
            os.remove(dest_file)
        except OSError:
            pass
        excel2 = RemoteExcel(os.path.abspath(dest_file))
        excel.xlBook.Worksheets(1).Copy(None, excel2.xlBook.Worksheets(1))
        excel2.xlApp.ActiveSheet.Name = name
        excel2.save()
        # excel2.xlBook.SaveAs(os.path.abspath("sample.xlsx"))
        excel.close()
    # 制作汇总表
    comment = """janis:
90~100分=特优 (绩效奖金2500)
80~90分=优秀 (绩效奖金1800)
70~79分=良好 (绩效奖金1300)
60~69分=普通 (绩效奖金800)
59分之下=不良 (绩效奖金0)
    """
    ws = excel2.xlBook.Worksheets(1)
    ws.Name = sheet_name
    # 单元数据表头
    excel2.set_cell(1, 2, 2, "团队")
    excel2.set_cell(1, 2, 3, "姓名")
    excel2.set_cell(1, 2, 4, "考核得分")
    excel2.set_cell(1, 2, 5, "最终绩效", comment=comment)
    excel2.set_cell(1, 3, 2, "英文排名")
    # 注解
    ws.Cells(2, 5).Comment.Shape.Width = 240
    ws.Cells(2, 5).Comment.Shape.Height = 115
    ws.Cells(2, 5).Comment.Shape.TextFrame.Characters().Font.Bold = False
    ws.Cells(2, 5).Comment.Shape.TextFrame.Characters(1, 7).Font.Bold = True

    # 填充具体数据
    name_list = ["耿二轩", "黄儒卿", "刘鹏雁", "马芳", "马永红", "谢亚丹", "岳静丽", "张九杰"]
    for i in range(3, 11):
        excel2.set_cell(1, i, 3, name_list[i - 3])
        excel2.set_cell(
            1, i, 4, '=INDIRECT(CONCATENATE(C{row_num},"!I25"))'.format(row_num=i))
        excel2.set_cell(
            1, i, 5, '=IF(D{row_num}>89,2500,IF(D{row_num}>79,1800,IF(D{row_num}>69,1300,IF(D{row_num}>59,800,0))))'.format(row_num=i))
    # 表的全局设置
    excel2.xlApp.Worksheets(1).Activate()
    excel2.xlApp.ActiveWindow.DisplayGridlines = False
    excel2.xlApp.Worksheets(1).Tab.Color = "&hFF"
    # 字体及单元格合并
    ws.Rows(2).RowHeight = 29
    ws.Rows(2).Font.Bold = True
    ws.Range("3:10").RowHeight = 25
    ws.Range("B3:B10").MergeCells = True
    # 列
    ws.Columns(1).Font.Bold = True
    ws.Columns(1).ColumnWidth = 2.75
    ws.Columns(2).ColumnWidth = 14.5
    ws.Columns(3).ColumnWidth = 13.5
    ws.Columns(4).ColumnWidth = 16.25
    ws.Columns(5).ColumnWidth = 18.25
    ws.Rows(1).RowHeight = 14.25
    # 对齐
    ws.Range("B2:E10").HorizontalAlignment = constants.xlCenter
    ws.Range("D3:D10").HorizontalAlignment = constants.xlRight
    # 每个单元边框
    ws.Range("B2:E10").Borders.LineStyle = 1
    ws.Range("B2:E10").Borders.Weight = 2
    # 最外边框粗体显示
    ws.Range("B2:B10").Borders(1).Weight = 3
    ws.Range("B2:E2").Borders(3).Weight = 3
    ws.Range("E2:E10").Borders(2).Weight = 3
    ws.Range("B10:E10").Borders(4).Weight = 3

    excel2.save()
    excel2.close()

if __name__ == '__main__':
    # example1
    import glob
    TODAY = time.strftime('%Y-%m-%d',    time.localtime(time.time()))
    filename = "郑州电商2016年07月英文排名团队绩效考核表.xlsx"
    # filename = "sample.xlsx"
    month = 8
    # excel = RemoteExcel(os.path.abspath(filename))
    # test = excel.xlBook.Worksheets(1).Cells(2, 5).Comment.Text()
    # print(dir(test), test)
    auto_build_excel(month=8)
