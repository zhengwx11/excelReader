#!/usr/bin/python
# -*- coding=utf-8 -*-
import sys
import xml.sax as SAX
import numpy
import zipfile as ZIP
import StringIO
# @string 形如ABA123的字符串
# @return ABA对应的数字（26进制）
def letterToDec(string):
    ret = 0
    for i in xrange(0, string.__len__()):
        c = ord(string[i])
        if(c >= ord('A') and c <= ord('Z')):
            ret *= 26
            ret += c - ord('A') + 1
        else:
            return ret, i

# @dim,形如"A1:ZA99"的字符串
# @return [L1,L2, N1,N2] ，行列两个方向的index范围
def parseDimension(dim):
    dim = str(dim)
    dims = dim.split(":")
    Rg = [0,0,0,0];
    for index in xrange(0,2):
        ret,i = letterToDec(dims[index])
        Rg[index] = ret
        Rg[index+2] = int(dims[index][i:])
    return Rg


# 读取sheet.xml用的handler
class SheetData(SAX.ContentHandler):
    def __init__(self):
        self.Tag = None # 解析过程中， 记录上一个标签的名字
        self.X = None   # 记录所有 id 对应的column值
        self.Y = None   # 记录所有 id 对应的row值
        self.V = None   # 记录所有的id， 是该位置的数据在sharedString中的索引值
        self.Rng = None # range, X, Y 的范围
        self.Row = 0    # 最近一个的row值
        self.Column = 0 # 最近一个column值
        self.Index = 0  # 解析的进度

    def startElement(self,tag,attributes):
        self.Tag = tag
        if tag == "dimension":
            ref = attributes['ref']
            self.Rng = parseDimension(ref)
            Len = self.Rng[3] - self.Rng[2] + 1
            Width = self.Rng[1] - self.Rng[0] + 1
            self.X = [None]*(Len*Width)
            self.Y = [None]*(Len*Width)
            self.V = [None]*(Len*Width)
        elif(tag == "row"):
            self.Row = int(attributes["r"])
        elif(tag == "c"):
            pos = attributes["r"]
            self.Column , _ = letterToDec(pos)

    def endElement(self, tag):
        if(tag == "v"):
            self.Tag = "VV"
            self.Index += 1

    def characters(self, content):
        if(self.Tag == "v"):
            self.X[self.Index] = self.Column
            self.Y[self.Index] = self.Row
            self.V[self.Index] = int(content)

# 读取sharedString的handler
class SharedString(SAX.ContentHandler):
    def __init__(self, length):
        self.datas = [None]*length
        self.Tag = None
        self.Index = 0
        reload(sys)
        sys.setdefaultencoding('utf-8')

    def startElement(self, tag, attr):
        self.Tag = tag

    def endElement(self,tag):
        if(tag == "t"):
            self.Index += 1

    def characters(self, content):
        if self.Tag == "t":
            if(self.datas[self.Index] is None):
                self.datas[self.Index] = str(content)
            else:
                self.datas[self.Index] += str(content)

# 使用handler解析filename
def parseXml(src, handler):
    parser = SAX.make_parser()
    parser.setFeature(SAX.handler.feature_namespaces,0)

    parser.setContentHandler(handler)
    parser.parse(src)
    return handler

# 根据两个xml的解析结果，得到一个字符串组成的二维表
def mergeTable(sheet, data):
    width = sheet.Rng[1] - sheet.Rng[0] + 1
    height = sheet.Rng[3] - sheet.Rng[2] + 1
    # python的二维数组初始化有特殊技巧
    # 随便来则容易出现几个维度公用内存的情况
    tables = [[None] * width for row in xrange(0,height)]

    for i in xrange(0, sheet.X.__len__()):
        x = sheet.X[i] - sheet.Rng[0]
        y = sheet.Y[i] - sheet.Rng[2]
        ID = sheet.V[i]
        tables[y][x] = data.datas[ID]
    return tables

# 给定两个文件流（sheet 和 sharedString）， 得出二维数组
def getDataTable(file1, file2):
    sheet_handler = SheetData()
    sheet_handler = parseXml(file1,sheet_handler)

    data_handler = SharedString(sheet_handler.Index)
    data_handler = parseXml(file2,data_handler)

    return mergeTable(sheet_handler, data_handler)


# 给定xlsx文件， 得出二维数组
def getDataTableFromXLSX(filename):
    # 设置系统使用Utf-8编码，需要reload才能调用此函数

    zipfile = ZIP.ZipFile(filename,allowZip64=True)
    f_sharedString = StringIO.StringIO(zipfile.read("xl/sharedStrings.xml"))
    f_sheet = StringIO.StringIO(zipfile.read("xl/worksheets/sheet1.xml"))
    return getDataTable(f_sheet, f_sharedString)
    
if __name__ == '__main__':
    table = getDataTableFromXLSX(sys.argv[1])
    for row in table:
        for column in row:
            print column
