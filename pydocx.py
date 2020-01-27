# -*- coding: utf-8 -*-
"""
Created on Sat Jan 24 14:04:27 2020
@author: Tao
"""
import win32com
import time

import win32com.client as wc

'''创建word应用对象'''
word = win32com.client.Dispatch('Word.Application')
'''设置是否显示，0不显示， 1显示'''
word.Visible = 1

'''获取文档对象'''
# 打开指定文档，打开默认文件夹是C:\users\xxx\document
doc = word.Documents.Open(r'D:\pycharmspace\py4office\b.docx')
# doc = word.Documents.Open(FileName=path, Encoding='gbk')
# 创建新的word文档
# new_doc = word.Documents.Add()

'''对段落进行操作'''
print(doc.Paragraphs.count)
print(doc.Paragraphs[28].Range.text)  # 下标从0开始
print(doc.Paragraphs[28])  # 目前这两种方法打印出来的一样

'''使用Documents.Range对象进行选择'''
# myRange1 = doc.Range(0, 0) # Range(0, 0)表示文档开头；Range()表示文档结尾
# myRange1.InsertBefore('Hello word')
# myRange2.InsertAfter('Bye word')
range1 = doc.Range(doc.Paragraphs(2).Range.Start, doc.Paragraphs(3).Range.End)
range2 = doc.Paragraphs(2).Range.Start

'''Selection, 选择特定内容'''
# 可以通过Range.Select()方法来获取，进而通过Selection对象进行复制、粘贴、粗体、格式等操作，例如：
# range1.Select()
# word.Selection.Font.Bold = True
# 也可以通过选择来选取特定内容

''' 现在可以添加表格，但是有个问题是表格默认没有边框
table = myRange1.Tables.Add(doc.Range(0,0),3,5) #建一张表格
table.Rows[0].Cells[0].Range.Text = u"列"
table.Rows[0].Cells[1].Range.Text = u"类型"
table.Rows[0].Cells[2].Range.Text = u"默认值"
table.Rows[0].Cells[3].Range.Text = u"是否为空"
table.Rows[0].Cells[4].Range.Text = u"列备注"
'''

#
# time.sleep(5)
# 查找替换
# word.Selection.Find.ClearFormatting()
# word.Selection.Find.Replacement.ClearFormatting()
# word.Selection.Find.Execute('ll', False, False, False, False, False, True, 1, True, 'ee', 2)
# time.sleep(3)
'''
上面涉及的 11 个参数说明
         (OldStr--搜索的关键字,
         True--区分大小写,
         True--完全匹配的单词，并非单词中的部分（全字匹配）,
         True--使用通配符,
         True--同音,
         True--查找单词的各种形式,
         True--向文档尾部搜索,
         1,
         True--带格式的文本,
         NewStr--替换文本,
         2--替换个数（0表示不替换，1表示只替换匹配到的第一个，2表示全部替换）
'''
word.Selection.Find.ClearFormatting()
# word.Selection.Find.Replacement.ClearFormatting()
# # 循环操作，将每个匹配到的关键词进行换色
# while word.Selection.Find.Execute('ee', False, False, False, False, False, True, 0, True, "", 0):
#     word.Selection.Range.HighlightColorIndex = WdColorIndex.wdYellow.value  # 替换背景颜色为绿色
#     word.Selection.Font.Color = 255  # 替换文字颜色为红色

'''
目前有两种方法来定位要写入数据的表格，一种是在模板中事先将表格创建好，利用table[index]方法来定位表格再进行修改
还有一种方法待研究，看看能否定位到表格中的特定内容的位置。
'''
# doc.PrintOut() 暂时没看到有什么用，
# doc.Save()
# 保存为PDF，但目前尝试不好使
# doc.SaveAs('xxx.pdf', 17)  # 这行代码目前试用是报错的。
doc.Close()  # 关闭 word 文档
# word.Documents.Close(wc.wdDoNotSaveChanges) # 保存并关闭 word 文档
word.Quit()  # 关闭 office
