# -*- coding: utf-8 -*-
"""
Created on Sat Jan 25 17:45:27 2020
@author: Tao
"""
import win32com.client
import time

from constants import WdPaperSize


word_app = win32com.client.Dispatch('Word.Application')
# word_app = win32com.client.gencache.EnsureDispatch('Word.Application')
word_app.Visible = 1

# 打开指定文档，打开默认文件夹是C:\users\xxx\document
doc = word_app.Documents.Open(r'D:\pycharmspace\py4office\c.docx')
print(win32com.client.constants.wdPaperEnvelope11)
"""设置纸张大小，这里有两种方法：
1、使用微软体系的枚举参数，但是win32com库比较特殊，代码无法提示，需要自己到MSDN上查询可选的值
2、第二种参照微软的enum，自定义python的enum。优点是import之后有代码提示，缺点是得使用.value函数
   取值，代码格式上显得有点不正宗
"""
# doc.PageSetup.PaperSize = win32com.client.constants.wdPaperEnvelope11
doc.PageSetup.PaperSize = WdPaperSize.wdPaperA3.value

# doc.PageSetup.PageWidth = 21 * 28.35  # 直接设置纸张大小, 使用该设置后PaperSize设置取消
# doc.PageSetup.PageHeight = 29.7 * 28.35  # 直接设置纸张大小
# doc.PageSetup.Orientation = 1  # 页面方向, 竖直=0, 水平=1 doc.PageSetup.TopMargin = 3*28.35
# 页边距上=3cm，1cm=28.35pt
doc.PageSetup.BottomMargin = 3 * 28.35  # 页边距下=3cm doc.PageSetup.LeftMargin = 2.5*28.35 # 页边距左=2.5cm doc.PageSetup.RightMargin = 2.5*28.35 # 页边距右=2.5cm
doc.PageSetup.TextColumns.SetCount(2)  # 设置页面分栏=2
time.sleep(3)
doc.Close()
word_app.Quit()
