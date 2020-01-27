import win32com
import win32com.client
from constants import WdColorIndex

word = win32com.client.gencache.EnsureDispatch('Word.Application')
doc = word.Documents.Open(r'D:\pycharmspace\py4office\c.docx')
range1 = doc.Range(doc.Paragraphs(2).Range.Start, doc.Paragraphs(3).Range.End)
# range1.HighlightColorIndex = win32com.client.constants.wdYellow  # 替换背景颜色为绿色
range1.HighlightColorIndex = WdColorIndex.wdYellow.value
range1.Select()
word.Selection.Font.Bold = True

doc.Close()
word.Quit()
