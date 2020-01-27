# -*- coding: utf-8 -*-
"""
Created on Sat Jan 25 16:10:27 2020
@author: Tao
"""

import docx


def read_docx(file_name):
    doc = docx.Document(file_name)
    content = '\n'.join([para.text for para in doc.paragraphs])
    return content


if __name__ == '__main__':
    print(read_docx(r'D:\pycharmspace\py4office\b.docx'))
