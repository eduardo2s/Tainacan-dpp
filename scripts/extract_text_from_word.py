# -*- coding: utf-8 -*-
"""
Created on Fri Jan 25 22:21:08 2019

@author: eduardo
"""
import os
word_text = []

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n'.join(paragraphs)


for root, dirs, files in os.walk(os.path.abspath("C:\\Users\\eduar\\Desktop\\fichas ufg\\")):
    for file in files:
        word_text.append(get_docx_text(os.path.join(root, file)))
        
import csv

with open('C:\\Users\\eduar\\Desktop\\fichas ufg\\itens.csv', 'w', encoding="utf-8") as myfile:
     wr = csv.writer(myfile, quoting=csv.QUOTE_MINIMAL)
     for val in word_text:
        wr.writerow([val])