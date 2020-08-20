from __future__ import print_function
import xlrd
import xlwt
import time
import os
import os.path
from mailmerge import MailMerge
from datetime import date

def docReplace(paragraph, document):
    for paragraph in document.paragraphs:
    if 'sea' in paragraph.text:
        print paragraph.text
        paragraph.text = 'new text containing ocean'