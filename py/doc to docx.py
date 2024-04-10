import os
import sys
import pickle
import re
import codecs
import string
from win32com import client as wc


word=wc.Dispatch("Word.Application")

rawpath='E:\\codeproiect\\py1\\报告模板'
for root,dirs,files in os.walk(rawpath):
   for i in files:
       # 找出文件中以.doc结尾并且不以~$开头的文件(~S是为了排除临时文件的)
       if i.endswith('.doc') and not i.startswith( '.$'):
          print(i)
          doc = word.Documents.Open(root +'\\'+ i)
          rename = os.path.splitext(i)
          # 将文件另存为.docx
          doc.SaveAs(root+'\\'+'renamedocx\\'+rename[0] +'.docx',12) # 12表示docx格式
          doc.Close()  # time.sleep(1)

word.Quit()


#document1 = Document("C:\Users\YH\Desktop\检验检测机构.docx")
#document1.save('new.doc')
