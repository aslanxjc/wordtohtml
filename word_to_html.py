#-*-coding:utf-8 -*-
from win32com import client as wc
import os
import time
import random
import re

def wordsToHtml(dir):
    #�������ļ��е�word�ĵ�ת����html�ļ�
    #��ɽWPS���ã����Ȱ����KWPS����ʽ��WPS
    #word = wc.Dispatch('KWPS.Application')
    word = wc.Dispatch('Word.Application')
    print word
    for path, subdirs, files in os.walk(dir):
        print 111111111111111111111
        for wordFile in files:
            wordFullName = os.path.join(path, wordFile)
            print wordFullName,2222222222222222222
            print "word:" + wordFullName
            doc = word.Documents.Open(wordFullName)
            wordFile2 = unicode(wordFile, "gbk")
            dotIndex = wordFile2.rfind(".")
            if(dotIndex == -1):
                print '********************ERROR: δȡ�ú�׺����'
            fileSuffix = wordFile2[(dotIndex + 1) : ]
            if(fileSuffix == "doc" or fileSuffix == "docx"):
                fileName = wordFile2[ : dotIndex]
                htmlName = fileName + ".html"
                htmlFullName = os.path.join(unicode(path, "gbk"), htmlName)
                # htmlFullName = unicode(path, "gbk") + "\\" + htmlName
                print htmlFullName
                doc.SaveAs(htmlFullName, 8)
                doc.Close()
                word.Quit()
                print ""
                print "Finished!"
 
if __name__ == "__main__":
    dir = "C:\\test"
    print 222222222222222
    wordsToHtml(dir)