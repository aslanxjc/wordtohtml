#-*-coding:utf-8 -*-
from win32com import client as wc
import os
import time
import random
import re
import win32com
from win32com.client import Dispatch, constants
import base64
from bs4 import BeautifulSoup
import json
import sys
import traceback
import pythoncom
reload(sys)
sys.setdefaultencoding('utf-8')

def rmdir(top):
    for root, dirs, files in os.walk(top, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))

def wordsToHtml(dir):
    #批量把文件夹的word文档转换成html文件
    #金山WPS调用，抢先版的用KWPS，正式版WPS
    #word = wc.Dispatch('KWPS.Application')
    print 2222222222222
    word = wc.Dispatch('Word.Application')
    print 333333333333
    
    print word
    for path, subdirs, files in os.walk(dir):
        for wordFile in files:
            wordFullName = os.path.join(path, wordFile)
            print "word:" + wordFullName
            doc = word.Documents.Open(wordFullName)
            wordFile2 = unicode(wordFile, "gbk")
            dotIndex = wordFile2.rfind(".")
            if(dotIndex == -1):
                print '********************ERROR: 未取得后缀名！'
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

def imgToBase64(imgfile):
    '''
    '''
    with open(imgfile, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())
    return encoded_string


def touchGenHtml(file="D:\\test.html"): 
    '''
    '''
    html_str = ""
    content = open(file).read()
    soup = BeautifulSoup(content,fromEncoding="gb18030")
    print dir(soup)
    #将所有图片转成base64显示
    for _img in soup.find_all("img"):
        print _img
        src = _img.attrs.get("src")
        if 'file:///' in src:
            src = src[8:]
        src = "D:\\WorkSpace\\wordtohtml\\wordtohtml\\upload\\out\\"+src
        #print src,888888888888888888888888888
        base64_str = imgToBase64(src)
        base64_str = "data:image/png;base64," + base64_str
        _img['src'] = base64_str
        
    #删除目录
    rmdir("D:\WorkSpace\wordtohtml\wordtohtml\upload\out")
    
    html_str = soup.prettify().replace("<?if !vml?>","<!--[if !vml]-->").replace("<?endif?>","<!--[endif]-->")
        
    return html_str
    

    
def touchHmtl(file="D:\\test.html"):
    '''
    '''
    content = open(file).read()
    soup = BeautifulSoup(content,fromEncoding="gb18030")
    #将所有图片转成base64显示
    for _img in soup.find_all("img"):
        print _img
        src = _img.attrs.get("src")
        if 'file:///' in src:
            src = src[8:]
        src = "D:\\WorkSpace\\wordtohtml\\wordtohtml\\upload\\out\\"+src

        base64_str = imgToBase64(src)
        base64_str = "data:image/png;base64," + base64_str
        _img['src'] = base64_str
    #获取H3标签后续的兄弟节点
    h3_next = soup.h3.next_siblings
    
    html_str = ""
    for _hn in h3_next:
        hn_str = str(_hn)
        html_str += hn_str
    #
    tg_jx_da = []
    h3s = soup.find_all("h3")

    for i,_h3 in enumerate(h3s):
    
        
        # if i==0:
            # html_str = html_str.split(str(_h3))[1]
        # if i>0:
            # _tbody = html_str.split(str(_h3))[0]
            # del_str = _tbody + str(_h3)
            # tg_jx_da.append(_tbody)
            
        _tbody = html_str.split(str(_h3))[0]
        del_str = _tbody + str(_h3)
        tg_jx_da.append(_tbody)
            
        #取出已经取到的
        #html_str = html_str.replace(tigan_str+str(_h3),"")
        html_str = html_str.replace(del_str,"")
        
    #从取出的题中分别提取题干，解析和答案
    data_lst = []
    for _tjd in tg_jx_da:
        tjd_soup = BeautifulSoup(_tjd)
        tjd_soup_str = str(tjd_soup)
        
        jd_lst = tjd_soup.find_all("h6")
        #jd_lst = tjd_soup.find_all("class_=MsoNormal")
        
        tg_str = ""
        jx_str = ""
        da_str = ""
        test_str = ""
        tmp_dct = {}
        da_lst = []
        for i,jd in enumerate(jd_lst):
            test_str = str(jd)
            _tjd_soup_str = tjd_soup_str
            
            
            #题干
            if i == 0:
                tg_str = _tjd_soup_str.split(str(jd))[0]
                #tg_str = _tjd_soup_str
                #tjd_soup_str = tjd_soup_str.replace(tg_str+str(jd),'')
                jx_str = _tjd_soup_str.split(str(jd))[1]
            #解析和答案
            if i == 1:
                
                da_str = _tjd_soup_str.split(str(jd))[1]
                
                #_tjd_soup_str = _tjd_soup_str.replace(tg_str+str(jd),'')
                
                jx_str = jx_str.split(str(jd))[0]
                
                #提取答案选项
                da_soup = BeautifulSoup(da_str)
                da_soup = str(da_soup)
                da_soup = da_soup.replace("<o:p></o:p>","")
                
                da_soup = BeautifulSoup(da_soup)
                
                das = da_soup.find_all('span')
                #print type(da_soup.text)
                das = da_soup.text.strip().replace("<o:p></o:p>","").split("、".decode("gb2312").encode("utf-8"))
                #das = da_soup.text.strip().replace("<o:p></o:p>","").split(u"、")

                #da_lst = []
                for _d in das:
                    da_lst.append(_d.strip())
                    #da_lst.append(_d.text)
                    #da_lst.append(_d.contents[0])
                    #with open('a.txt','a+') as f:
                    #    f.write(str(_d))
                    #    f.close()

        tmp_dct["tg_str"] = tg_str.replace("<?if !vml?>","<!--[if !vml]-->").replace("<?endif?>","<!--[endif]-->")
        tmp_dct["jx_str"] = jx_str
        tmp_dct["da_str"] = da_lst
        tmp_dct["test_str"] = da_str
        
        data_lst.append(tmp_dct)
    #删除目录
    rmdir("D:\WorkSpace\wordtohtml\wordtohtml\upload\out")
    data_json = json.dumps(data_lst)
    return data_lst

    #return soup.div.stripped_strings
           
                
def docToHtml(filename,filenameout):
    '''
    '''
    #w = win32com.client.DispatchEx('Word.Application')
    try:
        pythoncom.CoInitialize()
        w = win32com.client.Dispatch('Word.Application')
        w.Visible = 0
        w.DisplayAlerts = 0
        doc = w.Documents.Open( FileName = filename )
        wc = win32com.client.constants
        w.ActiveDocument.WebOptions.RelyOnCSS = 0
        w.ActiveDocument.WebOptions.OptimizeForBrowser = 0
        w.ActiveDocument.WebOptions.BrowserLevel = 0 # constants.wdBrowserLevelV4
        w.ActiveDocument.WebOptions.OrganizeInFolder = 0
        w.ActiveDocument.WebOptions.UseLongFileNames = 1
        w.ActiveDocument.WebOptions.RelyOnVML = 0
        w.ActiveDocument.WebOptions.AllowPNG = 0
        w.ActiveDocument.SaveAs( FileName = filenameout, FileFormat = 8 )
        doc.Close()
    except Exception,e:
        traceback.print_exc()
    #w.Documents.Close(0)
    #w.quit()
    return filenameout
    
 
if __name__ == "__main__":
    dir = "D:\wordtohtml"
    filename = "D:\wordtohtml3\shiti.docx"
    filenameout = "D:\wordtohtml3\shiti.html"
    # docToHtml(dir)
    # docToHtml(filename,filenameout)
    strings = touchHmtl(filenameout)