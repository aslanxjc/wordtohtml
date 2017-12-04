#-*-coding:utf-8 -*-
import os,time
import json
from django.core.serializers.json import DjangoJSONEncoder
from django.views.decorators.csrf import csrf_exempt
from django.http import HttpResponse
from django.shortcuts import render
from word_to_html import docToHtml,touchHmtl,touchGenHtml
from .add_image_to_word import add_image_to_word
from upload_to_oss import MyOSS
import traceback

filename = "D:\WorkSpace\wordtohtml\wordtohtml\upload\shiti.docx"
filenameout = "D:\WorkSpace\wordtohtml\wordtohtml\upload\out\shiti.html"


def test(request):
    '''
    '''
    return render(request,'test.html',{})


def rmdir(top):
    for root, dirs, files in os.walk(top, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))

def handle_uploaded_file(f,filename):
    # path = "C:\WorkSpace\wordtohtml\upload"
    # if not os.path.exists(path):
        # os.mkdir(path)
    # name = f.name
    # ext = name.split('.')[1]
    # filename = os.path.join(path,"word."+ext)

    try:
        with open(filename, 'wb+') as destination:
            for chunk in f.chunks():
                destination.write(chunk)
		destination.close()
        return filename
    except Exception,e:
		
        print e,777777777777777777777
		#traceback.print_exc()
        return None

@csrf_exempt
def wordToHtml(request):
    '''
    '''
    filename = "D:\WorkSpace\wordtohtml\wordtohtml\upload\shiti.docx"
    filenameout = "D:\WorkSpace\wordtohtml\wordtohtml\upload\out\shiti.html"
    ret = {}
    if request.method == "GET":
        return render(request,"test.html",{})

    file = request.FILES.get('file',None)
    print 8888888888888888888888
    print request.FILES,11111
    
    filename = handle_uploaded_file(file,filename)
    if filename:
        #word to html
        filenameout = docToHtml(filename,filenameout)
        #touch html
        data_lst = touchHmtl(filenameout)
        
        ret['code'] = 0
        ret['msg'] = 'success'
        ret['data'] = data_lst
        
        #rmdir(path)
        # return HttpResponse(html,content_type="text/html")
        # return render(request,html_dco,{})
        response = HttpResponse(json.dumps(ret),content_type="application/json")
        response['Access-Control-Allow-Origin'] = '*'
        
        return response
    else:
        ret['code'] = 1
        ret['msg'] = 'fail'
        ret['data'] = []
        
        response = HttpResponse(json.dumps(ret),content_type="application/json")
        response['Access-Control-Allow-Origin'] = '*'
        
        return response
        
@csrf_exempt
def addImgsToWord(request):
    '''
    '''
    ret = {}
    if request.method == "OPTIONS":
        return HttpResponse({})
    
    imgurls = request.POST.get('imgurls')
    imgurls = json.loads(imgurls)
    
    paper_title = request.POST.get('paper_title','')
    
    print imgurls,7777777777
    tmp_doc = add_image_to_word(imgurls)
    print tmp_doc,88888888888888888888
    
    #上传至oss
    myoss = MyOSS()
    ossfile = str(time.time()).replace('.','')+'.doc'
    #ossfile = "wxss/150212251033.doc"
    
    url = myoss.upload_from_local(tmp_doc,ossfile)
    os.remove(tmp_doc)
    
    ret['code'] = 0
    ret['msg'] = 'success'
    ret['data'] = {'url':"http://file.say365.xin/"+url}
    
    response = HttpResponse(json.dumps(ret),content_type="application/json")
    response['Access-Control-Allow-Origin'] = '*'
        
    return response
    
    
@csrf_exempt
def uploadWord(request):
    '''
    导入word文件
    '''
    ret = {}
    filename = "D:\WorkSpace\wordtohtml\wordtohtml\upload\shiti.docx"
    filenameout = "D:\WorkSpace\wordtohtml\wordtohtml\upload\out\shiti.html"
    
    if request.method == "OPTIONS":
        return HttpResponse({})
    file = request.FILES.get('file',None)
    print file,44444444444444444
    filename = handle_uploaded_file(file,filename)
    
    print filename,5555555555555555555
    html_str = ""
    if filename:
        filenameout = docToHtml(filename,filenameout)
        html_str = touchGenHtml(filenameout)
        
    ret["code"]= 0
    ret["msg"]= "success"
    ret["data"] = html_str
        
    response = HttpResponse(json.dumps(ret),content_type="application/json")
    response['Access-Control-Allow-Origin'] = '*'
    
    return response
    
    
