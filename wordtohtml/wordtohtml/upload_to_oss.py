#-*-coding:utf-8 -*-
import os
import oss2
import requests

class MyOSS:
    '''
    '''

    def __init__(self):
        '''
        '''
        #self.AccessKeyID = 'LTAIPp6YI1JUzpz2'
        #self.AccessKeySecret = 'UWc04c4TpUeVot9cYqbkRpU4YhE25S'
        #self.bucket_name = 'tederen'

        self.AccessKeyID = 'LTAIZL3kafPmPxaK'
        self.AccessKeySecret = 'tYr3nrBLeYfkjUI3r4MIp3JMrrjXH8'
        self.bucket_name = 'xzb365'

        self.root_name = 'wxss'

        self.auth = oss2.Auth(self.AccessKeyID,self.AccessKeySecret)
        #self.endpoint = 'http://oss-cn-shanghai.aliyuncs.com'
        self.endpoint = 'http://oss-cn-zhangjiakou.aliyuncs.com'

        self.bucket = oss2.Bucket(self.auth,self.endpoint,self.bucket_name)

        self.domain = 'http://tederen.oss-cn-shanghai.aliyuncs.com'

    def upload_from_str(self,content_str=None,filename=None):
        '''
        通过字符串上传
        byte,unicode,str
        '''
        filename = os.path.join(self.root_name,filename)

        result = self.bucket.put_object(filename,content_str)

        if result.status == 200:
            url = os.path.join(self.domain,filename)
        else:
            url = ''

        return url


    def upload_from_local(self,localfile=None,ossfile=None):
        '''
        上传本地文件到oss
        '''
        ossfile = os.path.join(self.root_name,ossfile).replace('\\','/')

        with open(localfile,'rb') as fileobj:
            result = self.bucket.put_object(ossfile,fileobj)

        if result.status == 200:
            url = os.path.join(self.domain,ossfile)
            url = ossfile
        else:
            url = ''

        return url

    def resumable_upload_from_local(self,localfile=None,ossfile=None):
        '''
        '''
        ossfile = os.path.join(self.root_name,ossfile)

        result = oss2.resumable_upload(self.bucket,ossfile,localfile,
            store=oss2.ResumableStore(root='/tmp'), 
            multipart_threshold=100*1024,
            part_size=100*1024,
            num_threads=4,
            progress_callback=self.progress_callback
        )

        if result.status == 200:
            url = os.path.join(self.domain,ossfile)
        else:
            url = ''

        return url


    def upload_from_url(self,url=None,ossfile=None):
        '''
        '''
        ossfile = os.path.join(self.root_name,ossfile)
        
        resp = requests.get(url)

        result = self.bucket.put_object(ossfile,resp)

        if result.status == 200:
            url = os.path.join(self.domain,ossfile)
        else:
            url = ''

        return url

    def progress_callback(self,consumed_bytes,total_bytes):
        '''
        '''
        if total_bytes:
            rate = int(100*float(consumed_bytes)/float(total_bytes))

            print '\r{0}%'.format(rate)


        



if __name__ == '__main__':
    myoss = MyOSS()
    #url = myoss.upload_from_str('test','test/test.txt')
    #print url

    #localfile = 'localtest.txt'
    #ossfile = 'localtooss/text.txt'
    #url = myoss.upload_from_local(localfile,ossfile)
    #print url

    #ossfile = 'test/test.jpg'
    #url = myoss.upload_from_url('http://img.ph.126.net/ocT0cPlMSiTs2BgbZ8bHFw==/631348372762626203.jpg',ossfile)
    #print url

    localfile = 'test.mp4'
    ossfile = 'localtooss/test.mp4'
    url = myoss.resumable_upload_from_local(localfile,ossfile)
    print url
