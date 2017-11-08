#-*-coding:utf-8 -*-
import win32com.client as win32

def add_image_to_word(images=[]):
    images = ["http://file.say365.xin/wxss/savepaperpieces/XZB2018RASXB0100100/1501839766.011-0_0.png",
                "http://file.say365.xin/wxss/savepaperpieces/XZB2018RASXB0100100/1501839765.692-0_0.png",
                "http://file.say365.xin/wxss/savepaperpieces/XZB2018RASXB0100100/1501839765.583-0_0.png"]
    try:
        word=win32.gencache.EnsureDispatch("Word.Application")
        print word,11111111111
        doc=word.Documents.Add()
        for img in images:
            word.Selection.InlineShapes.AddPicture(FileName=img,LinkToFile= False,SaveWithDocument=True)
        doc.SaveAs("c:\\tmp.doc")
        doc.Close(True)
        word.Application.Quit()
        return "c:\\tmp.doc"
    except Exception,e:
        print e,7777777777777777777
        return ""
        
add_image_to_word()