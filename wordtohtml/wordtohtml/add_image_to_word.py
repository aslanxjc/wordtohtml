#-*-coding:utf-8 -*-
import win32com.client as win32


#-*-coding:utf-8 -*-
import win32com.client as win32

def add_image_to_word(images=[]):

    try:
        word=win32.gencache.EnsureDispatch("Word.Application")
        word.DisplayAlerts  = False
        print word,11111111111
        doc=word.Documents.Add()
        for img in images:
            word.Selection.InlineShapes.AddPicture(FileName=img,LinkToFile= False,SaveWithDocument=True)
        doc.SaveAs("c:\\tmp.doc")
        doc.Close(True)
        #doc.Quit()
        word.Application.Quit()
        return "c:\\tmp.doc"
    except Exception,e:
        print e,7777777777777777777
        return ""