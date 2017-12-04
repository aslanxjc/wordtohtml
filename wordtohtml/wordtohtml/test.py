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
		
if __name__ == "__main__":
	f = open("a.txt","r")
	handle_uploaded_file(f,"D:\WorkSpace\wordtohtml\wordtohtml\upload\shiti.docx")