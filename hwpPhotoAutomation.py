from tkinter import *
from tkinter.filedialog import askopenfilenames,askopenfilename
import os
import win32com.client as win32

def image_event():
    btn1["state"]=DISABLED
    btn3["state"]=DISABLED

    global imagelist
    imagelist = askopenfilenames()

def file_event():
    btn2["state"]=DISABLED
    btn3["state"]=NORMAL

    global file_name
    file_name = askopenfilename()

def execute_event():
    btn1["state"]=NORMAL
    btn2["state"]=NORMAL
    
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 오브젝트 생성
    hwp.XHwpWindows.Item(0).Visible = True  # 숨김해제
    hwp.Open(file_name)

    for picture in imagelist:    
        hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
        temp=picture.rsplit("/",maxsplit=1)[1]
        temp=temp.rsplit(".",maxsplit=1)[0]
        temp=temp.replace('[꾸미기]', '')
        hwp.HParameterSet.HFindReplace.FindString = "NO."+temp
        print('No.'+temp)
        hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
        hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet);
        hwp.HAction.Run("MoveUp");
        hwp.InsertPicture(picture, Embedded=True, sizeoption=2) # 이미지 삽입

    print("실행")
    
root = Tk()
root.title("문서 사진 자동화 프로그램")
root.geometry("365x90+500+300")
root.resizable(False, False)

btn1 = Button(root, text='이미지 업로드', command=image_event)
btn1.grid(row=0, column=0, ipadx=10,ipady=10, padx=20, pady=20)

btn2 = Button(root, text='파일 선택', command=file_event)
btn2.grid(row=0, column=1, ipadx=10,ipady=10, padx=20, pady=20)

btn3 = Button(root, text='실행', command=execute_event)
btn3.grid(row=0, column=2, ipadx=10,ipady=10, padx=20, pady=20)

root.mainloop()

##
##for i in imagelist:
##    j=i.rsplit("/",maxsplit=1)[1]
##    print(j)
##    j=j.rsplit(".",maxsplit=1)[0]
##    print(j)
##    j=j.replace('[꾸미기]', '')
##    print(j)
##    
##
##root.destroy()
##

## BASE_DIR = imagelist[0].rsplit("/", maxsplit=1)[0]
## imagelist=[i.rsplit("/",maxsplit=1)[1] for i in imagelist]
