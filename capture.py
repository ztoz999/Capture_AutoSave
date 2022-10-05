import time
import keyboard
import os
import tkinter
import tkinter.ttk as ttk
from tkinter import font, filedialog, messagebox
from PIL import ImageGrab, Image, ImageTk

# 기본 시작 폴더
# var_initpath = os.path.join(os.path.dirname(__file__), "").replace("\\", "/")
# 기본 시작 폴더를 이전 사용한 경로로
var_initpath = os.path.dirname(".").replace("\\", "/")

def SelFolder():
    """ 캡쳐한 이미지를 저장할 폴더를 선택한다. """
    # 기본 저장 경로를 저장할 변수
    dirname = filedialog.askdirectory(title="저장 경로를 선택하세요", initialdir=var_initpath)
    if dirname == "":
        return None
    else:
        saveentry.delete(0, "end")
        saveentry.insert(0, dirname)

def Capture():
    """ 사용자가 F9 키를 누르면 스크린 샷 저장 """
    try:
        if os.path.exists(saveentry.get()):
            curr_time = time.strftime("_%Y%m%d-%H%M%S")
            
            # ImageGrab.grab() : 스크린샷을 가져와서 이미지 객채를 만든다.
            obj_img = ImageGrab.grab()
            imgpath = saveentry.get()
            imgname = "/image{0}.{1}".format(curr_time, cap_typecombo.get().lower())
            
            img = Image.new(mode="RGB", size=obj_img.size, color=(0,0,0))
            img.paste(im=obj_img, box=(0,0))
            img.save(imgpath + imgname)

            # 캡쳐 파일 리스트에 저장돤 이미지 이름 추가
            imgname = imgname.replace("/", "")
            if filelist.get(0) == "":
                filelist.insert(0, imgname)
            elif imgname != filelist.get("end"):
                filelist.insert("end", imgname)
            else:
                return None
        else: 
            messagebox.showerror("Error", message="폴더가 존재하지 않거나 잘못된 경로입니다.")

    except Exception as err:
        messagebox.showerror(title="에러", message="{}".format(err))
    
    
def Keycap():
    """ 키보드 이벤트 F9을 함수 Capture()와 연결 """
    
    if os.path.exists(saveentry.get()):
        messagebox.showinfo("Capture Start", message="F9 키를 누르면 스크린 샷이 저장됩니다.")
        capturebutton.config(state="disabled", text="F9키를 누르세요")
        cancelbutton.config(state="active")

        keyboard.add_hotkey("F9", lambda : Capture())
        windows.iconify()
    else:
        messagebox.showerror("Error", "올바른 저장 경로가 아닙니다.")
        
def Keycancel():
    """ 캡쳐룰 중지하고, 키보드 이벤트 F9을 제거 """
    capturebutton.config(state="active", text="스크린샷 만들기")
    cancelbutton.config(state="disabled")
    keyboard.remove_hotkey("F9")


# 이미지 합치기 탭 함수
def mg_SelFolder():
    """ 병합할 이미지들을 가져와서 파일 리스트에 보여준다. """    
    mg_imgpath = filedialog.askopenfilenames(title="이미지 파일들을 선택하세요", 
                                    filetypes=[("이미지 파일", ("*.png", "*.jpg", "*.bmp")), ("All Files", "*.*")], 
                                    initialdir=var_initpath)
    
    if mg_imgpath == "":
        return None
    else:
        for file in mg_imgpath:
            mg_filelist.insert("end", file)

def mg_SaveFolder():
    """ 병합한 이미지를 저장할 폴더를 선택한다. """
    path = filedialog.askdirectory(title="병합 이미지를 저장할 폴더를 선택하세요", initialdir=var_initpath)
    mg_savepath.delete(0, "end")
    mg_savepath.insert(0, path)
    
def mg_DelFile(mode:str):
    """ 선택한 항목 제거하기 """
    match mode:
        case "select":
            index = reversed(mg_filelist.curselection())
            for i in index:
                mg_filelist.delete(i)
        case "all":
                mg_filelist.delete(0, "end")
    
    img_showlabel.config(image="")

def mg_Merge():
    """ 새창으로 병합 과정을 보여준다. """
    # 새 창을 만든다
    toplevel = tkinter.Toplevel(windows)

    posx = windows.winfo_rootx() + round((windows.winfo_height() - 200) / 2)
    posy = windows.winfo_rooty() + round((windows.winfo_height() - 100) / 2)
    
    toplevel.geometry("300x100+{0}+{1}".format(posx, posy))
    
    toplevel.config(relief="groove", borderwidth=2, bg="#345")
    toplevel.overrideredirect(True)   
    toplevel.lift()
    
    # 파이썬 내장 이미지
    # image="::tk::icons::error"
    # image="::tk::icons::information"
    # image="::tk::icons::warning"
    # image="::tk::icons::question"
    tkinter.Label(toplevel, image="::tk::icons::information", bg="#345").pack(padx=5, pady=5)
    tkinter.Label(toplevel, text="병합중", foreground="white", bg="#345").pack(fill="both", expand=True)
    
    progressbar = ttk.Progressbar(toplevel, value=1.0, maximum=100)
    progressbar.config(mode="determinate")
    progressbar.pack(padx=5, pady=5)
    
    # 이미지 병합 함수 호출
    try:
        mg_MergeImage(progressbar)
    except Exception as err:
        messagebox.showerror(str(err))
    finally:
        toplevel.withdraw()

def mg_MergeImage(progressbar:ttk.Progressbar):
    """ 선택한 파일들을 병합한다. """
    try:
        if os.path.exists(mg_savepath.get()) == False:
            raise Exception("올바른 저장 폴더 경로가 아닙니다.") 
        if mg_filelist.get(0) == "":
            raise Exception("선택한 이미지가 없습니다.") 

        # 넓이 옵션
        var_imgwidth = mg_widthcomb.get()
        if var_imgwidth == "원본유지":
            var_imgwidth = -1
        else:
            var_imgwidth = int(var_imgwidth)

        # 이미지 간격 읽어와서 간격 offset 정하기
        offset_dict = dict(zip(var_imgspace, range(0, 61, 20)))
        yoffset = 0
        for x, y in offset_dict.items():
            
            if x == mg_spacecombo.get():
                yoffset = y

        # 이미지 확장자 타입 가져오기 
        mext = mg_typecombo.get().lower()

        # 병합 파일 목록의 이미지 경로를 가져와서 이미지 객체 리스트 만들기
        obj_images = [Image.open(x) for x in mg_filelist.get(0, "end")]
        
        # 개별 이미지의 사이즈 구하기
        image_size = []
        if var_imgwidth > -1:
            image_size = [(int(var_imgwidth), int(var_imgwidth * x.size[1] / x.size[0])) for x in obj_images]
        else:
            image_size = [(x.size[0], x.size[1]) for x in obj_images]
        
        # 넓이와 높이를 별도의 튜플로 분리
        mwidth, mheight = zip(*image_size)

        # 새 이미지 캔버스의 크기 구하기
        maxwidth, minwidth, totalheight = max(mwidth), min(mwidth), sum(mheight)

        if yoffset > 0:
            totalheight += (yoffset * (len(obj_images) - 1))

        # 병합할 이미지 켄버스 만들기
        mergedimag = Image.new(mode="RGB", size=(maxwidth, totalheight), color=(0, 0, 0))

        # 켄버스 이미지 더하기

        # 처음 더하는 이미지의 높이 위치좌표 조기화
        yspace = 0
        for index, img in enumerate(obj_images):
            # 이미지 크기 변경 옵션일 경우 이미지 리사이징
            if var_imgwidth > -1:
                img = img.resize(image_size[index], resample=Image.Resampling.LANCZOS)
            
            # 이미지 크기가 원본유지일 경우 작은 이미지를 넓이 기준 가운데로 위치 좌표 조정
            xspace = 0
            if img.size[0] < maxwidth:
                xspace = round((maxwidth - img.size[0]) / 2)
            
            # 이미지를 이어 붙이기
            mergedimag.paste(im=img, box=(xspace, yspace))
            
            # 이미지를 이어붙일 높이 좌표 조정
            yspace += mheight[index] + yoffset

            # 진행률을 프로그래스바에 적용
            percent = (index + 1) / len(obj_images) * 100    # 실제 percent 정보를 계산
            progressbar.config(value = percent)
            progressbar.update()

        # 병합된 이미지 저장
        
        # 저장할 이미지 이름 만들기
        mfilename = "merge" + time.strftime("_%Y%m%d-%H%M%S") + "." + mext
        # 저장할 이미지 경로 가져오기
        msavepath = os.path.join(mg_savepath.get(), mfilename)
        mergedimag.save(msavepath)
        
        messagebox.showinfo(title="작업 완료", message="파일이 병합되었습니다.")

    except Exception as err:
        messagebox.showerror(title="에러", message=str(err))
    finally:
        pass

def SelfileShow(event:tkinter.Event, tab:str):
    """ 목록에서 선택한 이미지의 섬네일을 보여준다. """
    
    selection = event.widget.curselection()
    
    if selection != ():
        try:
            index = selection[0]
            if tab == "capture":
                if index == 0:
                    return None
                filename = os.path.join(saveentry.get(), event.widget.get(index)).replace("\\", "/")
            else:
                filename = event.widget.get(index) 
            
            # 섬네일 만들기
            image = Image.open(filename)
            origin_width, origin_height = image.width, image.height
            
            image.thumbnail((300,300))
            
            thumbnail = Image.new(mode="RGB", size=image.size, color=(0, 0, 0))
            thumbnail.paste(im=image, box=(0,0))
            
            # 재사용할 섬네일 파일 만들기
            path = os.path.dirname(__file__).replace("\\", "/")

            thumbnail.save(path + "/thumbnail.jpg")
            
            # tkinter.photoimage는 jpg 인식 불가
            # PIL의 ImageTk.PhotoImage 는 인식 
            img = ImageTk.PhotoImage(file=path + "/thumbnail.jpg")
            
            # 미리보기 레이블에 보여주기
            img_showlabel.config(image=img)
            img_showinfo.config(text="SIZE = {} x {}".format(origin_width, origin_height))
            # 가비지 콜렉터 방지
            img_showlabel.image = img
        except Exception as err:
            print(str(err))

def exit():
    response = messagebox.askyesno("종료", message="종료하시겠습니까?")
    if response == True:
        windows.destroy()

##############################################################
# Windows 창 설정
##############################################################
windows = tkinter.Tk()
windows.title("Capture")
windows.geometry("840x450")
windows.resizable(False, False)

# 글꼴 지정 변수
fontname = font.Font(family="Malgun Gothic", size=10)
fontinfoname = font.Font(family="Malgun Gothic", size=12)

# 캡쳐 노트북 만들기
noteframe = ttk.Notebook(windows, width=500)
noteframe.pack(side="left", padx=5, pady=5)

###################################################################################
# 화면 스크린샷 만들기 Frame
###################################################################################
img_showframe = tkinter.LabelFrame(windows, text="이미지 미리보기", font=fontname)
img_showframe.pack(side="right", padx=5, pady=5, fill="both", expand=True)

img_showlabel = tkinter.Label(img_showframe)
img_showlabel.pack()

img_showinfo = tkinter.Label(img_showframe, font=fontinfoname)
img_showinfo.pack(side="bottom", padx=5, pady=5)

# 캡쳐 위젯 프레임 만들기
captureframe = tkinter.Frame(windows)
captureframe.pack(fill="x", expand=True, padx=2, pady=2, ipadx=2, ipady=2)

# 캡쳐 저장 경로 설정
cap_saveframe = tkinter.LabelFrame(captureframe, text="저장경로를 설정하세요", font=fontname)
cap_saveframe.pack(fill="x", expand=True)

# 기본 저장 경로를 보여줄 엔트리 박스
saveentry = tkinter.Entry(cap_saveframe, font=fontname)
saveentry.pack(side="left", fill="x", expand=True, padx=2, pady=2)

# 저장 폴더 선택 버튼 : 실행 함수 호출 연결
selectbutton = tkinter.Button(cap_saveframe, text="폴더 선택", font=fontname, command = SelFolder)

selectbutton.pack(side="right", padx=2, pady=2)

# 캡쳐 파일 목록 보여주기 프렘임 
cap_fileframe = tkinter.LabelFrame(captureframe, text="스크린샷 파일 목록", font=fontname)
cap_fileframe.pack(fill="both", expand=True, padx=5, pady=5)

# 파일 리스트박스와 연동할 스크롤 바
scrolly = tkinter.Scrollbar(cap_fileframe, orient="vertical")
scrolly.pack(side="right", fill="y")

# 캡쳐 파일 리스트박스
filelist = tkinter.Listbox(cap_fileframe, font=fontname)
filelist.insert(0,"스크린샷 파일들이 여기에 표시됩니다.")
filelist.bind("<<ListboxSelect>>", lambda event: SelfileShow(event, tab="capture"))
filelist.pack(side="left", fill="both", expand=True)

# 파일 목록 리스트박스와 스크롤바 연결
filelist.config(yscrollcommand=scrolly.set)
scrolly.config(command=filelist.yview)

# 스크린샷 저장 옵션
cap_optionframe = tkinter.LabelFrame(captureframe, text="스크린샷 이미지 저장 옵션", font=fontname)
cap_optionframe.pack(padx=5, pady=5, fill="x", expand=True)

# 스크린샷 저장 포맷
tkinter.Label(cap_optionframe, text="이미지 타입", font=fontname).pack(side="left", padx=2, pady=2)
var_captype = ["JPG", "PNG", "BMP"]
cap_typecombo = ttk.Combobox(cap_optionframe, values=var_captype, state="readonly", width=8, font=fontname)
cap_typecombo.current(0)
cap_typecombo.pack(side="left", padx=2, pady=2)

# 스크린샷 시작과 중지 버튼 : 실행 함수 호출 연결
cap_startframe = tkinter.Frame(captureframe, width=60)
cap_startframe.pack(fill="both", expand=True)

capturebutton = tkinter.Button(cap_startframe, text="스크린샷 만들기", state="normal", command=Keycap, font=fontname)
capturebutton.pack(side="left", fill="x", expand=True, padx=2, pady=2)

cancelbutton = tkinter.Button(cap_startframe, text="스크린샷 중지", state="disabled", command=Keycancel, font=fontname)
cancelbutton.pack(side="right", fill="x", expand=True, padx=2, pady=2)



###################################################################################
# 이미지 합치기 Frame
###################################################################################
mergeframe = tkinter.Frame(windows)
mergeframe.pack(fill="x", expand=True, padx=2, pady=2, ipadx=2, ipady=2)

# Image Folder 경로 설정
mg_folderframe = tkinter.Frame(mergeframe)
mg_folderframe.pack(fill="x", expand=True, anchor="n")

# 이미지를 가져올 폴더 산택 버튼 : 실행 함수 호출 연결
mg_selectbutton = tkinter.Button(mg_folderframe, text="이미지 파일 선택", command=mg_SelFolder, font=fontname)
mg_selectbutton.pack(side="left", padx=5, pady=2, fill="x", expand=True)

# 가져온 이미지 목록중 선택한 항목을 제거할 버튼 : 실행 함수 호출 연결
mg_delbutton = tkinter.Button(mg_folderframe, text="선택 항목 제거", font=fontname)
mg_delbutton.config(command=lambda: mg_DelFile("select"))
mg_delbutton.pack(side="left", padx=5, pady=2, fill="x", expand=True)

mg_delbutton = tkinter.Button(mg_folderframe, text="전체 항목 제거", font=fontname)
mg_delbutton.config(command=lambda: mg_DelFile("all"))
mg_delbutton.pack(side="right", padx=5, pady=2, fill="x", expand=True)

# 선택한 폴더의 이미지 목록 보여주기
mg_fileframe = tkinter.Frame(mergeframe)
mg_fileframe.pack(fill="both", expand=True, anchor="n")

# 병합할 파일 목록 리스트박스의 스크롤바 설정 
mg_scrolly = tkinter.Scrollbar(mg_fileframe, orient="vertical")
mg_scrollx = tkinter.Scrollbar(mg_fileframe, orient="horizontal")
mg_scrolly.pack(side="right", fill="y")
mg_scrollx.pack(side="bottom", fill="x")

# 병합할 파일 목록 리스트박스
mg_filelist = tkinter.Listbox(mg_fileframe, font=fontname, selectmode="extended")
mg_filelist.pack(side="left", fill="both", expand=True)
mg_filelist.bind("<<ListboxSelect>>", lambda event: SelfileShow(event, tab="merge"))

# 병합할 파일 목록 리스트박스와 스크롤바 연결
mg_filelist.config(xscrollcommand=mg_scrollx.set, yscrollcommand=mg_scrolly.set)
mg_scrollx.config(command=mg_filelist.xview)
mg_scrolly.config(command=mg_filelist.yview)

# 옵션 선택하기
mg_optionframe = tkinter.LabelFrame(mergeframe, text="옵션을 선택하세요", font=fontname)
mg_optionframe.pack(fill="x", expand=True, anchor="n")

# 이미지 포맷 옵션
tkinter.Label(mg_optionframe, text="이미지 타입", font=fontname).pack(side="left")
var_imgtype = ["JPG", "PNG", "BMP"]
mg_typecombo = ttk.Combobox(mg_optionframe, values=var_imgtype, state="readonly", width=8, font=fontname)
mg_typecombo.current(0)
mg_typecombo.pack(side="left", padx=2, pady=2)

# 이미지 사이즈 옵션
tkinter.Label(mg_optionframe, text="이미지 크기", font=fontname).pack(side="left")
var_imgwidth = ["원본유지", "1920", "1280", "1024", "800"]
mg_widthcomb = ttk.Combobox(mg_optionframe, values=var_imgwidth, state="readonly", width=8, font=fontname)
mg_widthcomb.pack(side="left", padx=2, pady=2)
mg_widthcomb.current(0)

# 이미지 간격 옵션
var_imgspace = ["없음", "좁게", "보통", "넓게"]
mg_spacecombo = ttk.Combobox(mg_optionframe, values=var_imgspace, state="readonly", width=8, font=fontname)
mg_spacecombo.pack(side="right", padx=2, pady=2)
mg_spacecombo.current(0)
tkinter.Label(mg_optionframe, text="이미지 간격", font=fontname).pack(side="right")

# 저장할 폴더를 선택하는 프레임
mg_saveframe = tkinter.Frame(mergeframe)
mg_saveframe.pack(fill="x", ipadx=2, ipady=2, padx=2, pady=2)

tkinter.Label(mg_saveframe, text="저장 폴더", font=fontname).pack(side="left", padx=2, pady=2)
mg_savepath = tkinter.Entry(mg_saveframe)
mg_savepath.pack(side="left", padx=2, pady=2, fill="x", expand=True) 

# 저장 폴더 선택 버튼
mg_savebutton = tkinter.Button(mg_saveframe, text="폴더 선택", command=mg_SaveFolder, font=fontname)
mg_savebutton.pack(side="right", padx=2, pady=2)

# 병합 버튼 : 실행 함수 연결
mg_button = tkinter.Button(mergeframe, text="이미지 합치기", command=mg_Merge, font=fontname)
mg_button.pack(fill="both", expand=True, padx=2, pady=2)
# 버튼 위젯의 슬립 방지
mg_button.update()

# 노트북 탭 추가
noteframe.add(captureframe, text=">>화면 스크린샷 만들기")
noteframe.add(mergeframe, text=">>이미지 합치기")

# windows 를 종료하기 전 확인
windows.protocol("WM_DELETE_WINDOW", exit)

# main widows 루프
windows.mainloop()



if __name__ == "__main__":
    pass