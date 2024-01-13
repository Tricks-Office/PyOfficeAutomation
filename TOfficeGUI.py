import tkinter as tk
from tkinter import ttk
from tkinter import font
import pandas as pd
from tkinter import filedialog
import TOffice as tof

# 마스터 데이터 업데이트하기
def update_master (df, indexnum, var):
    df['Item'][indexnum] = var
    df.to_csv('Master/Master.csv', index = False)

# 폴더 경로 변경하기
def change_TPath (df, index):
    folder_selected = filedialog.askdirectory()
    folder_selected = folder_selected.replace('/','\\')
    txtPath[index].delete('1.0', tk.END)
    txtPath[index].insert(tk.END, chars=folder_selected)
    update_master (df, index, folder_selected)

# 파일 경로 변경하기 (index 0 : 양식 파일, 1 : 사전 파일)
def change_File (df, index):
    file_selected = filedialog.askopenfile()
    file_selected = file_selected.name.replace('/','\\')
    txtFile[index].delete('1.0', tk.END)
    txtFile[index].insert(tk.END, chars=file_selected)
    update_master (df, index + 2, file_selected)    

# 마스터 데이터 가져오기
df = pd.read_csv('Master/Master.csv')

# GUI 창 생성하고 기본 정보 설정
win = tk.Tk()
win.geometry("680x580")
win.title('업무 자동화 트릭')

# 폴더 경로 설정하는 구역 frame1 설정하기 
frame1 = tk.LabelFrame(win, text = '폴더', padx = 5, pady=5, width = 100)
frame1.pack(padx = 10, pady=10)

# 폴더 경로의 각 구성 요소 생성하기
lbPath1 = tk.Label(frame1, text="작업대상",width = 10, padx = 5, pady=5)
txtPath = []
txtPath.append(tk.Text(frame1, width = 55, height = 1, padx = 5, pady=5, 
                        background='lightgrey'))
txtPath[0].insert(tk.INSERT, chars=df['Item'][0])
btnPath1 = tk.Button(frame1,text="경로 변경", width = 10,padx = 5, pady=5,
                    command = lambda : change_TPath(df,0))
lbPath2 = tk.Label(frame1, text="결과",width = 10, padx = 5, pady=5)
txtPath.append(tk.Text(frame1, width = 55, height = 1, padx = 5, pady=5, 
                        background='lightgrey'))
txtPath[1].insert(tk.INSERT, chars=df['Item'][1])
btnPath2 = tk.Button(frame1,text="경로 변경", width = 10,padx = 5, pady=5,
                    command = lambda : change_TPath(df,1))

# 폴더 경로의 각 구성 요소 배치하기
lbPath1.grid(row=0, column=0, padx = 5, pady=5)
txtPath[0].grid(row=0, column=1, padx = 5, pady=5)
btnPath1.grid(row=0, column=2, padx = 5, pady=5)
lbPath2.grid(row=1, column=0, padx = 5, pady=5)
txtPath[1].grid(row=1, column=1, padx = 5, pady=5)
btnPath2.grid(row=1, column=2, padx = 5, pady=5)




# 파일 경로 설정하는 구역 frame2 만들기
frame2 = tk.LabelFrame(win, text = '파일', padx = 5, pady=5, width = 100)
frame2.pack(padx = 10, pady=10)

# 파일 경로의 각 구성 요소 만들기
lbFile1 = tk.Label(frame2, text="양식",width = 10, padx = 5, pady=5)
txtFile = []
txtFile.append(tk.Text(frame2, width = 55, height = 1, padx = 5, pady=5, 
                        background='lightgrey'))
txtFile[0].insert(tk.INSERT, chars=df['Item'][2])
btnFile1 = tk.Button(frame2,text="경로 변경", width = 10,padx = 5, pady=5,
                    command = lambda : change_File(df,0))
lbFile2 = tk.Label(frame2, text="사전",width = 10, padx = 5, pady=5)
txtFile.append(tk.Text(frame2, width = 55, height = 1, padx = 5, pady=5, 
                        background='lightgrey'))
txtFile[1].insert(tk.INSERT, chars=df['Item'][3])
btnFile2 = tk.Button(frame2,text="경로 변경", width = 10, padx = 5, pady=5,
                    command = lambda : change_File(df,1))

# 파일 경로의 각 구성 요소 배치하기
lbFile1.grid(row=0, column=0, padx = 5, pady=5)
txtFile[0].grid(row=0, column=1, padx = 5, pady=5)
btnFile1.grid(row=0, column=2, padx = 5, pady=5)
lbFile2.grid(row=1, column=0, padx = 5, pady=5)
txtFile[1].grid(row=1, column=1, padx = 5, pady=5)
btnFile2.grid(row=1, column=2, padx = 5, pady=5)

# 글꼴 설정하는 구역 frame3 만들기
frame3 = tk.LabelFrame(win, text = '글꼴', padx = 5, pady=5)
frame3.pack(padx = 10, pady=10)

# 글꼴 구역의 각 구성 요소 생성하기
lbEFont = tk.Label(frame3, text = '영문 글꼴', width = 10, padx =5, pady =5)
cbEFont = ttk.Combobox(frame3, width = 25)
lbHFont = tk.Label(frame3, text = '한글 글꼴', width = 10, padx =5, pady =5)
cbHFont = ttk.Combobox(frame3, width = 25)

# 폰트 정보 콤보 박스에 리스트 추가하고 readonly 속성주기
fonts = list(font.families())
fonts.sort()
cbEFont['values'] = fonts
cbHFont['values'] = fonts
cbEFont['state'] = 'readonly'
cbHFont['state'] = 'readonly'

# 글꼴 초깃값 불러오기
cbEFont.current(fonts.index(df['Item'][4]))
cbHFont.current(fonts.index(df['Item'][5]))

# 글꼴 변경시 master data 저장하기
cbEFont.bind('<<ComboboxSelected>>', lambda _ : update_master(df, 4, cbEFont.get()))
cbHFont.bind('<<ComboboxSelected>>', lambda _ : update_master(df, 5, cbHFont.get()))

# 글꼴 구역의 각 구성 요소 배치하기
lbEFont.grid(row=0, column=0, padx =5, pady =5)
cbEFont.grid(row=0, column=1, padx =5, pady =5)
lbHFont.grid(row=0, column=2, padx =5, pady =5)
cbHFont.grid(row=0, column=3, padx =5, pady =5)

# 암호 설정하는 구역 frame4 만들기
frame4 = tk.Frame(win)
frame4.pack(padx = 10, pady=10)

# 암호 구역의 각 구성 요소 생성하기
lbpw = tk.Label(frame4, text = '암호', width = 10, padx =5, pady =5)
txtpw = tk.Text(frame4, width = 55, height = 1, padx = 5, pady=5)

# 암호 구역의 각 구성 요소 배치하기
lbpw.grid(row=0, column=0, padx =5, pady =5)
txtpw.grid(row=0, column=1, padx =5, pady =5)

# 암호 초기값 입력
txtpw.insert(tk.INSERT, chars=df['Item'][6])


# 버튼 설정하는 구역 frame5 만들기
frame5 = tk.Frame(win)
frame5.pack(padx = 10, pady=10)

# # 버튼 생성하기
# btnSaveImg = tk.Button(frame5, text = '문서 파일내 이미지 저장', width = 25, padx=5, pady=5)
# btnMergeXl = tk.Button(frame5, text = '같은 형식 엑셀 병합', width = 25, padx=5, pady=5)
# btnMergeXl2 = tk.Button(frame5, text = '유사 형식 엑셀 병합 (양식)', width = 25, padx=5, pady=5)
# btnPptFont = tk.Button(frame5, text = 'PPT 폰트 통일 (글꼴)', width = 25, padx=5, pady=5)
# btnPptDic = tk.Button(frame5, text = 'PPT 단어 일괄 변경 (사전)', width = 25, padx=5, pady=5)
# btnWordDic = tk.Button(frame5, text = '워드 단어 일괄 변경 (사전)', width = 25, padx=5, pady=5)
# btnDocPDF = tk.Button(frame5, text = '오피스 문서 PDF 전환', width = 25, padx=5, pady=5)
# btnImgPDF = tk.Button(frame5, text = '이미지 PDF 전환', width = 25, padx=5, pady=5)
# btnPDFpw = tk.Button(frame5, text = 'PDF 암호 설정 (암호)', width = 25, padx=5, pady=5)

# 버튼 생성하기
btnSaveImg = tk.Button(frame5, text = '문서 파일 이미지 저장', width = 25, padx=5, pady=5, 
                        command = lambda : tof.D_Img(df['Item'][0], df['Item'][1]))
btnMergeXl = tk.Button(frame5, text = '같은 형식 엑셀 병합', width = 25, padx=5, pady=5, 
                        command = lambda : tof.C_xl(df['Item'][0], df['Item'][1]))
btnMergeXl2 = tk.Button(frame5, text = '유사 형식 엑셀 병합', width = 25, padx=5, pady=5, 
                        command = lambda : tof.M_xl(df['Item'][0], df['Item'][1], df['Item'][2]))
btnPptFont = tk.Button(frame5, text = 'PPT 폰트 통일 (글꼴)', width = 25, padx=5, pady=5, 
                        command = lambda : tof.PptFont(df['Item'][0], df['Item'][1], df['Item'][4],
                        df['Item'][5]))
btnPptDic = tk.Button(frame5, text = 'PPT 단어 일괄 변경', width = 25, padx=5, pady=5, 
                        command = lambda : tof.PPT_Dic(df['Item'][0], df['Item'][1], 
                        df['Item'][3]))
btnWordDic = tk.Button(frame5, text = '워드 단어 일괄 변경', width = 25, padx=5, pady=5, 
                        command = lambda : tof.Doc_Dic(df['Item'][0], df['Item'][1],
                        df['Item'][3]))
btnDocPDF = tk.Button(frame5, text = '오피스 문서 PDF 전환', width = 25, padx=5, pady=5, 
                        command = lambda : tof.D_PDF(df['Item'][0], df['Item'][1]))
btnImgPDF = tk.Button(frame5, text = '이미지 PDF 전환', width = 25, padx=5, pady=5, 
                        command = lambda : tof.Img_PDF(df['Item'][0], df['Item'][1]))
btnPDFpw = tk.Button(frame5, text = 'PDF 암호 설정 (암호)', width = 25, padx=5, pady=5, 
                        command = lambda : tof.pw_PDF(df['Item'][0], df['Item'][1], 
                        txtpw.get('1.0','end-1c')))


# 버튼 배치하기
btnSaveImg.grid(row=0, column=0, padx =5, pady =5)
btnMergeXl.grid(row=0, column=1, padx =5, pady =5)
btnMergeXl2.grid(row=0, column=2, padx =5, pady =5)
btnPptFont.grid(row=1, column=0, padx =5, pady =5)
btnPptDic.grid(row=1, column=1, padx =5, pady =5)
btnWordDic.grid(row=1, column=2, padx =5, pady =5)
btnDocPDF.grid(row=2, column=0, padx =5, pady =5)
btnImgPDF.grid(row=2, column=1, padx =5, pady =5)
btnPDFpw.grid(row=2, column=2, padx =5, pady =5)

win.mainloop()