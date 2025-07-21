import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import xlwings as xw
from tkinter import scrolledtext
import sys
import matplotlib.pyplot as plt
from pathlib import Path

"""
_summary_
제작일 2025 07 21

버전 수정 사항
- job 파일에서 불러오는거면,,, 중심점 찾을 필요도 없고 cube size도 필요 없을거 같은데 
- 데이터 P1 부터 시작할 필요도 없고 job 파일 순서대로 엑셀에 넣으면 될듯?
- 먼 길 돌아온 느낌 
- rep 파일 하나라도 선택 필수
- 어떤 파일이든 입력하지 않으면 오류 발생하도록 

개선 필요
- 혹시나 10, 50, 100 파일의 포인트 순서가 다르다면? 
- 데이터셋 별 차이가 5개씩인데 1개 차이로 하는 게 맞긴함. 근데 데이터 셋 급증&연산 급증
    300개라면 약 150개 데이터셋 발생 
"""

def read_insert(path, speed, wb):
    try:
        df = pd.read_csv(path, sep = ",", encoding = "cp949", header=None)
        if df.shape[1] <=1:
            df = pd.read_csv(path, sep = "\t", encoding = "cp949", header=None)
    except Exception as e:
        print(f"reading error: {e}")
    sheet = wb.sheets[f'raw{speed}%']
    sheet.range((2, 2)).value = df.values.tolist()

# 파일 읽고 num 찾는 함수
def reading(path):
    try:
        df = pd.read_csv(path, sep = ",", encoding = "cp949", header=None)
        if df.shape[1] <=1:
            df = pd.read_csv(path, sep = "\t", encoding = "cp949", header=None)
    except Exception as e:
        print("file reading error")
        
    df = df.astype(float)
            
    return df # [num:]

# 중복 제거하기
def preprocessing1(df):
    diff = df.diff().abs()
    mask = (diff <= 30).all(axis=1)
    # 차이가 모두 1 이하면 뒷 행 제거 (mask가 True인 행 제거)
    df_filtered = df[~mask] #.reset_index(drop=True)
    deleted_df = df.loc[mask.index[mask].tolist()] # 삭제된 데이터
    print(" - - - - - - - - - - - - - - - - - - - - - - - -\n")
    print("중복으로 인해 제거된 데이터 ")
    if len(mask.index[mask].tolist()) == 0:
        print("없음")
    else:
        deleted_df.index += 1 # 인덱스 원래 0부터 시작하는데 보기 편하기 1 더함
        print(deleted_df.to_string(header=False)) # 열 이름 안보이게 header 없앰 
    same_list =[]
    same_list.extend(deleted_df.index)
    
    return df_filtered, same_list

# 누락 있을 때 그룹 째로 제거하는 함수 만들기...
def preprocessing2(df, point_list):
    df_index = pd.DataFrame(index=df.index)
    diff = (df - pd.Series(point_list[0])).abs()
    point1 = (diff<20).all(axis=1) # 포인트 1로 추정되는 행의 인덱스를 df_index에 저장
    df_index.loc[point1[point1].index, 'point'] = 1
    
    diff = (df - pd.Series(point_list[1])).abs()
    point2 = (diff<20).all(axis=1)
    df_index.loc[point2[point2].index, 'point'] = 2
    
    diff = (df - pd.Series(point_list[2])).abs()
    point3 = (diff<20).all(axis=1)
    df_index.loc[point3[point3].index, 'point'] = 3
    
    diff = (df - pd.Series(point_list[3])).abs()
    point4 = (diff<20).all(axis=1)
    df_index.loc[point4[point4].index, 'point'] = 4
    
    diff = (df - pd.Series(point_list[4])).abs()
    point5 = (diff<20).all(axis=1)
    df_index.loc[point5[point5].index, 'point'] = 5
    # 몇 번째 포인트인지 나타낸 df_index 만듦
    
    # 누락된 데이터 있으면 해당 그룹 제거
    answer_block = [1, 2, 3, 4, 5] 
    valid_blocks = []
    df_work = df_index.copy()
    #ans_list = df_index.head(5).values.reshape(-1).tolist()
    ans_list = answer_block
    #print(type(ans_list))
    while len(df_work) >= 5:
        found = False
        for i in range(len(df_work) - 4):
            window = df_work.iloc[i:i+5]
            if window['point'].tolist() == ans_list: # 일치하는지 확인 
                valid_blocks.append(window) # 일치하면 리스트에 저장
                df_work = df_work.drop(window.index) # 찾은걸 제거
                found = True
                break 
        if not found:
            break
        
    try:
        df_index_valid = pd.concat(valid_blocks) # 올바른 그룹만 저장됨
    except Exception as e:
        messagebox.showerror("오류", f"올바른 job 파일을 넣어주세요:\n{str(e)}")
        return
        
    removed_indices = df_index.index.difference(df_index_valid.index) # 비교해서 뭘 뺐는지 확인
    
    deleted_df = df.loc[removed_indices]
    print(" - - - - - - - - - - - - - - - - - - - - - - - -\n")
    print("\n누락으로 인해 제거된 그룹 데이터 ")
    # print(removed_indices.tolist())
    if len(removed_indices) == 0:
        print("없음")
    else:
        deleted_df.index += 1 # 보기쉽게 인덱스가 1로 시작하도록 함
        print(deleted_df.to_string(header=False)) 
    non_list = []
    non_list.extend(deleted_df.index) #전역변수 
    
    # 각 포인트별 인덱스 저장, 연속적이지 않은 그룹이 있으면 해당 인덱스 데이터 제거 
   
    # T/F 로 된 1열 데이터프레임 df1에서 True인 인덱스만 저장, 다른 데이터프레임 df2에 해당 인덱스 행에 1 넣기
    # df2에서 1~5가 반복되어야 하는데 그렇지 않은 인덱스 찾고 해당 그룹 제거
    # 제거하기 전 어떤 인덱스를 제거하는지 경고
    # 3열로 된 df3 에서 제거되지 않은 인덱스의 데이터를 추출 
    df_filtered = df.loc[df_index_valid.index]
    
    return df_filtered, non_list

# 데이터 받아서 같은 지점 데이터만 추출하는 전처리 함수
def point_extract(df, n): 
    row = df.iloc[n] # n행 기준으로 같은 지점만 필터링 
    diff = (df - row).abs()
    # 각 행별 차이가 15 이내인 행만 필터링 (모든 열이 조건을 만족해야 하면 all(axis=1))
    df = df[(diff <= 15).all(axis=1)].reset_index(drop=True)
    
    return(df[-30:]) # 굳이 슬라이싱 필요없긴 함 

# plot 추가하기
def chart_img(wb, speed):
    sheet = wb.sheets[f'POSE_{speed}%_mppi']
    data = []
    for i in range(46,106):
        data.append(sheet.range(f'Q{i}').value)
    df = pd.DataFrame(data)
    df_filtered = df[df <= 100].dropna()
    data = sum(df_filtered.values.tolist(),[])
    fig, axs = plt.subplots(1,2, figsize=(8,6))
    axs[0].set_xticks([0])
    axs[0].set_title("Box Plot")
    axs[0].set_xlabel("Rep")
    axs[0].boxplot(data, widths = 0.3)
    axs[0].scatter([1]*len(data), data, alpha = 0.7, s=20)
    axs[1].set_title("Histogram")
    axs[1].hist(data,bins=5,edgecolor='black')
    axs[1].set_xlabel("Rep")
    plt.tight_layout()
    plt.savefig('chart.png', dpi=300)
    plt.close()
    plot_file= Path('chart.png').resolve()
    sheet.pictures.add(str(plot_file), left=sheet.range('T48').left, top=sheet.range('T47').top)
    os.remove('chart.png')

# 전처리 결과 보여주기 raw 시트에 입력 
def result_show(wb,speed,same_list,non_list):
    raw_sheet = wb.sheets[f'raw{speed}%']
    raw_sheet.range('G2').value = "중복으로 제거된 데이터"
    for i in range(len(same_list)):
        raw_sheet.range(f'F{i+3}').value = same_list[i]
        raw_sheet.range(f'F{i+3}').font.color =(255,0,0)
        raw_sheet.range(f'G{i+3}').value = raw_sheet.range(f'B{same_list[i]+1}:D{same_list[i]+1}').value
    for row in same_list:
        cell = raw_sheet.range(f'B{row+1}:D{row+1}')
        cell.color = (255,0,0)
    
    raw_sheet.range('K2').value = "누락으로 제거된 데이터"
    for i in range(len(non_list)):
        raw_sheet.range(f'J{i+3}').value = non_list[i]
        raw_sheet.range(f'J{i+3}').font.color = (255,165,0)
        raw_sheet.range(f'K{i+3}').value = raw_sheet.range(f'B{non_list[i]+1}:D{non_list[i]+1}').value
    for row in non_list:
        cell = raw_sheet.range(f'B{row+1}:D{row+1}')
        cell.color = (255,165,0)

# 시트에 df 추가하기 특정 행, 열부터
def insert_data(wb,df,speed, start_row, start_col, i):
    sheet = wb.sheets[f'POSE_{speed}%_mppi']
    sheet.range((start_row, start_col)).value = df.values.tolist()
    # rep, acc 밑쪽에 들어가는지 확인하기 
    sheet.range((46+i, 2+3*(start_col-3)/7)).value = sheet.range((2,3+(start_col-3)/7),(3,3+(start_col-3)/7)).value
    return wb 

def processing_files(file_path1, file_path2, file_path3, output_path, point_li):
    print(f"현재 경로{os.getcwd()}")
    file_path = os.path.join(os.getcwd(), "ISO_form.xlsx")

    print(f"파일 경로: {file_path}")
    print(f"파일 존재 여부: {os.path.exists(file_path)}")
    print(f"파일 크기: {os.path.getsize(file_path)}")
    
    wb = xw.Book(file_path)
    speed_li = [10,50,100]
    # 각 시트 데이터 상단에 타겟 값 입력하기 
    for speed in speed_li:
        sheet1 =wb.sheets[f'POSE_{speed}%_mppi']
        sheet1.range((8,3)).value = point_li[0] 
        sheet1.range((8,10)).value = point_li[1]
        sheet1.range((8,17)).value = point_li[2]
        sheet1.range((8,24)).value = point_li[3]
        sheet1.range((8,31)).value = point_li[4]
    
    if(file_path1):    # rep10 데이터 넣기 
        read_insert(file_path1, 10, wb)
        speed = 10
        # p1~p5까지 수행하기 위한 for문 
        print("=====================================================================\n"*2)
        print(f"{file_path1} 처리")
        df = reading(file_path1)
        df, same_list = preprocessing1(df)
        df, non_list = preprocessing2(df, point_li)
        result_show(wb,speed,same_list,non_list) # 전처리 결과 입력하기
        
        i = 0
        while True:
            df1 = df[5*i:5*i+150]
            if len(df1)<150:
                break
            for j in range(5):
                df_p = point_extract(df1, j) # point j 30개 데이터 추출 
                insert_data(wb, df_p, speed, 11, 3+7*j, i) # 엑셀 파일에 넣기 
            i += 1
            
        # 최소 세트를 원래 데이터에 넣기
        min_set=wb.sheets[f'POSE_{speed}%_mppi'].range('V43').value
        df1 = df[int(5*(min_set-1)):int(5*(min_set-1)+150)]
        for j in range(5):
            df_p = point_extract(df1, j) # point j 30개 데이터 추출
            insert_data(wb, df_p, speed, 11, 3+7*j, min_set -1) # 엑셀 파일에 넣기
        try:
            chart_img(wb,10)
        except Exception as e:
            print(f"이미지 삽입 오류: {e}")
    
    if(file_path2): # rep50 데이터 넣기
        read_insert(file_path2, 50, wb)
        speed = 50
        print("=====================================================================\n"*2)
        print(f"{file_path2} 처리")
        df = reading(file_path2)
        df, same_list = preprocessing1(df)
        df, non_list = preprocessing2(df, point_li)
        result_show(wb,speed,same_list,non_list) # 전처리 결과 입력하기
        
        i = 0
        while True:
            df2 = df[5*i:5*i+150]
            if len(df2)<150:
                break
            for j in range(5):
                df_p = point_extract(df2, j)
                insert_data(wb, df_p, speed, 11, 3+7*j, i)
            i += 1
            
        min_set=wb.sheets[f'POSE_{speed}%_mppi'].range('V43').value
        df1 = df[int(5*(min_set-1)):int(5*(min_set-1)+150)]
        for j in range(5):
            df_p = point_extract(df1, j) # point j 30개 데이터 추출
            insert_data(wb, df_p, speed, 11, 3+7*j, min_set -1) # 엑셀 파일에 넣기
        try:
            chart_img(wb,50)
        except Exception as e:
            print(f"이미지 삽입 오류: {e}")
    
    if(file_path3): # rep100 데이터 넣기 
        read_insert(file_path3, 100, wb)
        speed = 100
        print("=====================================================================\n"*2)
        print(f"{file_path3} 처리")
        df = reading(file_path3)
        df, same_list = preprocessing1(df)
        df, non_list = preprocessing2(df, point_li)
        result_show(wb,speed,same_list,non_list) # 전처리 결과 입력하기
        
        i = 0
        while True:
            df3 = df[5*i:5*i+150]
            if len(df3)<150:
                break
            for j in range(5):
                df_p = point_extract(df3, j)
                insert_data(wb, df_p, speed, 11, 3+7*j, i)
            i += 1
        min_set=wb.sheets[f'POSE_{speed}%_mppi'].range('V43').value
        df1 = df[int(5*(min_set-1)):int(5*(min_set-1)+150)]
        for j in range(5):
            df_p = point_extract(df1, j) # point j 30개 데이터 추출
            insert_data(wb, df_p, speed, 11, 3+7*j, min_set -1) # 엑셀 파일에 넣기
        try:
            chart_img(wb,100)
        except Exception as e:
            print(f"이미지 삽입 오류: {e}")
        
    wb.save(output_path)
    
def run_process():
    file1 = entry_file1.get() if entry_file1.get() else False 
    file2 = entry_file2.get() if entry_file2.get() else False
    file3 = entry_file3.get() if entry_file3.get() else False
    
    if not any([file1, file2, file3]): # 전부 False 일 때
        messagebox.showerror("오류", "rep 파일을 선택해주세요.")
        return
    if not entry_job.get():
        messagebox.showerror("오류", "job 파일을 선택해주세요.")
        return
    if entry_save.get():
        filename = entry_save.get().strip() + '.xlsx'
    else:
        messagebox.showerror("오류", "파일명을 입력해주세요")
        return
    
    global point1,point2,point3,point4,point5 
    point_list = [point1,point2,point3,point4,point5]
    
    if os.path.basename(filename) != filename:
        messagebox.showerror("오류", "저장 파일명에는 경로를 포함하지 말고 입력하세요.")
        return

    try:
        processing_files(file1, file2, file3, filename, point_list)
        messagebox.showinfo("성공", f"파일이 성공적으로 저장되었습니다:\n{filename}")
    except Exception as e:
        messagebox.showerror("오류", f"파일 처리 중 오류 발생:\n{str(e)}")

# gui 위한 함수
def select_file1():
    path = filedialog.askopenfilenames(title="파일 선택")
    if path:
        for i in range(len(path)):
            if path[i].split('/')[-1].find('100')>0:
                entry_file3.delete(0, tk.END)
                entry_file3.insert(0, path[i])
            elif path[i].split('/')[-1].find('50')>0: 
                entry_file2.delete(0, tk.END)
                entry_file2.insert(0, path[i])
            elif path[i].split('/')[-1].find('10')>0: 
                entry_file1.delete(0, tk.END)
                entry_file1.insert(0, path[i])

def select_jobfile():
    path = filedialog.askopenfilename(title="job 파일 선택")
    if path: # 선택한 파일 열고 x, y, z 찾는 함수 구성하기 
        entry_job.delete(0, tk.END)
        entry_job.insert(0, path)
        # 엔트리에 이미 입력돼있으면 삭제하기
        df = pd.read_csv(path, sep = " ", encoding = "cp949", header=None)
        p1_list = df[6][1][1:].split(",")[:3]
        global point1 
        point1 = [float(p1_list[0].split(".")[0]),
                    float(p1_list[1].split(".")[0]),
                    float(p1_list[2].split(".")[0])]
        
        p2_list = df[6][3][1:].split(",")[:3]
        global point2
        point2 = [float(p2_list[0].split(".")[0]),
                    float(p2_list[1].split(".")[0]),
                    float(p2_list[2].split(".")[0])]
        
        p3_list = df[6][5][1:].split(",")[:3]
        global point3 
        point3 = [float(p3_list[0].split(".")[0]),
                    float(p3_list[1].split(".")[0]),
                    float(p3_list[2].split(".")[0])]
        
        p4_list = df[6][7][1:].split(",")[:3]
        global point4 
        point4 = [float(p4_list[0].split(".")[0]),
                    float(p4_list[1].split(".")[0]),
                    float(p4_list[2].split(".")[0])]
        
        p5_list = df[6][9][1:].split(",")[:3]
        global point5 
        point5 = [float(p5_list[0].split(".")[0]),
                    float(p5_list[1].split(".")[0]),
                    float(p5_list[2].split(".")[0])]
        
# 하단 코드 tkinter GUI 구성
root = tk.Tk()
root.title("JJH")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

tk.Label(root, text="※ 엑셀 서식 파일(form.xlsx)이 동일한 경로에 위치하여야 합니다.", 
         font=("Arial", 10, "bold")).pack(pady=(5, 0))

# 파일 3개 입력하기
tk.Label(frame, text="rep10 파일:").grid(row=4, column=0, sticky='e')
entry_file1 = tk.Entry(frame, width=50)
entry_file1.grid(row=4, column=1)

tk.Label(frame, text="rep50 파일:").grid(row=5, column=0, sticky='e')
entry_file2 = tk.Entry(frame, width=50)
entry_file2.grid(row=5, column=1)

tk.Label(frame, text="rep100 파일:").grid(row=6, column=0, sticky='e')
entry_file3 = tk.Entry(frame, width=50)
entry_file3.grid(row=6, column=1)

btn_file1 = tk.Button(frame, text="찾아보기", command=select_file1)
btn_file1.grid(row=5, column=2, padx=5)
frame.grid_rowconfigure(8, minsize=5) # 행 사이 간격

# job 파일 선택
tk.Label(frame, text="job 파일:").grid(row=9, column=0, sticky='e')
entry_job = tk.Entry(frame, width=50)
entry_job.grid(row=9, column=1)
btn_job = tk.Button(frame, text="찾아보기", command=select_jobfile)
btn_job.grid(row=9, column=2)

frame.grid_rowconfigure(11, minsize=5) # 행 사이 간격 

# 저장 파일 이름 설정
tk.Label(frame, text="저장 파일명:").grid(row=12, column=0, sticky='e')
entry_save = tk.Entry(frame, width=50)
entry_save.grid(row=12, column=1)

# 실행 버튼
btn_run = tk.Button(frame, text="실행", command=run_process, bg='lightblue')
btn_run.grid(row=13, column=1, padx=5)

# 창에서 로그 나타나도록
log_text = scrolledtext.ScrolledText(root, height=10, state='disabled', bg='black', fg='white')
log_text.pack(fill='both', expand=True)

# print 한 게 로그로 나오게 
def log_message(msg):
    log_text.config(state='normal')
    log_text.insert('end', msg + '\n')
    log_text.see('end')
    log_text.config(state='disabled')

# stdout/stderr 리디렉션 print가 로그로 출력되게?
class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        log_message(text.strip())

    def flush(self):
        pass

sys.stdout = TextRedirector(log_text)
sys.stderr = TextRedirector(log_text)

root.mainloop()
