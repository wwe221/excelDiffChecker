import pandas as pd
import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog
import time

def search_google_scholar(query):
    base_url = "https://scholar.google.com/scholar"
    params = {"q": query}
    response = requests.get(base_url, params=params,  headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Whale/3.23.214.10 Safari/537.36"})
    refrenced_cnt = None    
    print(query ," :: ", response.status_code)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        result = soup.select_one('.gs_fl.gs_flb > a:nth-child(3)')        
        refrenced_cnt = result.text.split("회")[0]
    print(refrenced_cnt)
    return refrenced_cnt

cnt = 0
def process_excel_file(input_excel_file):
    # 엑셀 파일 읽기
    df = pd.read_excel(input_excel_file)

    df['길이'] = df.iloc[:, 0].apply(lambda x: search_google_scholar(x))
    print(cnt)
    if cnt % 10 == 0:
        time.sleep(5)
    cnt += 1
    # 수정된 데이터프레임을 새로운 엑셀 파일에 저장
    df.to_excel(input_excel_file+"2", index=False)

    print(f"Google Scholar search results have been written to {input_excel_file}")

def open_file_dialog():
    root = tk.Tk()
    root.withdraw()  # Tkinter 창 숨기기

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")],
    )

    return file_path
if __name__ == "__main__":
    # 입력 및 출력 파일 경로 지정
    input_excel_file_path = open_file_dialog()  # 입력 엑셀 파일 경로를 지정하세요

    # 엑셀 파일 처리
    process_excel_file(input_excel_file_path)
