import pandas as pd
import tkinter as tk
from tkinter import ttk

# 엑셀 파일 경로 입력
EXCEL_FILE_PATH = "C:/CookAnalysis/Excel/RawData.xlsx"  

def show_selected_movies(country, genre):

    ## 정렬된 목록 보여주기
    def update_treeview():
        # 데이터 추가 (제목과 평점만 사용)
        for index, row in selected_movies.iterrows():
            tree.insert("", "end", values=(row["제목"], row["평점"], row["개봉연도"], row["OTT"]))

    ## Netflix 체크박스를 선택했을 때 실행
    def on_netflix_checked():
        pass

    ## Watcha 체크박스를 선택했을 때 실행
    def on_watcha_checked():
        pass

    ## Tiving 체크박스를 선택했을 때 실행
    def on_tiving_checked():
        pass

    ## 평점순 체크박스를 선택했을 때 실행
    def on_rating_checked():
        pass

    def on_release_checked():
        pass

    # 엑셀 파일 읽기
    df = pd.read_excel(EXCEL_FILE_PATH)

    # 선택된 국가와 장르에 해당하는 영화 데이터 필터링
    selected_movies = df[(df['국가'] == country) & (df['장르'] == genre)]

    # Tkinter 창 생성
    root = tk.Tk()
    root.geometry("600x330")
    root.title(f"{country} - {genre} 영화 목록")

    # 체크박스 변수 생성
    netflix_var = tk.BooleanVar(value=False)
    watcha_var = tk.BooleanVar(value=False)
    tiving_var = tk.BooleanVar(value=False)
    rating_var = tk.BooleanVar(value=False)
    release_var = tk.BooleanVar(value=False)
    # recommand_var = tk.BooleanVar(value=False)

    # 프레임 생성1 (OTT)
    checkbox_frame = tk.Frame(root)
    checkbox_frame.pack(side="top", pady=10)  # 체크박스 프레임을 위에 배치하고 간격을 추가
    tk.Checkbutton(checkbox_frame, text="Netflix", variable=netflix_var, command=on_netflix_checked).pack(side="left", padx=20)
    tk.Checkbutton(checkbox_frame, text="Watcha", variable=watcha_var, command=on_watcha_checked).pack(side="left", padx=20)
    tk.Checkbutton(checkbox_frame, text="Tiving", variable=tiving_var, command=on_tiving_checked).pack(side="left", padx=20)

    # 프레임 생성2 (정렬 기준)
    checkbox_frame2 = tk.Frame(root)
    checkbox_frame2.pack(side="top", pady=10)  # 체크박스 프레임을 위에 배치하고 간격을 추가
    tk.Checkbutton(checkbox_frame2, text="평점순", variable=rating_var, command=on_rating_checked ).pack(side="left", padx=20)
    tk.Checkbutton(checkbox_frame2, text="최근 개봉일순", variable=release_var, command=on_release_checked ).pack(side="left", padx=20)
   
    tk.Button(checkbox_frame2, text = "오늘의 영화 추천", fg = "black").pack(side="left", padx=20)

    # 표 생성
    tree = ttk.Treeview(root)

    # 열 제목 설정
    tree["columns"] = ("제목", "평점", "개봉연도", "OTT")
    tree.column("#0", width=0, stretch=tk.NO)  # 첫 번째 열(인덱스 열)은 보이지 않도록 설정

    # 나머지 열 설정
    for col in tree["columns"]:
        # 특정 열 크기 변경
        if col == "제목":
            tree.column(col, anchor="center", width=250)  # "제목" 열만 폭을 250으로 설정, 문자열 가운데 정렬
        else:
            tree.column(col, anchor="center", width=100)  # 나머지 열은 100으로 설정, 문자열 가운데 정렬
        tree.heading(col, text=col)

    update_treeview()  # 초기에 트리뷰 업데이트

    # 스크롤바 추가
    y_scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=y_scrollbar.set)
    y_scrollbar.pack(side="right", fill="y")

    # 표 출력
    tree.pack()

    # Tkinter 창 실행
    root.mainloop()

# 콤보박스에서 선택된 값 받아오기
def on_combobox_selected():
    selected_country = country_combo.get()
    selected_genre = genre_combo.get()
    show_selected_movies(selected_country, selected_genre)

# 엑셀 파일 읽기
df = pd.read_excel(EXCEL_FILE_PATH)

# 중복 없이 국가와 장르 목록 가져오기
countries = df['국가'].unique().tolist()
genres = df['장르'].unique().tolist()

# Tkinter 윈도우 생성
window = tk.Tk()
window.title("네이버 영화 크롤링 및 엑셀 파일 열기")

# 국가 선택 콤보박스
country_label = tk.Label(window, text="국가 선택:")
country_label.grid(row=0, column=0, padx=10, pady=5)
country_combo = ttk.Combobox(window, values=countries)
country_combo.grid(row=0, column=1, padx=10, pady=5)
country_combo.current(0)
country_combo.set("선택")

# 장르 선택 콤보박스
genre_label = tk.Label(window, text="장르 선택:")
genre_label.grid(row=1, column=0, padx=10, pady=5)
genre_combo = ttk.Combobox(window, values=genres)
genre_combo.grid(row=1, column=1, padx=10, pady=5)
genre_combo.current(0)
genre_combo.set("선택")

# 선택완료 버튼
submit_button = tk.Button(window, text="선택완료", command=on_combobox_selected)
submit_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

window.mainloop()