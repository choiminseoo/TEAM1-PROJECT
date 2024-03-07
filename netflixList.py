import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from tkinter import messagebox
from tkinter.simpledialog import *
import random

def crawl_movies(ott,country, genre,value):
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    #창 안뛰우는 코드
    #chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(options=chrome_options)
    url = "https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&ssc=tab.nx.all&query=%EB%84%B7%ED%94%8C%EB%A6%AD%EC%8A%A4+%EC%B6%94%EC%B2%9C+%EC%98%81%ED%99%94&oquery=%EB%84%B7%ED%94%8C%EB%A6%AD%EC%8A%A4+%EC%98%81%EA%B5%AD+%EC%BD%94%EB%AF%B8%EB%94%94+%EC%98%81%ED%99%94&tqi=iOJUPdqps8wssK9MbOwssssstd0-061618"
    driver.get(url)
    time.sleep(0.02)

    # 국가 선택
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger2 > a > span.menu._text").click()
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger2 > div > div > div > div > div > ul > li:nth-child(" + str(country) + ") > a").click()
    

    # 장르 선택
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger3 > a > span.menu._text").click()
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger3 > div > div > div > div > div > ul > li:nth-child(" + str(genre) + ") > a").click()
    time.sleep(0.01)

    movies = []
    net_movies = []
    wat_movies = []
    tv_movies= []
   #넷플릭스 크롤링
    while True:
        tag = driver.find_element(By.CLASS_NAME, '_panel')
        all_movies = tag.find_elements(By.CLASS_NAME, 'info_box')
        for movie in all_movies:
              names = movie.find_element(By.CLASS_NAME, "title")
              MovieName = names.find_element(By.TAG_NAME, "a").text
              try:
                    rating = float(movie.find_element(By.CLASS_NAME, "num").text)
              except:
                   rating = 0.01
              if MovieName == '':
                   continue
              movies.append([MovieName, rating])
              net_movies.append([MovieName, rating])
        try:
             next = driver.find_element( By.XPATH, '//*[@id="main_pack"]/section[1]/div[2]/div/div/div[3]/div/a[2]')
             nowpage = driver.find_element(By.CLASS_NAME, 'npgs_now._current').text
             nextpage = driver.find_element(By.CLASS_NAME, '_total').text
             next.click()
             time.sleep(0.01)
             if int(nowpage) == int(nextpage):
                  break
        except:
             if len(movies)> 0 :
                  break
             elif ott ==2 and len(net_movies) < 1 :
                  messagebox.showerror("에러","영화가 존재하지 않음")
                  break
             else :
                  break
## 왓챠 크롤링
    watchaUrl = "https://search.naver.com/search.naver?sm=tab_hty.top&where=nexearch&ssc=tab.nx.all&query=%EC%99%93%EC%B1%A0+%EC%98%81%ED%99%94+%EC%B6%94%EC%B2%9C&oquery=%EC%BF%A0%ED%8C%A1%ED%94%8C%EB%A0%88%EC%9D%B4+%EC%98%81%ED%99%94+%EC%B6%94%EC%B2%9C&tqi=iPsFnlpzLi0ss4GTBVNssssssPo-283503"
    driver.get(watchaUrl)

    time.sleep(0.02)

    # 국가 선택
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger2 > a > span.menu._text").click()
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger2 > div > div > div > div > div > ul > li:nth-child(" + str(country) + ") > a").click()
    

    # 장르 선택
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger3 > a > span.menu._text").click()
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger3 > div > div > div > div > div > ul > li:nth-child(" + str(genre) + ") > a").click()
    time.sleep(0.01)

    while True:
        tag = driver.find_element(By.CLASS_NAME, '_panel')
        all_movies = tag.find_elements(By.CLASS_NAME, 'info_box')
        for movie in all_movies:   ## movies 리스트에 크롤링한 영화 넣는 반복문
              names = movie.find_element(By.CLASS_NAME, "title")
              MovieName = names.find_element(By.TAG_NAME, "a").text
              try:
                    rating = float(movie.find_element(By.CLASS_NAME, "num").text)
              except:
                   rating = 0.01
              if MovieName == '':
                   continue
              wat_movies.append([MovieName, rating])
              if [MovieName, rating] in movies :
                   continue
              movies.append([MovieName, rating])

        try:        ## 다음창 넘기는 부분
             next = driver.find_element( By.XPATH, '//*[@id="main_pack"]/section[1]/div[2]/div/div/div[3]/div/a[2]')
             nowpage = driver.find_element(By.CLASS_NAME, 'npgs_now._current').text
             nextpage = driver.find_element(By.CLASS_NAME, '_total').text
             next.click()
             time.sleep(0.01)
             if int(nowpage) == int(nextpage):
                  break
        except:
             if len(movies) > 0 :
                  break
             elif ott == 3 and len(wat_movies) < 1 :
                   messagebox.showerror("에러","해당 영화가 존재하지 않습니다.")
                   break
             else:
                  break

    ## 티빙 크롤링         
    tvingUrl = "https://search.naver.com/search.naver?where=nexearch&sm=top_hty&fbm=0&ie=utf8&query=%ED%8B%B0%EB%B9%99+%EC%98%81%ED%99%94+%EC%B6%94%EC%B2%9C"
    driver.get(tvingUrl)

    time.sleep(0.02)

    # 국가 선택
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger2 > a > span.menu._text").click()
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger2 > div > div > div > div > div > ul > li:nth-child(" + str(country) + ") > a").click()
    

    # 장르 선택
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger3 > a > span.menu._text").click()
    driver.find_element(By.CSS_SELECTOR, "#main_pack > section.sc_new.cs_common_module.case_list.color_5._cs_contents_recommendation > div.cm_content_wrap > div > div > div.cm_tap_area > div > div > ul > li.tab._select_trigger3 > div > div > div > div > div > ul > li:nth-child(" + str(genre) + ") > a").click()
    time.sleep(0.01)

    while True:
        tag = driver.find_element(By.CLASS_NAME, '_panel')
        all_movies = tag.find_elements(By.CLASS_NAME, 'info_box')
        for movie in all_movies:   ## movies 리스트에 크롤링한 영화 넣는 반복문
              names = movie.find_element(By.CLASS_NAME, "title")
              MovieName = names.find_element(By.TAG_NAME, "a").text
              try:
                    rating = float(movie.find_element(By.CLASS_NAME, "num").text)
              except:
                   rating = 0.01
              if MovieName == '':
                   continue
              tv_movies.append([MovieName, rating])
              if [MovieName, rating] in movies :
                   continue
              movies.append([MovieName, rating])

        try:        ## 다음창 넘기는 부분
             next = driver.find_element( By.XPATH, '//*[@id="main_pack"]/section[1]/div[2]/div/div/div[3]/div/a[2]')
             nowpage = driver.find_element(By.CLASS_NAME, 'npgs_now._current').text
             nextpage = driver.find_element(By.CLASS_NAME, '_total').text
             next.click()
             time.sleep(0.01)
             if int(nowpage) == int(nextpage):
                  break
        except:
             if ott == 1 and len(movies) < 1 :
                  messagebox.showerror("에러","해당 영화가 존재하지 않습니다.")
                  break
             elif len(movies)> 0 :
                  break
             elif ott == 4 and len(tv_movies) < 1 :
                  messagebox.showerror("에러","해당 영화가 존재하지 않습니다.")
                  break   
             

    if len(movies) > 2:
        movies.sort(key=lambda x: x[1], reverse=True)
    if len(net_movies) > 2:
        net_movies.sort(key=lambda x: x[1], reverse=True)
    if len(wat_movies) > 2:
        wat_movies.sort(key=lambda x: x[1], reverse=True)
    if len(tv_movies) > 2:
        tv_movies.sort(key=lambda x: x[1], reverse=True)
    
           
    
                             

    # tkinter 윈도우 생성
    if len(movies) > 0 :
         root = tk.Tk()
         root.title("영화 평점 순위")
         
         tree = ttk.Treeview(root)
         tree["columns"] = ("영화 제목", "평점")
         tree.column("#0", width=0, stretch=tk.NO)
         tree.column("영화 제목", anchor=tk.CENTER, width=300)
         tree.column("평점", anchor=tk.CENTER, width=100)
        
         tree.heading("#0", text="", anchor=tk.CENTER)
         tree.heading("영화 제목", text="영화 제목", anchor=tk.CENTER)
         tree.heading("평점", text="평점", anchor=tk.CENTER)

         # 추천받기 버튼
         submit_button = tk.Button(root, text="영화  추천받기", command=lambda: randchoice(movies))
         submit_button.pack(side="top")


    # 트리뷰에 데이터 삽입
         if ott == 1:
              for idx, (movie, rating) in enumerate(movies, start=1):
                   if not rating < value:
                        tree.insert("", "end", text=str(idx), values=(movie, rating))
         elif ott == 2:
              for idx, (movie, rating) in enumerate(net_movies, start=1):
                   if not rating < value:
                        tree.insert("", "end", text=str(idx), values=(movie, rating))
         elif ott == 3:
              for idx, (movie, rating) in enumerate(wat_movies, start=1):
                   if not rating < value:
                        tree.insert("", "end", text=str(idx), values=(movie, rating))
         else:
              for idx, (movie, rating) in enumerate(tv_movies, start=1):
                   if not rating < value:
                        tree.insert("", "end", text=str(idx), values=(movie, rating))

    # 스크롤바 생성
         vsb = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
         vsb.pack(side='right', fill='y')
         tree.configure(yscrollcommand=vsb.set)

    # 트리뷰 배치
         tree.pack(expand=True, fill='both')

         driver.quit()
         root.mainloop()
         print(ott)

def on_submit():
    # 콤보박스에서 선택한 인덱스를 가져옴
    selected_ott = ott_combo.current() + 1
    selected_country = country_combo.current() + 1  # 인덱스는 0부터 시작하므로 +1
    selected_genre = genre_combo.current() + 1      # 인덱스는 0부터 시작하므로 +1
    selected_value = value_combo.current() 
    crawl_movies(selected_ott,selected_country, selected_genre,selected_value)

# 특정 영화를 출력
def randchoice(movies):
     selected_movie = random.choice(movies)
     movie_name = selected_movie[0]
     chrome_options = Options()
     chrome_options.add_experimental_option("detach", True)
     url = f"https://search.daum.net/search?nil_suggest=btn&w=tot&DA=SBC&q=%EC%98%81%ED%99%94+{movie_name}"
     driver = webdriver.Chrome(options=chrome_options)
     driver.get(url)

ott = ["전체","넷플릭스","왓챠","티빙"]
country = ["전체", "한국", "미국", "일본", "중국", "대만", "영국", "해외"]
genre = ["전체", "공포", "스릴러", "로맨스", "코미디", "로맨틱코미디", "액션", "하이틴", "가족", "어린이", "SF", "판타지", "수사", "좀비", "재난", "전쟁", "에로",
         "BL", "추리", "범죄", "반전", "학교", "시대극", "우주", "시트콤", "음악", "무협", "슬픈", "감동적인", "힐링되는", "명작고전", "실화바탕", "웹툰원작", "마블", "디즈니", "연애"]
value = []
for i in range(0,11,1):
     value.append(i)

# Tkinter 윈도우 생성
window = tk.Tk()
window.title("ott 영화 크롤링 및  파일 열기")

# OTT 선택 콤보박스
ott_label = tk.Label(window, text="OTT 선택:")
ott_label.grid(row=0, column=0, padx=10, pady=5)
ott_combo = ttk.Combobox(window, values=ott)
ott_combo.grid(row=0, column=1, padx=10, pady=5)
ott_combo.current(0)

# 국가 선택 콤보박스
country_label = tk.Label(window, text="국가 선택:")
country_label.grid(row=1, column=0, padx=10, pady=5)
country_combo = ttk.Combobox(window, values=country)
country_combo.grid(row=1, column=1, padx=10, pady=5)
country_combo.current(0)

# 장르 선택 콤보박스
genre_label = tk.Label(window, text="장르 선택:")
genre_label.grid(row=2, column=0, padx=10, pady=5)
genre_combo = ttk.Combobox(window, values=genre)
genre_combo.grid(row=2, column=1, padx=10, pady=5)
genre_combo.current(0)

# 평점 선택 콤보박스
value_label = tk.Label(window, text="평점 필터:")
value_label.grid(row=3, column=0, padx=10, pady=5)
value_combo = ttk.Combobox(window, values=value)
value_combo.grid(row=3, column=1, padx=10, pady=5)
value_combo.current(0)

# 선택완료 버튼
submit_button = tk.Button(window, text="선택완료", command=on_submit)
submit_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

window.mainloop()
