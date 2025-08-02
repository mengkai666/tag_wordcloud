import pandas as pd
import requests
from bs4 import BeautifulSoup
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
import time
import random

# 豆瓣电影 Top 20 爬取
def fetch_movies():
    headers = {"User-Agent": "Mozilla/5.0"}
    url = "https://movie.douban.com/top250?start=0&filter="
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    movies = soup.find_all('div', class_='hd')[:20]
    data = [{"title": m.a.span.text, "category": "豆瓣TOP"} for m in movies]
    df = pd.DataFrame(data)
    df.to_excel("movies.xlsx", index=False)
    return df

# 生成词云并写入Excel
def generate_wordcloud():
    df = pd.read_excel("movies.xlsx")
    categories = df["category"].astype(str)
    text = " ".join(categories)
    wc = WordCloud(font_path="simhei.ttf", background_color="white", width=800, height=400).generate(text)
    wc.to_file("wordcloud.png")
    wb = load_workbook("movies.xlsx")
    ws = wb.create_sheet("词云图")
    img = Image("wordcloud.png")
    ws.add_image(img, "B2")
    wb.save("movies.xlsx")

if __name__ == "__main__":
    fetch_movies()
    time.sleep(random.uniform(1, 2))  # 延时防止封IP
    generate_wordcloud()
