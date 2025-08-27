import requests
import pandas as pd
import time
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt

from matplotlib.font_manager import FontProperties


import os
import platform

if platform.system() == "Windows":
    FONT_PATH = r"C:\Windows\Fonts\msyh.ttc"
else:
    FONT_PATH =   "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
    





# ---------------- 财联社 ----------------
def fetch_cls_top20():
    url = "https://api3.cls.cn/v1/hot_stock"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    }
    params = {
        "app": "cailianpress",
        "os": "android",
        "sv": "835",
        "sign": "e89e141e1391c13c7d2b99d8c142848c"
    }
    try:
        resp = requests.get(url, headers=headers, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        stocks = data.get("data", [])[:20]
        return [item["stock"]["name"] for item in stocks]
    except Exception as e:
        print("❌ 财联社请求失败:", e)
        return []


# ---------------- 东方财富 ----------------
def fetch_eastmoney_top20():
    url = "http://push2.eastmoney.com/api/qt/clist/get"
    params = {
        "pn": "1",   # 页数
        "pz": "20",  # 每页数量
        "po": "1",
        "np": "1",
        "ut": "bd1d9ddb04089700cf9c27f6f7426281",
        "fltt": "2",
        "invt": "2",
        "fid": "f62",   # 按人气排序
        "fs": "m:0+t:6,m:0+t:80",  # A股 主板+创业板
        "fields": "f12,f14,f2,f3,f62"
    }
    try:
        res = requests.get(url, params=params, timeout=10)
        res.raise_for_status()
        data = res.json()
        stocks = data.get("data", {}).get("diff", [])[:20]
        # 返回股票名称列表（后续保存 Excel 用）
        return [stock.get("f14", "--") for stock in stocks]
    except Exception as e:
        print("❌ 东方财富请求失败:", e)
        return []


# ---------------- 同花顺 ----------------
def fetch_ths_top20():
    url = "https://dq.10jqka.com.cn/fuyao/hot_list_data/out/hot_list/v1/stock"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
        "Referer": "https://eq.10jqka.com.cn/",
        "Origin": "https://eq.10jqka.com.cn",
        "Accept": "application/json, text/plain, */*"
    }
    params = {"stock_type": "a", "type": "hour", "list_type": "normal"}
    try:
        resp = requests.get(url, headers=headers, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        stocks = data.get("data", {}).get("stock_list", [])[:20]
        return [item["name"] for item in stocks]
    except Exception as e:
        print("❌ 同花顺请求失败:", e)
        return []

# ---------------- 保存 Excel ----------------
def save_to_excel(cls_names, eastmoney_names, ths_names):
    df = pd.DataFrame({
        "财联社": cls_names + [""] * (20 - len(cls_names)),
        "东方财富": eastmoney_names + [""] * (20 - len(eastmoney_names)),
        "同花顺": ths_names + [""] * (20 - len(ths_names))
    })
    file_name = f"HotStock_Top20_{time.strftime('%Y%m%d')}.xlsx"
    df.to_excel(file_name, index=False)
    print(f"✅ 已保存到 Excel：{file_name}")
    return file_name

# ---------------- 生成词云 ----------------
def generate_wordcloud(file_path):
    df = pd.read_excel(file_path)
    cls_names = df.iloc[:, 0].dropna().tolist()
    eastmoney_names = df.iloc[:, 1].dropna().tolist()
    ths_names = df.iloc[:, 2].dropna().tolist()

    print("📌 财联社前20:", cls_names)
    print("📌 东方财富前20:", eastmoney_names)
    print("📌 同花顺前20:", ths_names)

    def weighted_list(names):
        weighted = []
        for i, name in enumerate(names):
            weight = len(names) - i
            weighted.extend([name] * weight)
        return weighted

    all_stocks = weighted_list(cls_names) + weighted_list(eastmoney_names) + weighted_list(ths_names)
    counter = Counter(all_stocks)



    wc = WordCloud(
        width=800,
        height=400,
        background_color="white",
        colormap="tab10",
        font_path=FONT_PATH,
        
    ).generate_from_frequencies(counter)

    font_prop = FontProperties(fname=FONT_PATH)

    plt.figure(figsize=(12, 6))
    plt.imshow(wc, interpolation='bilinear')
    plt.axis("off")

    # 使用 FontProperties 指定字体
    plt.title("热门股票词云（按排名加权）", fontsize=16, fontproperties=font_prop)


    img_path = f"HotStock_WordCloud_{time.strftime('%Y%m%d')}.png"
    plt.savefig(img_path, dpi=900, bbox_inches='tight')
    plt.show()
    print(f"✅ 图片已保存：{img_path}")

# ---------------- 主程序 ----------------
if __name__ == "__main__":
    cls_names = fetch_cls_top20()
    eastmoney_names = fetch_eastmoney_top20()
    ths_names = fetch_ths_top20()

    excel_file = save_to_excel(cls_names, eastmoney_names, ths_names)
    generate_wordcloud(excel_file)







