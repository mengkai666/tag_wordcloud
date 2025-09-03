import requests
import pandas as pd
import time
from collections import Counter

from matplotlib.font_manager import FontProperties
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import platform

# ----------------- 字体路径 -----------------
if platform.system() == "Windows":
    FONT_PATH = r"C:\Windows\Fonts\msyh.ttc"
else:
    FONT_PATH = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"

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
    url_rank = "https://emappdata.eastmoney.com/stockrank/getAllCurrentList"
    headers = {
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/138.0.0.0 Mobile Safari/537.36",
        "Referer": "https://vipmoney.eastmoney.com/",
        "Origin": "https://vipmoney.eastmoney.com",
        "Content-Type": "application/json",
    }
    payload = {
        "rankType": "1",
        "pageSize": 100,
        "pageIndex": 1
    }
    try:
        resp = requests.post(url_rank, headers=headers, json=payload, timeout=10)
        resp.raise_for_status()
        data_rank_json = resp.json()
        if not data_rank_json or "data" not in data_rank_json:
            print("❌ 东方财富接口返回异常:", data_rank_json)
            return []
        data_rank = data_rank_json['data']
        top20_codes = [item['sc'] for item in data_rank[:20]]

        # 构造第二接口 secids
        secids = ",".join([("1."+c[2:] if c.startswith("SH") else "0."+c[2:]) for c in top20_codes])
        url_detail = "https://push2.eastmoney.com/api/qt/ulist.np/get"
        params = {
            "ut": "f057cbcbce2a86e2866ab8877db1d059",
            "fltt": 2,
            "invt": 2,
            "fields": "f14,f12",
            "secids": secids
        }
        resp_detail = requests.get(url_detail, params=params, timeout=10)
        resp_detail.raise_for_status()
        data_detail_json = resp_detail.json()
        if not data_detail_json or "data" not in data_detail_json or "diff" not in data_detail_json['data']:
            print("❌ 东方财富第二接口异常:", data_detail_json)
            return []

        data_detail = data_detail_json['data']['diff']
        # 返回股票名称列表
        return [stock['f14'] for stock in data_detail]
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
    file_name = f"HotStock_Top20.xlsx"
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

    img_path = f"HotStock_WordCloud.png"
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




