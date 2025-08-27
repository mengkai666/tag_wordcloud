import requests
import pandas as pd
import time
from collections import Counter
from wordcloud import WordCloud
import matplotlib.pyplot as plt

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
    import json
    secids = "1.600410,0.002261,0.002131,0.002600,1.600111,0.002241,1.688256,1.600118,1.600010,0.002402," \
             "0.002272,1.603019,0.002217,0.002555,0.002298,0.002456,1.600460,0.300059,1.601606,0.002681"
    url = f"https://push2.eastmoney.com/api/qt/ulist.np/get?fltt=2&np=3&ut=a79f54e3d4c8d44e494efb8f748db291&invt=2&secids={secids}&fields=f14"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        text = resp.text
        if text.startswith("qa_wap_jsonpCB"):
            text = text[text.find("(")+1:text.rfind(")")]
        data = json.loads(text)
        stocks = data.get("data", {}).get("diff", [])[:20]
        return [item["f14"] for item in stocks]
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
        font_path=r"C:\Windows\Fonts\msyh.ttc"
    ).generate_from_frequencies(counter)

    plt.figure(figsize=(12, 6))
    plt.imshow(wc, interpolation='bilinear')
    plt.axis("off")
    plt.title("热门股票词云（按排名加权）", fontsize=16, fontname="Microsoft YaHei")

    # 前20高频股票
    top20 = counter.most_common(20)
    for i, (name, freq) in enumerate(top20):
        plt.text(820, 20+i*20, f"{i+1}. {name} ({freq})",
                 fontsize=12, color=plt.cm.Reds(freq/top20[0][1]), fontname="Microsoft YaHei")

    plt.tight_layout()
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
