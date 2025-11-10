import os, re
from bs4 import BeautifulSoup
from collections import Counter
import pandas as pd
import matplotlib.pyplot as plt
from wordcloud import WordCloud

# === 设置 ===
input_folder = "."
output_file = "商品分析结果.xlsx"

html_files = [f for f in os.listdir(input_folder) if f.endswith(".html")]
print(f"检测到 {len(html_files)} 个 HTML 文件：", html_files)

data = []
all_keywords = []
all_title_words = []

# === 解析 ===
for file in html_files:
    print(f"正在处理：{file}")
    try:
        os.makedirs(file.split('.')[0])
    except:
        pass
    with open(os.path.join(input_folder, file), "r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    # 每个商品项通过 data-sku 来识别
    items = soup.find_all("div", attrs={"data-sku": True})
    print(f"本页商品数量: {len(items)}")

    for item in items:
        # 商品标题
        title_tag = item.select_one("span._text_1g56m_31")
        title = title_tag.get_text(strip=True) if title_tag else ""

        # 商品关键词（在 _common-wrap_9uih3_1 ... 中）
        keyword_spans = item.select("div._common-wrap_9uih3_1 span")
        keywords = [s.get_text(strip=True) for s in keyword_spans if s.get_text(strip=True)]
        all_keywords.extend(keywords)

        # 价格
        price_tag = item.select_one("span._price_uqsva_14")
        price_text = ""
        if price_tag:
            parts = price_tag.find_all(text=True)
            price_text = "".join(p for p in parts if p.strip())
            price_match = re.search(r"\d+(\.\d+)?", price_text)
            price = float(price_match.group()) if price_match else None
        else:
            price = None

        # 销量
        sales_tag = item.select_one("div._goods_volume_container_1xkku_1")
        sales = sales_tag.get_text(strip=True) if sales_tag else ""

        # 商家
        shop_tag = item.select_one("a._name_d19t5_35 span")
        shop = shop_tag.get_text(strip=True) if shop_tag else ""

        # 累计标题词（标题拆词）
        title_words = re.findall(r"[\u4e00-\u9fa5]+|[a-zA-Z]+", title)
        all_title_words.extend(title_words)

        data.append({
            "商品名称": title,
            "价格": price,
            "商家": shop,
            "销量": sales,
            "关键词列表": ", ".join(keywords)
        })

    # === 汇总 ===
    df = pd.DataFrame(data)
    df.drop_duplicates(subset=["商品名称"], inplace=True)
    df.reset_index(drop=True, inplace=True)

    # 高频关键词统计
    keyword_counts = Counter(all_keywords)
    title_word_counts = Counter(all_title_words)
    common_keywords = keyword_counts.most_common(30)
    common_title_words = title_word_counts.most_common(30)

    # === 保存 Excel ===
    with pd.ExcelWriter(file.split('.')[0] + "\\" + output_file) as writer:
        df.to_excel(writer, index=False, sheet_name="商品数据")
        pd.DataFrame(common_keywords, columns=["关键词", "出现次数"]).to_excel(writer, index=False, sheet_name="高频关键词")
        pd.DataFrame(common_title_words, columns=["标题词", "出现次数"]).to_excel(writer, index=False, sheet_name="标题词频")

    print(f"✅ 数据分析完成，结果保存至 {file.split('.')[0]}\\{output_file}")

    # === 可视化 ===
    plt.rcParams["font.sans-serif"] = ["SimHei"]
    plt.rcParams["axes.unicode_minus"] = False

    # 1️⃣ 关键词词频柱状图
    if common_keywords:
        words, counts = zip(*common_keywords)
        plt.figure(figsize=(10,5))
        plt.bar(words, counts, color="cornflowerblue")
        plt.title(f"{file.split('.')[0]}\\商品关键词词频统计")
        plt.xticks(rotation=60)
        plt.tight_layout()
        plt.savefig(f"{file.split('.')[0]}\\关键词词频统计.png", dpi=200)
        plt.close()

    # 2️⃣ 价格分布
    valid_prices = df["价格"].dropna()
    if not valid_prices.empty:
        plt.figure(figsize=(8,4))
        plt.hist(valid_prices, bins=20, color="lightgreen", edgecolor="black")
        plt.title("价格分布")
        plt.xlabel("价格（元）")
        plt.ylabel("商品数量")
        plt.tight_layout()
        plt.savefig(f"{file.split('.')[0]}\\价格分布.png", dpi=200)
        plt.close()

    # 3️⃣ 词云
    print("🎨 正在生成词云图...")
    wordcloud = WordCloud(
        font_path="C:/Windows/Fonts/simhei.ttf",
        width=1000,
        height=600,
        background_color="white",
        max_words=200,
        colormap="viridis"
    ).generate_from_frequencies(dict(keyword_counts))

    wordcloud.to_file(f"{file.split('.')[0]}\\关键词词云.png")
    print(f"✅ 词云生成完成：{file.split('.')[0]}\\关键词词云.png")

    print(f"\n📊 输出图表：{file.split('.')[0]}\\词频统计.png、{file.split('.')[0]}\\价格分布.png、{file.split('.')[0]}\\关键词词云.png")
