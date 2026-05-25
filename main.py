import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import os
from datetime import datetime

BASE_URL = "https://www.yyds567.com"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
    "Referer": BASE_URL
}

# =========================
# 账号密码
# =========================
USERNAME = "qqaazz"
PASSWORD = "aa123456"

# =========================
# 测试模式
# True  = 只下载1~5
# False = 正式模式每天100条
# =========================
TEST_MODE = True

# =========================
# Session
# =========================
session = requests.Session()
session.headers.update(HEADERS)

# =========================
# 登录
# =========================
def login():

    print("正在登录账号...")

    login_url = BASE_URL + "/wp-admin/admin-ajax.php"

    data = {
        "action": "user_login",
        "username": USERNAME,
        "password": PASSWORD,
        "remember": "1"
    }

    try:

        r = session.post(
            login_url,
            data=data,
            timeout=20
        )

        print("登录返回：")
        print(r.text)

        if r.status_code == 200:

            print("登录请求成功")

            return True

        return False

    except Exception as e:

        print("登录失败：", e)

        return False

# =========================
# 获取最终跳转链接
# =========================
def get_final_link(url):

    try:

        if not url:
            return ""

        if url.startswith("/"):
            url = BASE_URL + url

        r = session.get(
            url,
            timeout=20,
            allow_redirects=True
        )

        return r.url

    except Exception as e:

        print("获取跳转失败：", e)

        return ""

# =========================
# 获取游戏详情
# =========================
def get_game_data(game_id):

    detail_url = f"{BASE_URL}/{game_id}.html"

    print("\n正在处理:", detail_url)

    try:

        r = session.get(
            detail_url,
            timeout=20
        )

        if r.status_code != 200:

            print("访问失败：", r.status_code)

            return None

        r.encoding = r.apparent_encoding

        soup = BeautifulSoup(r.text, "html.parser")

        # =========================
        # 游戏名称
        # =========================
        name_tag = soup.find("h1")

        name = name_tag.get_text(strip=True) if name_tag else ""

        # =========================
        # 封面图
        # =========================
        img_tag = soup.select_one(".entry-content img") or soup.find("img")

        icon_link = ""

        if img_tag and img_tag.get("src"):

            icon_link = img_tag.get("src")

            if icon_link.startswith("/"):
                icon_link = BASE_URL + icon_link

        # =========================
        # 网盘链接
        # =========================
        storage_baidu = ""
        storage_tianyi = ""
        storage_xunlei = ""

        for a in soup.find_all("a"):

            href = a.get("href", "")
            text_a = a.get_text(strip=True)

            # 百度网盘
            if "百度" in text_a or "pan.baidu.com" in href:

                storage_baidu = get_final_link(href)

            # 天翼云盘
            elif "天翼" in text_a or "cloud.189.cn" in href:

                storage_tianyi = get_final_link(href)

            # 迅雷云盘
            elif "迅雷" in text_a or "xunlei" in href:

                storage_xunlei = get_final_link(href)

        # =========================
        # 解压密码
        # =========================
        password = ""

        for li in soup.find_all("li"):

            label = li.find("p", class_="data-label")
            info = li.find("p", class_="info")

            if label and info:

                if "解压密码" in label.get_text(strip=True):

                    password = info.get_text(strip=True)

                    break

        # =========================
        # 返回数据
        # =========================
        return {
            "name": name,
            "icon_link": icon_link,
            "storage_baidu": storage_baidu,
            "storage_tianyi": storage_tianyi,
            "storage_xunlei": storage_xunlei,
            "password": password,
            "detail_url": detail_url
        }

    except Exception as e:

        print("处理失败：", e)

        return None

# =========================
# 保存 Excel
# =========================
def save_excel(data):

    today = datetime.now().strftime("%m-%d")

    filename = f"游戏下载链接{today}.xlsx"

    df = pd.DataFrame(data)

    df.to_excel(
        filename,
        index=False,
        engine="openpyxl"
    )

    print("已保存：", filename)

# =========================
# 保存进度
# =========================
def save_progress(index):

    with open(
        "progress.txt",
        "w",
        encoding="utf-8"
    ) as f:

        f.write(str(index))

# =========================
# 读取进度
# =========================
def load_progress():

    if not os.path.exists("progress.txt"):

        return 0

    with open(
        "progress.txt",
        "r",
        encoding="utf-8"
    ) as f:

        return int(f.read().strip())

# =========================
# 主程序
# =========================
def main():

    # =========================
    # 登录
    # =========================
    if not login():

        print("登录失败，程序结束")

        return

    # =========================
    # 读取Excel
    # =========================
    df = pd.read_excel(
        "下载列表.xlsx",
        usecols=[1],
        header=None
    )

    game_ids = df[1].dropna().astype(str).tolist()

    print(f"\n总数量：{len(game_ids)}")

    # =========================
    # 测试模式 / 正式模式
    # =========================
    if TEST_MODE:

        start = 0
        end = 5

        print("\n当前为【测试模式】")
        print("只下载：1 ~ 5")

    else:

        start = load_progress()
        end = start + 100

        print("\n当前为【正式模式】")
        print(f"本次下载：{start+1} ~ {end}")

    # 截取范围
    game_ids = game_ids[start:end]

    # =========================
    # 开始处理
    # =========================
    all_data = []

    for i, gid in enumerate(game_ids, start=start):

        gid = gid.strip()

        if not gid:
            continue

        result = get_game_data(gid)

        if result:

            all_data.append(result)

            # 实时保存Excel
            save_excel(all_data)

        # 保存进度
        save_progress(i + 1)

        # 延时
        time.sleep(2)

    print("\n全部完成！")

# =========================
# 启动
# =========================
if __name__ == "__main__":

    main()

    input("\n按回车退出...")
