import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import os
from datetime import datetime

BASE_URL = "https://www.yyds567.com"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

USERNAME = "qqaazz"
PASSWORD = "aa123456"

session = requests.Session()
session.headers.update(HEADERS)

# =========================
# 登录
# =========================
def login():

    print("正在登录...")

    login_url = BASE_URL + "/wp-admin/admin-ajax.php"

    data = {
        "action": "user_login",
        "username": USERNAME,
        "password": PASSWORD,
        "remember": "1"
    }

    r = session.post(login_url, data=data)

    print(r.text)

    return r.status_code == 200

# =========================
# 获取跳转链接
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

    except:
        return ""

# =========================
# 获取详情
# =========================
def get_game_data(game_id):

    detail_url = f"{BASE_URL}/{game_id}.html"

    print("处理:", detail_url)

    try:

        r = session.get(detail_url, timeout=20)

        r.encoding = r.apparent_encoding

        soup = BeautifulSoup(r.text, "html.parser")

        name_tag = soup.find("h1")
        name = name_tag.get_text(strip=True) if name_tag else ""

        img_tag = soup.select_one(".entry-content img") or soup.find("img")

        icon_link = ""

        if img_tag and img_tag.get("src"):

            icon_link = img_tag.get("src")

            if icon_link.startswith("/"):
                icon_link = BASE_URL + icon_link

        storage_baidu = ""
        storage_tianyi = ""
        storage_xunlei = ""

        for a in soup.find_all("a"):

            href = a.get("href", "")
            text_a = a.get_text(strip=True)

            if "百度" in text_a or "pan.baidu.com" in href:
                storage_baidu = get_final_link(href)

            elif "天翼" in text_a or "cloud.189.cn" in href:
                storage_tianyi = get_final_link(href)

            elif "迅雷" in text_a or "xunlei" in href:
                storage_xunlei = get_final_link(href)

        password = ""

        for li in soup.find_all("li"):

            label = li.find("p", class_="data-label")
            info = li.find("p", class_="info")

            if label and info:

                if "解压密码" in label.get_text(strip=True):

                    password = info.get_text(strip=True)

                    break

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

        print("失败:", e)

        return None

# =========================
# 保存进度
# =========================
def save_progress(index):

    with open("progress.txt", "w", encoding="utf-8") as f:
        f.write(str(index))

# =========================
# 读取进度
# =========================
def load_progress():

    if not os.path.exists("progress.txt"):
        return 0

    with open("progress.txt", "r", encoding="utf-8") as f:
        return int(f.read().strip())

# =========================
# 主程序
# =========================
def main():

    if not login():

        print("登录失败")

        return

    df = pd.read_excel(
        "下载列表.xlsx",
        usecols=[1],
        header=None
    )

    game_ids = df[1].dropna().astype(str).tolist()

    # 读取进度
    start = load_progress()

    end = start + 100

    print(f"今日处理范围: {start+1} ~ {end}")

    game_ids = game_ids[start:end]

    all_data = []

    for i, gid in enumerate(game_ids, start=start):

        gid = gid.strip()

        if not gid:
            continue

        result = get_game_data(gid)

        if result:
            all_data.append(result)

        save_progress(i + 1)

        time.sleep(2)

    # 日期文件名
    today = datetime.now().strftime("%m-%d")

    filename = f"游戏下载链接{today}.xlsx"

    df = pd.DataFrame(all_data)

    df.to_excel(
        filename,
        index=False,
        engine="openpyxl"
    )

    print("保存完成:", filename)

if __name__ == "__main__":
    main()
