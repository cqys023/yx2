
import requests
from bs4 import BeautifulSoup
import time
import pandas as pd
import os
import re
import json
from datetime import datetime
from urllib.parse import urljoin
from getpass import getpass

# =====================================================
# 基础配置
# =====================================================

BASE_URL = "https://www.yyds567.com"

CATEGORY_URL = BASE_URL + "/qbyx/"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
    "Referer": BASE_URL
}

# 每天采集数量
DAILY_LIMIT = 100

# 每条数据采集间隔
SLEEP_TIME = 2

# 文件目录
DATA_DIR = "data"

# 总表
MASTER_FILE = os.path.join(DATA_DIR, "游戏下载链接总表.xlsx")

# 进度文件
PROGRESS_FILE = "progress.json"

# =====================================================
# 登录账号
# 建议运行时输入，不要把密码写在代码里
# =====================================================

USERNAME = "qqaazz"
PASSWORD = "aa123456"

# =====================================================
# Session
# =====================================================

session = requests.Session()
session.headers.update(HEADERS)


# =====================================================
# 登录
# =====================================================

def login():

    print("\n正在登录账号...")

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
            timeout=30
        )

        print("登录返回：")
        print(r.text[:500])

        if r.status_code == 200:

            print("登录请求已返回，继续检查采集页面")

            return True

        print("登录失败，HTTP状态码：", r.status_code)

        return False

    except Exception as e:

        print("登录异常：", e)

        return False


# =====================================================
# 生成分页地址
# =====================================================

def get_page_url(page):

    if page <= 1:

        return CATEGORY_URL

    return f"{CATEGORY_URL}page/{page}"


# =====================================================
# 获取列表页游戏链接
# 根据用户提供的 HTML 结构提取
# =====================================================

def get_game_list(page):

    page_url = get_page_url(page)

    print("\n正在读取列表页：", page_url)

    try:

        r = session.get(
            page_url,
            timeout=30
        )

        if r.status_code != 200:

            print("列表页访问失败：", r.status_code)

            return None

        r.encoding = r.apparent_encoding

        soup = BeautifulSoup(r.text, "html.parser")

        # 用户提供的结构：
        # div.posts-wrapper article.post

        articles = soup.select(
            "div.posts-wrapper article.post"
        )

        game_list = []

        for article in articles:

            # 优先提取游戏标题链接
            a = article.select_one(
                "h2.entry-title a[href]"
            )

            # 如果没有标题链接，则使用封面链接
            if not a:

                a = article.select_one(
                    ".entry-media a[href]"
                )

            if not a:

                continue

            detail_url = urljoin(
                BASE_URL,
                a.get("href", "")
            )

            # 只接受数字 ID 的游戏详情页
            match = re.search(
                r"/(\d+)\.html(?:\?.*)?$",
                detail_url
            )

            if not match:

                continue

            game_id = match.group(1)

            # 去重当前页面链接
            if any(x["game_id"] == game_id for x in game_list):

                continue

            game_list.append({
                "game_id": game_id,
                "detail_url": detail_url
            })

        print("本页找到游戏数量：", len(game_list))

        return game_list

    except Exception as e:

        print("列表页读取失败：", e)

        return None


# =====================================================
# 获取最终跳转链接
# =====================================================

def get_final_link(url):

    try:

        if not url:

            return ""

        url = urljoin(BASE_URL, url)

        r = session.get(
            url,
            timeout=30,
            allow_redirects=True
        )

        return r.url

    except Exception as e:

        print("获取跳转失败：", e)

        return ""


# =====================================================
# 获取游戏详情
# =====================================================

def get_game_data(game_id, detail_url):

    print("\n-----------------------------------")
    print("正在采集游戏 ID：", game_id)
    print("详情地址：", detail_url)

    try:

        r = session.get(
            detail_url,
            timeout=30
        )

        if r.status_code != 200:

            print("详情页访问失败：", r.status_code)

            return None

        r.encoding = r.apparent_encoding

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        # 游戏名称
        name_tag = soup.find("h1")

        name = (
            name_tag.get_text(strip=True)
            if name_tag else ""
        )

        # 封面图片
        img_tag = (
            soup.select_one(".entry-content img")
            or soup.select_one(".entry-media img")
            or soup.find("img")
        )

        icon_link = ""

        if img_tag:

            icon_link = (
                img_tag.get("data-src")
                or img_tag.get("data-original")
                or img_tag.get("src")
                or ""
            )

            if icon_link:

                icon_link = urljoin(
                    BASE_URL,
                    icon_link
                )

        # 网盘链接
        storage_baidu = ""
        storage_tianyi = ""
        storage_xunlei = ""

        for a in soup.find_all("a", href=True):

            href = a.get("href", "").strip()

            text_a = a.get_text(
                strip=True
            )

            if not href:

                continue

            # 百度网盘
            if (
                "百度" in text_a
                or "pan.baidu.com" in href
            ):

                storage_baidu = get_final_link(href)

            # 天翼云盘
            elif (
                "天翼" in text_a
                or "cloud.189.cn" in href
            ):

                storage_tianyi = get_final_link(href)

            # 迅雷云盘
            elif (
                "迅雷" in text_a
                or "xunlei" in href
            ):

                storage_xunlei = get_final_link(href)

        # 解压密码
        password = ""

        for li in soup.find_all("li"):

            label = li.find(
                "p",
                class_="data-label"
            )

            info = li.find(
                "p",
                class_="info"
            )

            if label and info:

                label_text = label.get_text(
                    strip=True
                )

                if "解压密码" in label_text:

                    password = info.get_text(
                        strip=True
                    )

                    break

        if not name:

            print("没有获取到游戏名称，跳过")

            return None

        result = {

            "game_id": game_id,

            "name": name,

            "icon_link": icon_link,

            "storage_baidu": storage_baidu,

            "storage_tianyi": storage_tianyi,

            "storage_xunlei": storage_xunlei,

            "password": password,

            "detail_url": detail_url,

            "采集时间": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        }

        print("游戏名称：", name)
        print("百度网盘：", storage_baidu)
        print("天翼云盘：", storage_tianyi)
        print("迅雷云盘：", storage_xunlei)
        print("解压密码：", password)

        return result

    except Exception as e:

        print("详情采集失败：", e)

        return None


# =====================================================
# 读取 Excel
# =====================================================

def load_excel(filename):

    if not os.path.exists(filename):

        return []

    try:

        df = pd.read_excel(filename)

        return df.to_dict("records")

    except Exception as e:

        print("读取 Excel 失败：", e)

        return []


# =====================================================
# 保存 Excel
# =====================================================

def save_excel(data, filename):

    if not data:

        return

    os.makedirs(
        DATA_DIR,
        exist_ok=True
    )

    df = pd.DataFrame(data)

    temp_file = filename + ".tmp.xlsx"

    df.to_excel(
        temp_file,
        index=False,
        engine="openpyxl"
    )

    os.replace(
        temp_file,
        filename
    )

    print("Excel 已保存：", filename)


# =====================================================
# 读取进度
# =====================================================

def load_progress():

    default = {

        "page": 1,

        "index": 0,

        "date": "",

        "daily_count": 0
    }

    if not os.path.exists(PROGRESS_FILE):

        return default

    try:

        with open(
            PROGRESS_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            data = json.load(f)

        for key in default:

            if key not in data:

                data[key] = default[key]

        return data

    except Exception:

        return default


# =====================================================
# 保存进度
# =====================================================

def save_progress(progress):

    temp_file = PROGRESS_FILE + ".tmp"

    with open(
        temp_file,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            progress,
            f,
            ensure_ascii=False,
            indent=4
        )

    os.replace(
        temp_file,
        PROGRESS_FILE
    )


# =====================================================
# 记录采集失败
# =====================================================

def save_failed(game_id, detail_url):

    filename = os.path.join(
        DATA_DIR,
        "采集失败记录.txt"
    )

    with open(
        filename,
        "a",
        encoding="utf-8"
    ) as f:

        f.write(
            f"{datetime.now()} | {game_id} | {detail_url}\n"
        )


# =====================================================
# 主程序
# =====================================================

def main():

    os.makedirs(
        DATA_DIR,
        exist_ok=True
    )

    # 登录
    if not login():

        print("登录失败，程序结束")

        return

    # 读取总表
    all_data = load_excel(
        MASTER_FILE
    )

    # 已采集 ID
    seen_ids = set()

    for item in all_data:

        gid = str(
            item.get("game_id", "")
        ).strip()

        if gid and gid != "nan":

            seen_ids.add(gid)

    print("\n总表已有资源：", len(seen_ids))

    # 读取进度
    progress = load_progress()

    today = datetime.now().strftime(
        "%Y-%m-%d"
    )

    # 今天的 Excel
    daily_file = os.path.join(
        DATA_DIR,
        f"游戏下载链接{today}.xlsx"
    )

    # 读取当天已经采集的数据
    daily_data = load_excel(
        daily_file
    )

    # 统计当天数量
    daily_ids = set()

    for item in daily_data:

        gid = str(
            item.get("game_id", "")
        ).strip()

        if gid and gid != "nan":

            daily_ids.add(gid)

    # 日期变化，重置每日计数
    if progress["date"] != today:

        progress["date"] = today

        progress["daily_count"] = len(
            daily_ids
        )

        print("\n新的一天，继续上次的分页位置")

    else:

        progress["daily_count"] = max(
            int(progress.get("daily_count", 0)),
            len(daily_ids)
        )

    page = int(progress["page"])

    index = int(progress["index"])

    daily_count = int(
        progress["daily_count"]
    )

    print("\n===================================")
    print("每日采集目标：", DAILY_LIMIT)
    print("今天已经采集：", daily_count)
    print("今天还需采集：", max(0, DAILY_LIMIT - daily_count))
    print("当前页码：", page)
    print("当前页位置：", index + 1)
    print("===================================")

    # =================================================
    # 开始分页采集
    # =================================================

    while daily_count < DAILY_LIMIT:

        print("\n当前采集页码：", page)

        game_list = get_game_list(page)

        # 页面请求失败，不推进进度，方便下次重试
        if game_list is None:

            print("列表页请求失败，停止本次任务")

            break

        # 当前页没有游戏
        if len(game_list) == 0:

            print("当前页没有找到游戏，采集结束")

            break

        # 当前页采集完毕，进入下一页
        if index >= len(game_list):

            page += 1

            index = 0

            progress["page"] = page

            progress["index"] = index

            save_progress(progress)

            continue

        # 当前游戏
        game = game_list[index]

        game_id = game["game_id"]

        detail_url = game["detail_url"]

        print(
            f"\n今日进度：{daily_count}/{DAILY_LIMIT}"
        )

        print(
            f"分页位置：第 {page} 页，第 {index + 1} 个"
        )

        # 已经采集过的游戏跳过
        if game_id in seen_ids:

            print("该游戏已经采集过，跳过：", game_id)

            index += 1

            progress["page"] = page

            progress["index"] = index

            save_progress(progress)

            continue

        # 采集详情
        result = get_game_data(
            game_id,
            detail_url
        )

        if result:

            # 加入总表
            all_data.append(result)

            # 加入当天数据
            daily_data.append(result)

            # 更新去重集合
            seen_ids.add(game_id)

            daily_ids.add(game_id)

            # 计数
            daily_count += 1

            progress["daily_count"] = daily_count

            # 实时保存总表
            save_excel(
                all_data,
                MASTER_FILE
            )

            # 实时保存当天 Excel
            save_excel(
                daily_data,
                daily_file
            )

            print(
                f"\n采集成功：{daily_count}/{DAILY_LIMIT}"
            )

        else:

            print("本条采集失败，记录后继续下一条")

            save_failed(
                game_id,
                detail_url
            )

        # 无论成功或失败，都推进当前页位置
        index += 1

        progress["page"] = page

        progress["index"] = index

        progress["daily_count"] = daily_count

        progress["date"] = today

        save_progress(progress)

        # 延时
        time.sleep(SLEEP_TIME)

    # =================================================
    # 结束
    # =================================================

    print("\n===================================")

    print("本次任务结束")

    print("今天成功采集：", daily_count)

    print("每日目标：", DAILY_LIMIT)

    print("当前页码：", progress["page"])

    print("当前页位置：", progress["index"] + 1)

    if daily_count >= DAILY_LIMIT:

        print("今天已经完成 100 个采集目标")

    else:

        print("未达到目标，可能是页面结束或访问失败")

    print("总表：", MASTER_FILE)

    print("今日文件：", daily_file)

    print("进度文件：", PROGRESS_FILE)

    print("===================================")


# =====================================================
# 启动
# =====================================================

if __name__ == "__main__":

    main()
