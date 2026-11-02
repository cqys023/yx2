
import os
import re
import json
import time
import hashlib
import html
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime
from urllib.parse import urljoin

# =====================================================
# 基础配置
# =====================================================

BASE_URL = "https://www.yyds567.com"
CATEGORY_URL = BASE_URL + "/qbyx/"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
    "Referer": BASE_URL
}

DAILY_LIMIT = 100
SLEEP_TIME = 2

DATA_DIR = "data"
MASTER_FILE = os.path.join(DATA_DIR, "游戏下载链接总表.xlsx")
PROGRESS_FILE = "progress.json"
DETAIL_FAILED_FILE = os.path.join(DATA_DIR, "详情页上传失败记录.json")

# =====================================================
# yyds567 网站登录账号
# =====================================================

USERNAME = os.getenv("YYDS_USERNAME", "").strip()
PASSWORD = os.getenv("YYDS_PASSWORD", "").strip()

# =====================================================
# 宝塔 API 配置
# 由 GitHub Actions Secrets 提供
# =====================================================

GAME_API_URL = os.getenv("GAME_API_URL", "").strip()
GAME_API_TOKEN = os.getenv("GAME_API_TOKEN", "").strip()

# =====================================================
# Session
# =====================================================

session = requests.Session()
session.headers.update(HEADERS)


# =====================================================
# 登录
# =====================================================

def login():

    if not USERNAME or not PASSWORD:
        raise RuntimeError("缺少 YYDS_USERNAME 或 YYDS_PASSWORD Secrets")

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

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        articles = soup.select(
            "div.posts-wrapper article.post"
        )

        game_list = []
        page_seen = set()

        for article in articles:

            a = article.select_one(
                "h2.entry-title a[href]"
            )

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

            match = re.search(
                r"/(\d+)\.html(?:\?.*)?$",
                detail_url
            )

            if not match:
                continue

            game_id = match.group(1)

            if game_id in page_seen:
                continue

            page_seen.add(game_id)

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

            text_a = a.get_text(strip=True)

            if not href:
                continue

            if (
                "百度" in text_a
                or "pan.baidu.com" in href
            ):

                storage_baidu = get_final_link(href)

            elif (
                "天翼" in text_a
                or "cloud.189.cn" in href
            ):

                storage_tianyi = get_final_link(href)

            elif (
                "迅雷" in text_a
                or "xunlei" in href
            ):

                storage_xunlei = get_final_link(href)

        # 解压密码
        extract_password = ""

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

                label_text = label.get_text(strip=True)

                if "解压密码" in label_text:

                    extract_password = info.get_text(strip=True)
                    break

        if not name:

            print("没有获取到游戏名称，跳过")
            return None

        result = {
            "game_id": str(game_id),
            "name": name,
            "icon_link": icon_link,
            "storage_baidu": storage_baidu,
            "storage_tianyi": storage_tianyi,
            "storage_xunlei": storage_xunlei,
            "password": extract_password,
            "detail_url": detail_url,
            "采集时间": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        }

        print("游戏名称：", name)
        print("百度网盘：", storage_baidu)
        print("天翼云盘：", storage_tianyi)
        print("迅雷云盘：", storage_xunlei)
        print("解压密码：", extract_password)

        return result

    except Exception as e:

        print("详情采集失败：", e)
        return None


# =====================================================
# 采集详情页 HTML
# =====================================================

def build_detail_html(detail_url):

    print("\n正在采集详情页 HTML：", detail_url)

    response = session.get(
        detail_url,
        timeout=45
    )

    response.raise_for_status()
    response.encoding = response.apparent_encoding

    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )

    # 找到文章正文
    article = soup.find("article")

    if not article:
        raise RuntimeError("没有找到 article 标签")

    # 删除分享区域和脚本
    for node in article.select(
        ".entry-share, script, noscript"
    ):

        node.decompose()

    # 视频添加 16:9 容器
    for video in article.find_all("video"):

        parent = video.parent

        if (
            parent
            and parent.name == "div"
            and "video-container" in parent.get("class", [])
        ):

            continue

        wrapper = soup.new_tag(
            "div",
            attrs={"class": "video-container"}
        )

        video.wrap(wrapper)

    # 提取 article 的直接子元素
    content_elements = article.find_all(
        recursive=False
    )

    if not content_elements:
        raise RuntimeError("article 中没有正文内容")

    content = "".join(
        str(element)
        for element in content_elements
    )

    # HTML 模板
    template = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">

<title>游戏详情</title>

<base href="__BASE_URL__">

<link rel="stylesheet"
href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">

<style>

html {
    background-color: #f0f0f0;
}

body {
    font-family: sans-serif;
    line-height: 1.6;
    border: 1px solid #ccc;
    padding: 20px;
    background-color: white;
    margin: 20px auto;
    max-width: 800px;
    color: #222;
}

p {
    margin-bottom: 15px;
}

img {
    display: block;
    max-width: 100%;
    height: auto;
    margin: 10px auto;
}

.post-note {
    background-color: #f8d7da;
    border: 1px solid #f5c6cb;
    color: #721c24;
    padding: 10px;
    margin-top: 10px;
    border-radius: 5px;
}

.entry-share {
    display: none !important;
}

small {
    font-size: 0.8em;
}

strong {
    font-weight: bold;
}

.container {
    max-width: 800px;
}

.video-container {
    position: relative;
    padding-bottom: 56.25%;
    height: 0;
    overflow: hidden;
    margin-bottom: 15px;
    background: #000;
}

.video-container video {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
}

pre, table {
    max-width: 100%;
    overflow-x: auto;
}

</style>

</head>

<body>

<div class="container">

__CONTENT__

</div>

</body>
</html>"""

    final_html = template.replace(
        "__BASE_URL__",
        html.escape(detail_url, quote=True)
    ).replace(
        "__CONTENT__",
        content
    )

    return final_html


# =====================================================
# 详情页失败任务队列
# =====================================================

def load_detail_failed():

    if not os.path.exists(DETAIL_FAILED_FILE):
        return []

    try:

        with open(
            DETAIL_FAILED_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            data = json.load(f)

        return data if isinstance(data, list) else []

    except Exception as e:

        print("读取详情页失败队列异常：", e)
        return []


def save_detail_failed(tasks):

    os.makedirs(DATA_DIR, exist_ok=True)

    temp_file = DETAIL_FAILED_FILE + ".tmp"

    with open(
        temp_file,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            tasks,
            f,
            ensure_ascii=False,
            indent=2
        )

    os.replace(
        temp_file,
        DETAIL_FAILED_FILE
    )


def queue_detail_task(game_id, detail_url):

    tasks = load_detail_failed()

    for task in tasks:

        if str(task.get("game_id")) == str(game_id):

            task["detail_url"] = detail_url
            save_detail_failed(tasks)
            return

    tasks.append({
        "game_id": str(game_id),
        "detail_url": detail_url
    })

    save_detail_failed(tasks)


# =====================================================
# 检查宝塔 details 目录中是否已经存在详情 HTML
# =====================================================

def get_existing_detail_ids(game_ids):
    """向 PHP 查询服务器真实文件；只把存在且非空的 HTML 视为已完成。"""
    normalized_ids = []
    seen = set()
    for value in game_ids:
        game_id = str(value).strip()
        if not game_id or game_id.lower() == "nan":
            continue
        if game_id not in seen:
            seen.add(game_id)
            normalized_ids.append(game_id)

    existing = set()
    for start in range(0, len(normalized_ids), 500):
        batch = normalized_ids[start:start + 500]
        result = api_post(
            {
                "action": "check_details",
                "game_ids": batch
            },
            "检查宝塔详情 HTML 文件"
        )
        existing.update(
            str(x).strip()
            for x in result.get("existing_ids", [])
        )

    return existing


# =====================================================
# 上传单个详情页到宝塔
# =====================================================

def upload_detail_page(game_id, detail_url):

    if not GAME_API_URL or not GAME_API_TOKEN:
        raise RuntimeError("缺少 GAME_API_URL 或 GAME_API_TOKEN")

    print("\n正在下载游戏详情页：", game_id)

    page_html = build_detail_html(detail_url)

    print("详情 HTML 下载成功")
    print("HTML 大小：", round(len(page_html.encode("utf-8")) / 1024, 2), "KB")

    headers = {
        "Authorization": f"Bearer {GAME_API_TOKEN}",
        "Content-Type": "application/json; charset=utf-8"
    }

    payload = {
        "action": "upload_detail",
        "game_id": str(game_id),
        "html": page_html
    }

    # 不转义中文，减少上传数据体积
    request_data = json.dumps(
        payload,
        ensure_ascii=False
    ).encode("utf-8")

    response = requests.post(
        GAME_API_URL,
        data=request_data,
        headers=headers,
        timeout=180
    )

    print("详情页上传 HTTP 状态：", response.status_code)
    print("详情页上传返回：", response.text[:2000])

    response.raise_for_status()

    try:
        result = response.json()
    except ValueError as e:
        raise RuntimeError(
            "详情页接口没有返回有效 JSON：" + response.text[:1000]
        ) from e

    if not result.get("success"):
        raise RuntimeError(f"详情页上传失败：{result}")

    print("详情页已保存：", result.get("detail_html_url", ""))

    return result.get("detail_html_url", "")



# =====================================================
# 重试历史详情页上传失败任务
# =====================================================

def retry_failed_details():

    tasks = load_detail_failed()

    print("\n" + "=" * 50)
    print("详情页任务队列检查")
    print("当前待处理详情页任务数：", len(tasks))
    print("任务队列文件：", DETAIL_FAILED_FILE)
    print("=" * 50)

    if not tasks:
        print("详情页任务队列为空")
        return

    # 先检查宝塔服务器上的真实文件。文件已经存在且非空，就从重试队列移除。
    try:
        queued_ids = [str(task.get("game_id", "")).strip() for task in tasks]
        existing_ids = get_existing_detail_ids(queued_ids)
    except Exception as e:
        print("检查宝塔详情文件失败，本轮暂不处理详情任务：", e)
        return

    already_done = [
        task for task in tasks
        if str(task.get("game_id", "")).strip() in existing_ids
    ]
    tasks = [
        task for task in tasks
        if str(task.get("game_id", "")).strip() not in existing_ids
    ]

    if already_done:
        print("宝塔中已经存在详情 HTML，跳过采集和上传：", len(already_done), "条")
        save_detail_failed(tasks)

    if not tasks:
        print("所有详情页文件均已存在，无需采集或上传")
        return

    remaining = []

    for task in tasks:

        game_id = str(
            task.get("game_id", "")
        ).strip()

        detail_url = str(
            task.get("detail_url", "")
        ).strip()

        if not game_id or not detail_url:
            continue

        success = False

        for attempt in range(1, 4):

            try:

                print(
                    f"\n处理详情页 {game_id}，"
                    f"第 {attempt}/3 次"
                )

                upload_detail_page(
                    game_id,
                    detail_url
                )

                success = True
                break

            except Exception as e:

                print(
                    f"详情页 {game_id} 上传失败：{e}"
                )

                if attempt < 3:
                    time.sleep(3)

        if not success:
            remaining.append(task)

    save_detail_failed(remaining)

    print("\n详情页处理结束")
    print("上传成功数量：", len(tasks) - len(remaining))
    print("仍然失败数量：", len(remaining))


# =====================================================
# 根据当天 Excel 自动补充详情页任务
# =====================================================

def rebuild_daily_detail_tasks(daily_data):

    print("\n" + "=" * 50)
    print("开始检查当天 Excel 中的详情页任务")
    print("=" * 50)

    if not daily_data:
        print("当天没有游戏记录")
        return

    # 批量检查宝塔 details 目录，避免重复采集和上传已有详情页。
    daily_ids_to_check = [
        str(item.get("game_id", "")).strip()
        for item in daily_data
        if str(item.get("game_id", "")).strip()
        and str(item.get("game_id", "")).strip().lower() != "nan"
    ]

    try:
        existing_ids = get_existing_detail_ids(daily_ids_to_check)
    except Exception as e:
        print("检查宝塔详情文件失败，本轮不新增详情页任务：", e)
        return

    added = 0
    skipped_existing = 0

    for item in daily_data:

        game_id = str(
            item.get("game_id", "")
        ).strip()

        detail_url = str(
            item.get("detail_url", "")
        ).strip()

        if not game_id or game_id.lower() == "nan":
            continue

        if not detail_url or detail_url.lower() == "nan":
            print(f"游戏 {game_id} 缺少详情页地址")
            continue

        if game_id in existing_ids:
            skipped_existing += 1
            print(f"宝塔详情 HTML 已存在，跳过：{game_id}")
            continue

        queue_detail_task(
            game_id,
            detail_url
        )

        added += 1

    tasks = load_detail_failed()

    print("当天 Excel 游戏数量：", len(daily_data))
    print("宝塔中已存在、跳过的详情页：", skipped_existing)
    print("检查并补充任务数量：", added)
    print("当前待上传任务总数：", len(tasks))
    print("=" * 50)



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
            f"{datetime.now()} | "
            f"{game_id} | {detail_url}\n"
        )


# =====================================================
# 上传数据到宝塔 MySQL
# =====================================================



def clean_excel_row(row):
    """将 Excel 行清理为可 JSON 序列化的数据。"""
    item = {}
    for key, value in row.items():
        try:
            missing = pd.isna(value)
            if bool(missing):
                item[key] = ""
                continue
        except (TypeError, ValueError):
            pass

        if hasattr(value, "isoformat"):
            try:
                item[key] = value.isoformat(sep=" ")
            except TypeError:
                item[key] = value.isoformat()
        elif key == "game_id":
            item[key] = str(value).strip()
        elif isinstance(value, (str, int, float, bool)) or value is None:
            item[key] = value
        else:
            item[key] = str(value)
    return item


def game_data_hash(item):
    """必须与 PHP database_game_hash() 的字段顺序和分隔符完全一致。"""
    fields = [
        "name", "icon_link", "storage_baidu", "storage_tianyi",
        "storage_xunlei", "password", "detail_url"
    ]
    values = []
    for field in fields:
        value = item.get(field, "")
        values.append("" if value is None else str(value))
    return hashlib.sha256("\x1f".join(values).encode("utf-8")).hexdigest()


def api_post(payload, action_name, attempts=3):
    """调用宝塔 API，统一处理重试、HTTP 错误和 JSON 错误。"""
    if not GAME_API_URL or not GAME_API_TOKEN:
        raise RuntimeError("缺少 GAME_API_URL 或 GAME_API_TOKEN")

    headers = {
        "Authorization": f"Bearer {GAME_API_TOKEN}",
        "Content-Type": "application/json; charset=utf-8"
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            response = requests.post(
                GAME_API_URL, data=body, headers=headers,
                timeout=(20, 180)
            )
            if response.status_code != 200:
                raise RuntimeError(
                    f"HTTP {response.status_code}: {response.text[:1200]}"
                )
            try:
                result = response.json()
            except ValueError as exc:
                raise RuntimeError(
                    f"接口未返回有效 JSON：{response.text[:1000]}"
                ) from exc
            if not isinstance(result, dict) or not result.get("success"):
                raise RuntimeError(f"接口返回失败：{result}")
            return result
        except Exception as exc:
            last_error = exc
            print(f"{action_name} 第 {attempt}/{attempts} 次失败：{exc}")
            if attempt < attempts:
                time.sleep(attempt * 5)

    raise RuntimeError(f"{action_name} 重试 {attempts} 次仍失败：{last_error}")


def upload_to_mysql_api(games):
    """只上传有变化的资料，最后单独同步 Excel 全量 ID 排序。"""
    if not GAME_API_URL or not GAME_API_TOKEN:
        raise RuntimeError("缺少 GAME_API_URL 或 GAME_API_TOKEN")
    if not games:
        print("Excel 总表没有游戏记录，跳过数据库同步")
        return

    clean_games = []
    ordered_ids = []
    seen_ids = set()
    hash_items = []

    for row in games:
        item = clean_excel_row(row)
        game_id = str(item.get("game_id", "")).strip()
        if not game_id or game_id.lower() == "nan":
            continue
        if game_id in seen_ids:
            raise RuntimeError(f"Excel 总表存在重复 game_id：{game_id}，请先处理重复记录")
        seen_ids.add(game_id)
        item["game_id"] = game_id

        # 摘要字段统一转成字符串；与 PHP 数据库字段映射一致。
        for field in (
            "name", "icon_link", "storage_baidu", "storage_tianyi",
            "storage_xunlei", "password", "detail_url"
        ):
            value = item.get(field, "")
            item[field] = "" if value is None else str(value)

        clean_games.append(item)
        ordered_ids.append(game_id)
        hash_items.append({"game_id": game_id, "data_hash": game_data_hash(item)})

    if not clean_games:
        print("Excel 总表没有有效 game_id，跳过数据库同步")
        return

    print("\n===================================")
    print("开始检查游戏资料变化")
    print("Excel 总表数量：", len(clean_games))
    check_result = api_post(
        {"action": "check_changes", "items": hash_items},
        "检查游戏资料变化"
    )

    changed_ids = {str(value).strip() for value in check_result.get("changed_ids", [])}
    changed_games = [item for item in clean_games if item["game_id"] in changed_ids]

    print("需要新增或更新：", len(changed_games))
    print("资料未变化、跳过：", len(clean_games) - len(changed_games))

    batch_size = 100
    uploaded_count = 0
    total_batches = (len(changed_games) + batch_size - 1) // batch_size

    for start in range(0, len(changed_games), batch_size):
        batch = changed_games[start:start + batch_size]
        batch_number = start // batch_size + 1
        result = api_post(
            {"games": batch},
            f"上传变化资料第 {batch_number}/{total_batches} 批"
        )
        uploaded_count += int(result.get("upserted", len(batch)))
        print(f"第 {batch_number}/{total_batches} 批成功，处理 {result.get('upserted', len(batch))} 条")
        time.sleep(1)

    # 只有所有变化资料都上传成功后才同步顺序；失败会直接抛错，不会执行到这里。
    print("开始按 Excel 总表完整顺序同步 sort_order……")
    order_result = api_post(
        {"action": "sync_order", "game_ids": ordered_ids},
        "同步 Excel 排序"
    )

    print("\n===================================")
    print("数据库增量同步完成")
    print("总表游戏数：", len(clean_games))
    print("新增或更新资料数：", uploaded_count)
    print("资料未变化跳过数：", len(clean_games) - len(changed_games))
    print("排序同步数：", order_result.get("ordered", 0))
    print("不在 Excel 中、保留在数据库末尾的记录：", order_result.get("unlisted", 0))
    print("===================================")

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

        raise RuntimeError(
            "网站登录失败，程序结束"
        )

    # 先重试历史详情页任务
    retry_failed_details()

    # 读取总表
    all_data = load_excel(MASTER_FILE)

    # 已采集 ID
    seen_ids = set()

    for item in all_data:

        gid = str(
            item.get("game_id", "")
        ).strip()

        if gid and gid.lower() != "nan":
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

    # 读取当天数据
    daily_data = load_excel(daily_file)

    daily_ids = set()

    for item in daily_data:

        gid = str(
            item.get("game_id", "")
        ).strip()

        if gid and gid.lower() != "nan":
            daily_ids.add(gid)

    # 日期变化，重置每日计数
    if progress["date"] != today:

        progress["date"] = today

        progress["daily_count"] = len(daily_ids)

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

        if game_list is None:

            print("列表页请求失败，停止本次任务")
            break

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

        # 采集游戏信息
        result = get_game_data(
            game_id,
            detail_url
        )

        if result:

            # 将详情页任务加入失败重试队列
            queue_detail_task(
                game_id,
                detail_url
            )

            all_data.append(result)
            daily_data.append(result)

            seen_ids.add(game_id)
            daily_ids.add(game_id)

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

        # 推进当前页位置
        index += 1

        progress["page"] = page
        progress["index"] = index
        progress["daily_count"] = daily_count
        progress["date"] = today

        save_progress(progress)

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
    print("总表：", MASTER_FILE)
    print("今日文件：", daily_file)
    print("进度文件：", PROGRESS_FILE)
    print("===================================")
    
    # =================================================
    # 上传基本信息到 MySQL
    # =================================================

   

    # =================================================
    # 按 Excel 总表顺序同步全部游戏到 MySQL
    # =================================================

    try:
    
        upload_to_mysql_api(all_data)
    
    except Exception as e:
    
        print("\nExcel 总表同步数据库出现问题：", e)

    # 根据当天采集的全部游戏记录生成详情页任务
    print("\n===================================")
    print("开始根据今日游戏记录生成详情页任务...")
    print("===================================")

    rebuild_daily_detail_tasks(daily_data)

    # 开始上传详情页 HTML
    print("\n===================================")
    print("开始处理游戏详情页 HTML 上传任务...")
    print("===================================")

    retry_failed_details()

    print("\n===================================")
    print("本次详情页 HTML 任务处理结束")
    print("===================================")


if __name__ == "__main__":
    main()
