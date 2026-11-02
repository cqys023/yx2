
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

USERNAME = os.getenv("YYDS_USERNAME", "qqaazz").strip()
PASSWORD = os.getenv("YYDS_PASSWORD", "aa123456").strip()

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
