# -*- coding: utf-8 -*-
"""
QDII 基金限额公告监控脚本 v3.1
- 逐只打开天天基金公告页,抓取近 N 天限额相关公告
- 页面正常判据:表格里存在带日期的数据行(不再依赖链接数,根治AJAX慢加载误判)
- 支持翻页抓取(最多3页,防公告大户单页漏抓)
- 与 seen_announcements.json 对比,新公告推送飞书
- 推送成功后才更新已见记录(事务性,失败自动下轮重试)
"""

# ============================================================
# 一、全局配置
# ============================================================
import json
import os
import random
import re
import time
import datetime
import pathlib

import pandas as pd
import requests
from playwright.sync_api import sync_playwright

APP_DIR = pathlib.Path(__file__).resolve().parent

FUND_LIST_CSV = APP_DIR / "fund_list.csv"          # 基金清单(可选,存在则优先)
SEEN_RECORD_FILE = APP_DIR / "seen_announcements.json"  # 已推送记录(持久化到仓库)
EXCEL_FILE = APP_DIR / "近5天限额公告.xlsx"
LOG_FILE = APP_DIR / "fund_monitor.log"

RECENT_DAYS = 5            # 检查窗口:近几天的公告
PAGE_TIMEOUT = 30000       # 单页加载超时(毫秒)
MAX_ATTEMPTS = 3           # 页面异常时最大尝试次数
RETRY_WAIT_SECONDS = 2     # 重试前的等待秒数
MAX_PAGES = 3              # 每只基金最多抓前几页公告(防公告太多翻页漏抓)
TABLE_WAIT_TIMEOUT = 8000  # 等待公告表格骨架渲染的超时(毫秒)
DEBUG_DIR = APP_DIR / "debug"   # 页面异常时的截图/HTML留证目录

FEISHU_WEBHOOK_URL = os.environ.get("FEISHU_WEBHOOK_URL", "")

REQUEST_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# 限额公告关键词(正则,命中任意一条即保留)
KEYWORD_PATTERNS = [
    re.compile(r"调整.{0,25}大额申购"),
    re.compile(r"暂停(大额)?(申购|定期定额)"),
    re.compile(r"(申购|定期定额).{0,10}限制"),
    re.compile(r"限制(大额)?(申购|定期定额).{0,10}金额"),
    re.compile(r"恢复(大额)?(申购|定期定额)"),
    re.compile(r"规模上限"),
    re.compile(r"单日.{0,10}(金额|上限)"),
]

# ★ 完整基金列表(53只)——仅当仓库里不存在 fund_list.csv 时使用
FALLBACK_FUND_DICT = {
    # ---- 全球/其他 QDII(29只)----
    "017730": "嘉实全球产业升级",
    "501225": "景顺长城全球半导体芯片",
    "017653": "创金合信全球芯片产业",
    "019155": "易方达全球配置",
    "005698": "华夏全球科技先锋",
    "017091": "景顺长城纳斯达克科技",
    "018229": "易方达全球优质企业",
    "012920": "易方达全球成长精选",
    "000043": "嘉实美国成长",
    "002230": "华夏大中华",
    "016701": "银华海外数字经济量化",
    "270042": "广发纳斯达克100ETF联接",
    "017436": "华宝纳斯达克精选",
    "501312": "华宝海外科技",
    "008253": "华宝致远",
    "164212": "天弘全球新能源汽车",
    "019454": "中韩半导体",
    "006373": "国富全球科技互联",
    "270023": "广发全球精选",
    "501226": "长城全球新能源车",
    "002891": "华夏移动互联",
    "378006": "摩根全球新兴市场",
    "016664": "天弘全球高端制造",
    "539002": "建信新兴市场",
    "457001": "国富亚洲机会",
    "100055": "富国全球科技互联",
    "006555": "浦银全球智能科技",
 
