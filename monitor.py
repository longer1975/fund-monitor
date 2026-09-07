# -*- coding: utf-8 -*-
"""
QDII 基金限额公告监控脚本
- 逐只打开天天基金公告页,抓取近 N 天限额相关公告
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

FUND_LIST_CSV = APP_DIR / "fund_list.csv"          # 基金清单(两列:代码,名称)
SEEN_RECORD_FILE = APP_DIR / "seen_announcements.json"  # 已推送记录(持久化到仓库)
EXCEL_FILE = APP_DIR / "近5天限额公告.xlsx"
LOG_FILE = APP_DIR / "fund_monitor.log"

RECENT_DAYS = 2            # 只看近几天的公告
WAIT_AFTER_LOAD = 1500     # 页面加载后固定等待毫秒数(保证 Ajax 表格渲染,勿调低,防漏抓)
PAGE_TIMEOUT = 30000       # 单页加载超时(毫秒)
MAX_ATTEMPTS = 3           # 页面异常时最大尝试次数
RETRY_WAIT_SECONDS = 2     # 重试前的等待秒数

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

# ★ 备用基金列表:仅当仓库里不存在 fund_list.csv 时使用
#   如果你原 monitor.py 里有完整的 53 只基金 fund_dict,请整体替换下面这个字典
FALLBACK_FUND_DICT = {
    "016701": "银华海外数字经济量化选股混合",
    "019736": "宝盈纳斯达克100指数",
    "016055": "博时纳斯达克100ETF联接",
    "040046": "华安纳斯达克100ETF联接",
    "019172": "摩根纳斯达克100指数",
    "019441": "万家纳斯达克100指数",
    "018043": "天弘纳斯达克100指数",
    "016532": "嘉实纳斯达克100ETF联接",
    "019547": "招商纳斯达克100ETF联接",
    "007721": "天弘标普500发起(QDII)",
    "018064": "华夏标普500ETF发起式",
    "096001": "大成标普500等权重指数",
    "270042": "广发纳斯达克100ETF联接",
    "000041": "华夏全球精选混合(QDII)",
    "000834": "国泰纳斯达克100ETF联接",
    "050025": "博时标普500ETF联接",
}

# ============================================================
# 二、日志(同时输出到控制台和文件)
# ============================================================
import logging

logger = logging.getLogger("fund_monitor")
logger.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
_sh = logging.StreamHandler()
_sh.setFormatter(_fmt)
logger.addHandler(_sh)
try:
    _fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    _fh.setFormatter(_fmt)
    logger.addHandler(_fh)
except Exception:
    pass


# ============================================================
# 三、基金清单加载
# ============================================================
def load_fund_dict():
    """优先读 fund_list.csv(代码,名称);不存在则用脚本内置备用列表"""
    if FUND_LIST_CSV.exists():
        try:
            df = pd.read_csv(FUND_LIST_CSV, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            code_col = next((c for c in df.columns
                             if c in ("代码", "code", "基金代码")), df.columns[0])
            name_col = next((c for c in df.columns
                             if c in ("名称", "name", "基金名称", "简称")), df.columns[1])
            fund_dict = {}
            for _, row in df.iterrows():
                code = str(row[code_col]).strip().split(".")[0].zfill(6)
                name = str(row[name_col]).strip()
                if code and code != "nan":
                    fund_dict[code] = name
            logger.info("已从 %s 读取 %d 只基金", FUND_LIST_CSV.name, len(fund_dict))
            return fund_dict
        except Exception as e:
            logger.info("读取基金列表失败(%s),改用脚本内置列表", e)
    else:
        logger.info("未找到 %s,使用脚本内置列表(%d只)",
                    FUND_LIST_CSV.name, len(FALLBACK_FUND_DICT))
    return dict(FALLBACK_FUND_DICT)


# ============================================================
# 四、日期窗口
# ============================================================
def fetch_recent_start_date(recent_days):
    """返回起始日期(含今天往前推 recent_days 天)"""
    return datetime.date.today() - datetime.timedelta(days=recent_days - 1)


# ============================================================
# 五、已见记录读写
# ============================================================
def load_seen_records():
    if SEEN_RECORD_FILE.exists():
        try:
            with open(SEEN_RECORD_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return set(data) if isinstance(data, list) else set()
        except Exception as e:
            logger.info("读取已见记录失败(%s),按空记录处理", e)
    return set()


def save_seen_records(seen_set):
    with open(SEEN_RECORD_FILE, "w", encoding="utf-8") as f:
        json.dump(sorted(seen_set), f, ensure_ascii=False, indent=2)


# ============================================================
# 六、飞书推送(文本消息)
# ============================================================
def send_feishu(new_records):
    if not FEISHU_WEBHOOK_URL:
        logger.info("未配置飞书 Webhook,跳过推送")
        return
    new_count = len(new_records)
    lines = [f"发现新的限额公告：{new_count}条"]
    for r in new_records:
        lines.append(f"- {r['基金代码']} {r['公告日期']} {r['公告标题']}")
        lines.append(f"  {r['公告链接']}")
    payload = {"msg_type": "text", "content": {"text": "\n".join(lines)}}
    resp = requests.post(FEISHU_WEBHOOK_URL, json=payload, timeout=10)
    if resp.status_code == 200:
        logger.info("飞书消息发送成功")
    else:
        raise RuntimeError(f"飞书发送失败:{resp.status_code} {resp.text}")


# ============================================================
# 七、页面公告提取
# ============================================================
def extract_page_records(*, page, code, fund_name, source_url, start_date, end_date):
    rows = page.locator("table tr")
    row_count = rows.count()
    records = []
    for i in range(row_count):
        try:
            cells = rows.nth(i).locator("td")
            if cells.count() < 2:
                continue
            date_text = cells.nth(0).inner_text().strip()
            try:
                row_date = datetime.datetime.strptime(date_text, "%Y-%m-%d").date()
            except ValueError:
                continue
            if not (start_date <= row_date <= end_date):
                continue
            link = rows.nth(i).locator("a").first
            if link.count() == 0:
                continue
            title = link.inner_text().strip()
            if not title:
                continue
            if not any(p.search(title) for p in KEYWORD_PATTERNS):
                continue
            href = link.get_attribute("href") or ""
            if href.startswith("/"):
                url = "https://fundf10.eastmoney.com" + href
            elif href.startswith("http"):
                url = href
            else:
                url = source_url
            records.append({
                "基金代码": code,
                "基金名称": fund_name,
                "公告日期": date_text,
                "公告标题": title,
                "公告链接": url,
            })
        except Exception:
            continue
    return records


# ============================================================
# 八、单只基金查询(v3:智能重试,只在页面疑似异常时重试)
# ============================================================
def query_one_fund(page, index, total, code, fund_name, start_date, end_date):
    code = str(code).zfill(6)
    source_url = f"https://fundf10.eastmoney.com/jjgg_{code}.html"
    logger.info("[%02d/%d] 正在查询：%s %s", index, total, code, fund_name)

    records = []
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            page.goto(source_url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)
            try:
                page.wait_for_selector("a", state="attached", timeout=3000)
            except Exception:
                pass
            page.wait_for_timeout(WAIT_AFTER_LOAD)

            records = extract_page_records(
                page=page, code=code, fund_name=fund_name,
                source_url=source_url, start_date=start_date, end_date=end_date,
            )
            if records:
                logger.info("找到额度相关公告：%s %s", code, fund_name)
                for r in records:
                    logger.info("- %s %s", r["公告日期"], r["公告标题"])
                return records

            # 没抓到时先判断:页面加载正常(链接多)= 真没有;链接少 = 页面异常,值得重试
            try:
                link_count = page.evaluate("() => document.querySelectorAll('a').length")
            except Exception:
                link_count = 0

            if link_count >= 15:
                logger.info("没有找到额度相关公告：%s %s", code, fund_name)
                return []

            if attempt < MAX_ATTEMPTS:
                logger.info("  第%d次页面异常（仅%d个链接），%d秒后重试...",
                            attempt, link_count, RETRY_WAIT_SECONDS)
                time.sleep(RETRY_WAIT_SECONDS)
            else:
                logger.info("没有找到额度相关公告：%s %s（页面连续%d次加载异常）",
                            code, fund_name, MAX_ATTEMPTS)

        except Exception as error:
            if attempt < MAX_ATTEMPTS:
                logger.info("  查询异常(%s)，%d秒后重试...", error, RETRY_WAIT_SECONDS)
                time.sleep(RETRY_WAIT_SECONDS)
            else:
                logger.info("查询失败：%s %s", code, error)
    return records


# ============================================================
# 九、全量查询
# ============================================================
def query_all_funds(page, fund_dict, start_date, end_date):
    total = len(fund_dict)
    all_records = []
    for idx, (code, name) in enumerate(fund_dict.items(), start=1):
        all_records.extend(
            query_one_fund(page, idx, total, code, name, start_date, end_date)
        )
        time.sleep(random.uniform(1.0, 2.0))  # 每只之间随机间隔,降低风控概率
    return all_records


# ============================================================
# 十、主流程
# ============================================================
def main():
    start_date = fetch_recent_start_date(RECENT_DAYS)
    end_date = datetime.date.today()
    logger.info("本次检查窗口：%s ~ %s", start_date, end_date)

    fund_dict = load_fund_dict()
    seen = load_seen_records()

    all_records = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=REQUEST_HEADERS["User-Agent"])
        page = context.new_page()
        try:
            all_records = query_all_funds(page, fund_dict, start_date, end_date)
        finally:
            browser.close()

    # 保存 Excel(无论是否有公告都生成,便于 artifact 查看全量)
    df = pd.DataFrame(all_records, columns=["基金代码", "基金名称", "公告日期", "公告标题", "公告链接"])
    df.to_excel(EXCEL_FILE, index=False)
    logger.info("Excel已保存：%s", EXCEL_FILE)

    # 对比已见记录,筛出新公告
    new_records = []
    for r in all_records:
        key = f"{r['基金代码']}|{r['公告日期']}|{r['公告标题']}"
        if key not in seen:
            r["_key"] = key
            new_records.append(r)

    if not new_records:
        logger.info("本轮没有新的限额公告")
        return

    logger.info("发现新的限额公告：%d条", len(new_records))
    for r in new_records:
        logger.info("- %s %s %s", r["基金代码"], r["公告日期"], r["公告标题"])

    try:
        send_feishu(new_records)
    except Exception as e:
        logger.info("推送失败(%s),记录不更新,下轮自动重试", e)
        return

    # ★ 只有推送成功才写记录(事务性)
    seen.update(r["_key"] for r in new_records)
    save_seen_records(seen)


if __name__ == "__main__":
    main()
