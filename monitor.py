# -*- coding: utf-8 -*-
"""
QDII 基金限额公告监控脚本 v3.6
- 逐只打开天天基金公告页,抓取"过去2天 + 未来4天"窗口内的限额/限购相关公告
- 通用抓取：不假设 table/ul/li DOM，抓取公告列表区域内 <a>，从其"最近行容器"解析日期

v3.6 相比 v3.5 的改动：

1) ★最重要：从"53只基金全跑完才统一推送一次"改成"每查完一只基金，
   立刻判断这只有没有新公告，有就马上推送+写入已见记录"。
   原来的写法如果跑到中途被 GitHub Actions 的 15 分钟超时强制杀掉，
   前面已经抓到的新公告会跟着一起丢失、你也收不到提醒；改成按基金
   为单位后，只要这只基金处理完了，它的新公告就已经安全落地。

2) 恢复"赎回噪音过滤"：QDII基金几乎每天都可能因为境外市场节假日/
   估值不确定，发布"XX年X月X日暂停申购、赎回及定期定额投资业务"这类
   常规通知，标题里同时出现"申购"和"赎回"但并不代表真的调整了限额。
   现在的规则是：标题里如果同时出现"赎回"，又没有"大额/限购/限额/
   限制/上限/金额/额度"这类真正表示限购限额的字眼，就判定为常规通知
   予以排除，避免刷屏。

3) 加入运行时间预算（软死线）：本轮运行超过 SOFT_DEADLINE_SECONDS后，
   不再继续查剩下的基金，而是直接结束本轮、保存已有结果。因为查询窗口
   本身覆盖"过去2天+未来4天"，这一轮没查完的基金，下一轮还会覆盖到，
   不会被永久漏掉；但可以避免被 workflow 的 timeout-minutes 硬杀死、
   导致 Excel/已见记录写到一半或者干脆没保存。

4) 缩短 goto 超时上限（90秒一次太长了，正常公告页几秒内就能加载完，
   等这么久基本等于在浪费本来就紧张的15分钟预算），改成更合理的区间。
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
import logging

import pandas as pd
import requests
from playwright.sync_api import sync_playwright

APP_DIR = pathlib.Path(__file__).resolve().parent

FUND_LIST_CSV = APP_DIR / "fund_list.csv"          # 基金清单(可选,存在则优先)
SEEN_RECORD_FILE = APP_DIR / "seen_announcements.json"  # 已推送记录(持久化到仓库)
EXCEL_FILE = APP_DIR / "限额公告.xlsx"
LOG_FILE = APP_DIR / "fund_monitor.log"

# 时间窗口：过去2天 + 未来4天（含端点）
PAST_DAYS = 2
FUTURE_DAYS = 4

# 基础超时/重试
# v3.6：90秒一次goto太长了，正常公告页几秒能加载完，缩短上限，
# 把省下来的时间留给后面的基金，减少被workflow超时强杀的概率
GOTO_TIMEOUTS = [12000, 20000, 30000]  # 第1/2/3次 goto 超时(毫秒)
MAX_ATTEMPTS = 3           # 页面异常时最大尝试次数
MAX_PAGES = 3              # 每只基金最多翻页数（仍保留硬上限，避免极端情况）
TABLE_WAIT_TIMEOUT = 8000  # 等待公告列表骨架渲染的超时(毫秒)
DEBUG_DIR = APP_DIR / "debug"   # 页面异常时的截图/HTML留证目录

MAX_DIAG_TITLES = 30  # 未命中关键词时，窗口内标题最多打印多少条（防日志爆）

# v3.6新增：运行时间预算（软死线），单位秒。
# workflow 的 timeout-minutes 建议设15分钟，这里留够导出Excel、
# git提交推送的余量，实际抓取阶段最多跑到这个时长就主动收尾。
SOFT_DEADLINE_SECONDS = 11 * 60

FEISHU_WEBHOOK_URL = os.environ.get("FEISHU_WEBHOOK_URL", "")

REQUEST_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# 限额/限购关键词（子串匹配，命中任意一个即视为"可能相关"）
LIMIT_KEYWORDS_STRICT = [
    "限购", "限额", "限大额", "大额申购",
    "暂停申购", "恢复申购",
    "暂停大额申购", "恢复大额申购",
    "申购上限", "单日上限", "规模上限",
    "暂停定投", "暂停定期定额", "恢复定期定额",
    "暂停转换转入", "恢复转换转入",
]

# v3.6：真正表示"限购/限额调整"的强信号词。
# 标题里如果同时出现"赎回"，必须再命中下面任意一个词，才认定是
# 真正的限额公告；否则判定为"境外市场节假日/估值不确定导致的
# 当天暂停申购赎回"这类常规通知，予以排除。
STRONG_LIMIT_SIGNALS = ["大额", "限购", "限额", "限制", "上限", "金额", "额度"]

# 明确排除：这些通常不是"限额/限购"类
EXCLUDE_KEYWORDS = [
    "年度报告", "中期报告", "季度报告", "定期报告",
    "招募说明书", "更新招募说明书", "基金合同", "托管协议",
    "分红", "收益分配",
    "净值", "临时公告",
    "基金经理", "经理变更",
    "节假日", "市场节假日",
    "暂停申购赎回安排", "暂停赎回安排", "申购赎回安排",
    "开放日安排", "交易日安排",
    "清明节", "劳动节", "端午节", "中秋节", "国庆节", "春节", "元旦",
]


# ============================================================
# 二、日志(同时输出到控制台和文件)
# ============================================================
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
# 完整基金列表(53只)——仅当仓库里不存在 fund_list.csv 时使用
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
    "001668": "汇添富全球移动互联",
    "017144": "华宝海外新能源汽车",
    # ---- 纳斯达克100系(16只)----
    "160213": "国泰纳斯达克100指数",
    "016055": "博时纳斯达克100ETF联接",
    "040046": "华安纳斯达克100ETF联接",
    "019172": "摩根纳斯达克100指数",
    "019441": "万家纳斯达克100指数",
    "018043": "天弘纳斯达克100指数",
    "016532": "嘉实纳斯达克100ETF联接",
    "019547": "招商纳斯达克100ETF联接",
    "000834": "大成纳斯达克100ETF联接",
    "015299": "华夏纳斯达克100ETF联接",
    "161130": "易方达纳斯达克100ETF联接",
    "019736": "宝盈纳斯达克100指数",
    "016452": "南方纳斯达克100指数",
    "539001": "建信纳斯达克100指数",
    "019524": "华泰柏瑞纳斯达克100ETF联接",
    "018966": "汇添富纳斯达克100ETF联接",
    # ---- 标普系(8只)----
    "017641": "摩根标普500指数(QDII)",
    "050025": "博时标普500ETF联接A",
    "519981": "长信标普100",
    "161125": "易方达标普500指数人民币",
    "017028": "国泰标普500ETF发起联接",
    "007721": "天弘标普500发起(QDII)",
    "018064": "华夏标普500ETF发起式",
    "096001": "大成标普500等权重指数",
}


def load_fund_dict():
    """优先读 fund_list.csv(代码,名称);不存在则用脚本内置完整列表"""
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
# 四、日期窗口 + 文字/日期处理
# ============================================================
def fetch_window_dates(past_days, future_days):
    """返回窗口日期：[今天-past_days, 今天+future_days]（含端点）"""
    today = datetime.date.today()
    start_date = today - datetime.timedelta(days=past_days)
    end_date = today + datetime.timedelta(days=future_days)
    return start_date, end_date


def clean_text(value):
    """清理网页文字"""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\n", "").replace("\r", "").replace("\t", "")
    text = text.replace(" ", "").replace("　", "")
    return text.strip()


def normalize_title(title):
    """统一公告标题格式"""
    title = clean_text(title)
    title = title.replace("（", "(").replace("）", ")")
    return title


def parse_date(text):
    """
    从文字中提取日期，兼容多种写法：
    2026-08-29 / 2026/08/29 / 2026.08.29 / 2026年08月29日
    """
    text = clean_text(text)
    match = re.search(r"(20\d{2})[-年/.](\d{1,2})[-月/.](\d{1,2})", text)
    if not match:
        return None
    try:
        return datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def format_date(value):
    """格式化日期为 YYYY-MM-DD"""
    if value is None:
        return ""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    return str(value)


def is_limit_title(title):
    """
    判断公告标题是否限额相关。

    v3.6：恢复"赎回噪音过滤"——标题里如果同时出现"赎回"，
    又没有命中 STRONG_LIMIT_SIGNALS 里的任何一个真正表示限购限额的
    字眼，判定为境外市场节假日/估值不确定导致的常规"当天暂停申购
    赎回"通知，不算限额公告，予以排除。
    """
    t = normalize_title(title)

    if any(x in t for x in EXCLUDE_KEYWORDS):
        return False

    if not any(k in t for k in LIMIT_KEYWORDS_STRICT):
        return False

    if "赎回" in t and not any(k in t for k in STRONG_LIMIT_SIGNALS):
        return False

    return True


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


def make_key(record):
    return f"{record['基金代码']}|{record['公告日期']}|{record['公告标题']}"


# ============================================================
# 六、飞书推送(文本消息)
# ============================================================
def send_feishu(new_records):
    if not new_records:
        return
    if not FEISHU_WEBHOOK_URL:
        logger.info("未配置飞书 Webhook,跳过推送")
        return
    new_count = len(new_records)
    lines = [f"发现新的限额公告：{new_count}条"]
    for r in new_records:
        lines.append(f"- {r['基金代码']} {r['基金名称']} {r['公告日期']} {r['公告标题']}")
        lines.append(f"  {r['公告链接']}")
    payload = {"msg_type": "text", "content": {"text": "\n".join(lines)}}
    resp = requests.post(FEISHU_WEBHOOK_URL, json=payload, timeout=10)
    if resp.status_code == 200:
        logger.info("飞书消息发送成功")
    else:
        raise RuntimeError(f"飞书发送失败:{resp.status_code} {resp.text}")


# ============================================================
# 七、通用公告抓取（不假设table/ul/li具体结构）
# ============================================================
def extract_raw_items(page):
    """
    通用抓取：尽量限定在"公告列表区域"内找 <a>，找不到再退回整页；
    同时过滤明显不是详情链接的 href（# / javascript: / 空）
    """
    try:
        raw_items = page.evaluate(
            r"""
            () => {
                const result = [];
                const clean = (value) => {
                    return String(value || "")
                        .replace(/[\s\u3000]+/g, "")
                        .trim();
                };
                const scope =
                    document.querySelector("#jjgg, .jjgg, .fundNoticeList, #gg, .gg, .txt_in")
                    || document;
                const links = Array.from(scope.querySelectorAll("a"));
                for (const link of links) {
                    const title = clean(link.innerText || link.textContent);
                    if (!title) continue;
                    const href = (link.getAttribute("href") || "").trim();
                    if (!href || href === "#" || href.toLowerCase().startsWith("javascript:")) continue;
                    let row = link.closest("tr") || link.closest("li");
                    if (!row) row = link.closest(".list-item, .notice-item, .item, .box");
                    if (!row) row = link.parentElement;
                    if (!row) row = link;
                    const rowText = clean(row.innerText || row.textContent);
                    result.push({ title: title, rowText: rowText, href: href });
                }
                return result;
            }
            """
        )
    except Exception as error:
        logger.info("      页面公告提取失败：%s", error)
        return []
    return raw_items or []


def parse_records_from_raw_items(raw_items, code, fund_name, source_url, start_date, end_date):
    """
    从 extract_raw_items 抓到的原始条目里，解析出：
    - dated_item_count：能解析出日期的条目数（用于判定公告是否加载）
    - window_titles：落在日期窗口内的全部公告标题（无论是否命中关键词，用于诊断漏抓）
    - records：落在日期窗口内 且 命中限额关键词的公告（真正要推送/导出的）
    - min_date_in_page：该页所有"可解析日期条目"的最早日期（用于翻页提前停止）
    """
    dated_item_count = 0
    window_titles = []
    records = []
    min_date_in_page = None

    for item in raw_items:
        title = normalize_title(item.get("title", ""))
        row_text = clean_text(item.get("rowText", ""))
        href = item.get("href", "") or ""

        if len(title) < 8:
            continue

        announcement_date = parse_date(row_text) or parse_date(title)
        if announcement_date is None:
            continue

        dated_item_count += 1
        if (min_date_in_page is None) or (announcement_date < min_date_in_page):
            min_date_in_page = announcement_date

        if not (start_date <= announcement_date <= end_date):
            continue

        window_titles.append((format_date(announcement_date), title))

        if not is_limit_title(title):
            continue

        if href.startswith("//"):
            url = "https:" + href
        elif href.startswith("/"):
            url = "https://fundf10.eastmoney.com" + href
        elif href.startswith("http"):
            url = href
        else:
            url = source_url

        records.append({
            "基金代码": code,
            "基金名称": fund_name,
            "公告日期": format_date(announcement_date),
            "公告标题": title,
            "公告链接": url,
        })

    return dated_item_count, window_titles, records, min_date_in_page


def save_debug_snapshot(page, code, tag):
    """页面异常时保存截图和HTML,便于排查"""
    try:
        DEBUG_DIR.mkdir(exist_ok=True)
        ts = datetime.datetime.now().strftime("%H%M%S")
        page.screenshot(path=str(DEBUG_DIR / f"{code}_{tag}_{ts}.png"), full_page=False)
        with open(DEBUG_DIR / f"{code}_{tag}_{ts}.html", "w", encoding="utf-8") as f:
            f.write(page.content())
    except Exception:
        pass


def backoff_sleep(attempt_index_zero_based: int):
    """指数退避 + 随机抖动（attempt=0/1/2 -> 约2/4/8秒）"""
    base = 2 ** (attempt_index_zero_based + 1)
    wait = base + random.uniform(0.0, 1.5)
    time.sleep(wait)


def find_clickable_next_button(page):
    """更稳健地找"下一页"：必须可见，并且 class/aria-disabled 不像 disabled。"""
    candidates = page.locator("a", has_text="下一页")
    try:
        n = candidates.count()
    except Exception:
        return None

    for i in range(n):
        btn = candidates.nth(i)
        try:
            if not btn.is_visible():
                continue
            cls = (btn.get_attribute("class") or "").lower()
            aria = (btn.get_attribute("aria-disabled") or "").lower()
            if "disabled" in cls or "disable" in cls or "nobtn" in cls or "ban" in cls:
                continue
            if aria in ("true", "1"):
                continue
            return btn
        except Exception:
            continue
    return None


# ============================================================
# 八、单只基金查询(通用抓取+重试+翻页+未命中诊断+失败留证)
# ============================================================
def query_one_fund(page, index, total, code, fund_name, start_date, end_date):
    code = str(code).zfill(6)
    source_url = f"https://fundf10.eastmoney.com/jjgg_{code}.html"
    logger.info("[%02d/%d] 正在查询：%s %s", index, total, code, fund_name)

    for attempt in range(MAX_ATTEMPTS):
        try:
            timeout = GOTO_TIMEOUTS[min(attempt, len(GOTO_TIMEOUTS) - 1)]
            page.goto(source_url, wait_until="domcontentloaded", timeout=timeout)

            try:
                page.wait_for_selector("a", state="attached", timeout=TABLE_WAIT_TIMEOUT)
            except Exception:
                pass
            page.wait_for_timeout(600)

            raw_items = extract_raw_items(page)
            dated_item_count, window_titles, records, min_date_in_page = parse_records_from_raw_items(
                raw_items, code, fund_name, source_url, start_date, end_date
            )

            if dated_item_count == 0:
                if attempt < MAX_ATTEMPTS - 1:
                    logger.info("  第%d次页面异常（未解析到任何带日期的公告条目），退避后重试...",
                                attempt + 1)
                    backoff_sleep(attempt)
                    continue
                save_debug_snapshot(page, code, "tablefail")
                logger.info("没有找到额度相关公告：%s %s（连续%d次页面异常，已存debug截图）",
                            code, fund_name, MAX_ATTEMPTS)
                return []

            if (min_date_in_page is not None) and (min_date_in_page < start_date):
                pass
            else:
                for _page_no in range(2, MAX_PAGES + 1):
                    next_btn = find_clickable_next_button(page)
                    if next_btn is None:
                        break
                    try:
                        next_btn.click()
                    except Exception:
                        break

                    try:
                        page.wait_for_selector("a", state="attached", timeout=TABLE_WAIT_TIMEOUT)
                    except Exception:
                        pass
                    page.wait_for_timeout(900)

                    more_raw_items = extract_raw_items(page)
                    _more_dated, more_window_titles, more_records, more_min_date = parse_records_from_raw_items(
                        more_raw_items, code, fund_name, source_url, start_date, end_date
                    )
                    records.extend(more_records)
                    window_titles.extend(more_window_titles)

                    if (more_min_date is not None) and (more_min_date < start_date):
                        break

            unique, seen_keys = [], set()
            for r in records:
                k = (r["基金代码"], r["公告日期"], r["公告标题"])
                if k not in seen_keys:
                    seen_keys.add(k)
                    unique.append(r)
            records = unique

            if records:
                logger.info("找到额度相关公告：%s %s", code, fund_name)
                for r in records:
                    logger.info("- %s %s", r["公告日期"], r["公告标题"])
            else:
                shown = set()
                printed = 0
                for d, t in window_titles:
                    if (d, t) in shown:
                        continue
                    shown.add((d, t))
                    logger.info("  窗口内公告(未命中关键词)：%s %s", d, t)
                    printed += 1
                    if printed >= MAX_DIAG_TITLES:
                        logger.info("  ...窗口内未命中标题太多，仅展示前%d条", MAX_DIAG_TITLES)
                        break
                logger.info("没有找到额度相关公告：%s %s", code, fund_name)

            return records

        except Exception as error:
            if attempt < MAX_ATTEMPTS - 1:
                logger.info("  查询异常(%s)，退避后重试...", error)
                backoff_sleep(attempt)
            else:
                save_debug_snapshot(page, code, "gotofail")
                logger.info("查询失败：%s %s", code, error)

    return []


# ============================================================
# 九、全量查询
# v3.6核心改动：不再"全部查完再统一推送"，改成每查完一只基金
# 立刻判断新公告、立刻推送+写入已见记录，避免中途被超时杀掉时
# 已经抓到的新公告白白丢失。同时加入运行时间预算，快到期时提前
# 收尾，剩下的基金留给下一轮（查询窗口本身有过去2天余量，不会
# 永久漏掉）。
# ============================================================
def query_all_funds(page, fund_dict, start_date, end_date, seen, deadline_ts):
    total = len(fund_dict)
    all_records = []
    processed_count = 0

    for idx, (code, name) in enumerate(fund_dict.items(), start=1):
        if time.time() >= deadline_ts:
            remaining = total - processed_count
            logger.info("已接近本轮时间预算上限，提前结束本轮，剩余%d只基金留给下一轮继续检查",
                        remaining)
            break

        records = query_one_fund(page, idx, total, code, name, start_date, end_date)
        all_records.extend(records)
        processed_count += 1

        # 立刻判断这只基金有没有新公告，有就马上推送+写入已见记录
        new_records_this_fund = []
        for r in records:
            key = make_key(r)
            if key not in seen:
                r["_key"] = key
                new_records_this_fund.append(r)

        if new_records_this_fund:
            logger.info("发现新的限额公告：%d条（%s %s）", len(new_records_this_fund), code, name)
            try:
                send_feishu(new_records_this_fund)
                seen.update(r["_key"] for r in new_records_this_fund)
                save_seen_records(seen)
            except Exception as e:
                logger.info("推送失败(%s)，本条记录不写入历史，下轮自动重试", e)

        time.sleep(random.uniform(0.8, 1.8))

    return all_records, processed_count


# ============================================================
# 十、主流程
# ============================================================
def main():
    start_time = time.time()
    deadline_ts = start_time + SOFT_DEADLINE_SECONDS

    start_date, end_date = fetch_window_dates(PAST_DAYS, FUTURE_DAYS)
    logger.info("本次检查窗口：%s ~ %s（过去%d天 + 未来%d天）",
                start_date, end_date, PAST_DAYS, FUTURE_DAYS)

    fund_dict = load_fund_dict()
    seen = load_seen_records()

    all_records = []
    processed_count = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-gpu",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )
        context = browser.new_context(user_agent=REQUEST_HEADERS["User-Agent"])

        def block_unneeded(route):
            try:
                r = route.request
                if r.resource_type in ("image", "media", "font"):
                    return route.abort()
            except Exception:
                pass
            return route.continue_()

        context.route("**/*", block_unneeded)

        page = context.new_page()
        page.set_default_timeout(30000)
        page.set_default_navigation_timeout(30000)

        try:
            all_records, processed_count = query_all_funds(
                page, fund_dict, start_date, end_date, seen, deadline_ts
            )
        finally:
            browser.close()

    # 保存 Excel(无论是否有公告都生成,便于 artifact 查看全量；
    # 即使因为时间预算提前收尾，也只会缺剩余未查基金的数据，不影响已查部分)
    df = pd.DataFrame(all_records, columns=["基金代码", "基金名称", "公告日期", "公告标题", "公告链接"])
    df.to_excel(EXCEL_FILE, index=False)

    elapsed = time.time() - start_time
    logger.info("Excel已保存：%s（本轮共检查%d/%d只基金，耗时%.0f秒）",
                EXCEL_FILE, processed_count, len(fund_dict), elapsed)
    logger.info("本轮结束（新公告已在查询过程中逐只推送完毕，此处不再重复推送）")


if __name__ == "__main__":
    main()
