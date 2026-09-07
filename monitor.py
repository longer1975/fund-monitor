# -*- coding: utf-8 -*-
"""
QDII 基金限额公告监控脚本 v3.4
- 逐只打开天天基金公告页,抓取“过去2天 + 未来4天”窗口内的限额/限购相关公告
- v3.3核心：不假设 table/ul/li DOM，通用抓取 <a> + 最近行容器解析日期
- v3.4改动：
  1) 时间窗口改为：今天往前2天 ~ 今天往后4天（含端点）
  2) 降噪：优先限定在“公告列表区域”范围内找 <a>，并过滤明显非详情链接
  3) 降噪：关键词改为“强匹配 + 排除项”
  4) 降噪：未命中诊断标题最多打印 MAX_DIAG_TITLES 条
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
import logging

import pandas as pd
import requests
from playwright.sync_api import sync_playwright

APP_DIR = pathlib.Path(__file__).resolve().parent

FUND_LIST_CSV = APP_DIR / "fund_list.csv"                # 基金清单(可选,存在则优先)
SEEN_RECORD_FILE = APP_DIR / "seen_announcements.json"   # 已推送记录(持久化到仓库)
EXCEL_FILE = APP_DIR / "限额公告.xlsx"
LOG_FILE = APP_DIR / "fund_monitor.log"

# 时间窗口：过去2天 + 未来4天（含端点）
PAST_DAYS = 2
FUTURE_DAYS = 4

PAGE_TIMEOUT = 30000       # 单页加载超时(毫秒)
MAX_ATTEMPTS = 3           # 页面异常时最大尝试次数
RETRY_WAIT_SECONDS = 2     # 重试前的等待秒数
MAX_PAGES = 3              # 每只基金最多抓前几页公告(防公告太多翻页漏抓)
TABLE_WAIT_TIMEOUT = 8000  # 等待公告列表骨架渲染的超时(毫秒)
DEBUG_DIR = APP_DIR / "debug"   # 页面异常时的截图/HTML留证目录

MAX_DIAG_TITLES = 30       # 未命中关键词时，窗口内标题最多打印多少条（防日志爆）

FEISHU_WEBHOOK_URL = os.environ.get("FEISHU_WEBHOOK_URL", "")

REQUEST_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "zh-CN,zh;q=0.9",
}

# 更严格的“限额/限购”关键词（强匹配，减少误抓）
LIMIT_KEYWORDS_STRICT = [
    "限购", "限额", "限大额", "大额申购",
    "暂停申购", "恢复申购",
    "暂停大额申购", "恢复大额申购",
    "申购上限", "单日上限", "规模上限",
    "暂停定投", "暂停定期定额", "恢复定期定额",
    "暂停转换转入", "恢复转换转入",
]

# 明确排除：这些通常不是“限额/限购”类
EXCLUDE_KEYWORDS = [
    "年度报告", "中期报告", "季度报告", "定期报告",
    "招募说明书", "更新招募说明书", "基金合同", "托管协议",
    "分红", "收益分配",
    "净值", "临时公告",
    "基金经理", "经理变更",
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


# ============================================================
# 四、日期窗口 + 文字/日期处理
# ============================================================
def fetch_window_dates(past_days, future_days):
    """返回窗口日期：[今天- past_days, 今天+ future_days]（含端点）"""
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
    """判断公告标题是否限额相关（强匹配 + 排除项）"""
    t = normalize_title(title)
    if any(x in t for x in EXCLUDE_KEYWORDS):
        return False
    return any(k in t for k in LIMIT_KEYWORDS_STRICT)


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
# 七、通用公告抓取（不假设table/ul/li具体结构）
# ============================================================
def extract_raw_items(page):
    """
    通用抓取：尽量限定在“公告列表区域”内找 <a>，找不到再退回整页；
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

                // 尽量把范围限定在“公告列表”附近（不同基金页结构不一）
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
    - dated_item_count：窗口无关的、能解析出日期的条目总数（用于判断页面是否真的加载了公告数据）
    - window_titles：落在日期窗口内的全部公告标题（无论是否命中关键词，用于诊断漏抓）
    - records：落在日期窗口内 且 命中限额关键词的公告（真正要推送/导出的）
    """
    dated_item_count = 0
    window_titles = []
    records = []

    for item in raw_items:
        title = normalize_title(item.get("title", ""))
        row_text = clean_text(item.get("rowText", ""))
        href = item.get("href", "") or ""

        if len(title) < 8:
            continue

        # 优先从整行文字找日期，找不到再退回标题本身
        announcement_date = parse_date(row_text)
        if announcement_date is None:
            announcement_date = parse_date(title)

        if announcement_date is None:
            continue

        dated_item_count += 1

        if not (start_date <= announcement_date <= end_date):
            continue

        window_titles.append((format_date(announcement_date), title))

        if not is_limit_title(title):
            continue

        if href.startswith("/"):
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

    return dated_item_count, window_titles, records


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


# ============================================================
# 八、单只基金查询(通用抓取+重试+翻页+未命中诊断+失败留证)
# ============================================================
def query_one_fund(page, index, total, code, fund_name, start_date, end_date):
    code = str(code).zfill(6)
    source_url = f"https://fundf10.eastmoney.com/jjgg_{code}.html"
    logger.info("[%02d/%d] 正在查询：%s %s", index, total, code, fund_name)

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            page.goto(source_url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)

            # 等页面里出现<a>标签(尽力等待,超时不报错,后面还有dated_item_count判定兜底)
            try:
                page.wait_for_selector("a", state="attached", timeout=TABLE_WAIT_TIMEOUT)
            except Exception:
                pass
            page.wait_for_timeout(800)  # 出现<a>后给Ajax数据渲染留缓冲

            raw_items = extract_raw_items(page)
            dated_item_count, window_titles, records = parse_records_from_raw_items(
                raw_items, code, fund_name, source_url, start_date, end_date
            )

            # 公告数据是否真的加载了（不限定table还是ul/li结构）
            if dated_item_count == 0:
                if attempt < MAX_ATTEMPTS:
                    logger.info("  第%d次页面异常（未解析到任何带日期的公告条目），%d秒后重试...",
                                attempt, RETRY_WAIT_SECONDS)
                    time.sleep(RETRY_WAIT_SECONDS)
                    continue
                save_debug_snapshot(page, code, "tablefail")
                logger.info("没有找到额度相关公告：%s %s（连续%d次页面异常，已存debug截图）",
                            code, fund_name, MAX_ATTEMPTS)
                return []

            # 第1页正常:翻页继续抓(最多 MAX_PAGES 页,翻不动就停)
            for _page_no in range(2, MAX_PAGES + 1):
                try:
                    next_btn = page.locator("a", has_text="下一页")
                    if next_btn.count() == 0:
                        break
                    next_btn.first.click()
                    page.wait_for_timeout(1200)  # 翻页后等Ajax刷新

                    more_raw_items = extract_raw_items(page)
                    _more_dated, more_window_titles, more_records = parse_records_from_raw_items(
                        more_raw_items, code, fund_name, source_url, start_date, end_date
                    )
                    records.extend(more_records)
                    window_titles.extend(more_window_titles)
                except Exception:
                    break

            # 翻页可能造成跨页重复,去重
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
                # 有日期条目但没命中关键词：打印窗口内部分标题用于诊断（限量）
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
            if attempt < MAX_ATTEMPTS:
                logger.info("  查询异常(%s)，%d秒后重试...", error, RETRY_WAIT_SECONDS)
                time.sleep(RETRY_WAIT_SECONDS)
            else:
                logger.info("查询失败：%s %s", code, error)
    return []


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
    start_date, end_date = fetch_window_dates(PAST_DAYS, FUTURE_DAYS)
    logger.info("本次检查窗口：%s ~ %s（过去%d天 + 未来%d天）",
                start_date, end_date, PAST_DAYS, FUTURE_DAYS)

    fund_dict = load_fund_dict()
    seen = load_seen_records()

    all_records = []
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled",
                  "--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage"],
        )
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

    # 只有推送成功才写记录(事务性)
    seen.update(r["_key"] for r in new_records)
    save_seen_records(seen)


if __name__ == "__main__":
    main()
