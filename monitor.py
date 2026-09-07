# -*- coding: utf-8 -*-
"""
基金限额公告监控（GitHub Actions 单轮模式版）
由 workflow 每10分钟触发一次，每次只执行一轮查询
v2：增加页面加载智能等待、抓取重试机制、基金间隔延迟，修复偶发漏抓
"""

import json
import os
import re
import time
from copy import copy
from datetime import date, datetime, timedelta

import pandas as pd
import requests
from playwright.sync_api import sync_playwright


# ============================================================
# 一、基金清单
# ============================================================

fund_dict = {
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
# 二、运行参数
# ============================================================

# 查询最近几个自然日
RECENT_DAYS = 2

# 允许公告日期晚于系统日期几天
ALLOW_FUTURE_DAYS = 4

# 页面最长等待时间，单位：毫秒
PAGE_TIMEOUT = 6000

# ★ v2修复：页面加载完成后等待动态内容，从300毫秒提高到1500毫秒
#   东财公告列表是 Ajax 动态注入，300毫秒经常等不到公告渲染完成
WAIT_AFTER_LOAD = 1500

# 单只基金最大尝试次数（抓取失败自动重试）
MAX_ATTEMPTS = 3

# 重试间隔，单位：秒
RETRY_WAIT_SECONDS = 2

# 是否后台运行
HEADLESS = True

# 第一次运行是否发送当前已有公告
# False：第一次只建立历史记录，不发送旧公告
# True：第一次也发送当前公告
FIRST_RUN_NOTIFY = False

# 已发送公告记录（GitHub Actions 中随仓库提交持久化）
SEEN_FILE = "seen_announcements.json"

# Excel文件（运行时生成，由 workflow 上传为 artifact）
EXCEL_FILE = "近5天限额公告.xlsx"


# ============================================================
# 三、飞书配置
# webhook 从环境变量读取，不硬编码到代码里
# （GitHub 仓库代码公开可见，Secrets 才是安全存放位置）
# ============================================================

FEISHU_WEBHOOK_URL = os.environ.get("FEISHU_WEBHOOK_URL", "")


# ============================================================
# 四、公告关键词
# ============================================================

STRICT_LIMIT_KEYWORDS = [
    "大额申购", "大额购买", "暂停大额申购", "暂停大额购买",
    "恢复大额申购", "恢复大额购买", "限制大额申购", "限制大额购买",
    "申购上限", "购买上限", "业务上限", "业务限额", "金额限制",
    "限制金额", "额度上限", "额度调整", "调整额度", "调整限额",
    "调整大额", "调整申购", "调整定投", "调整定期定额",
    "暂停定投", "恢复定投", "暂停定期定额", "恢复定期定额",
    "定期定额投资业务", "定期定额申购业务", "定期定额业务",
    "定投业务", "不定额投资", "转换转入", "申购业务限制",
    "申购业务上限", "申购业务限额",
]

EXCLUDE_TITLE_KEYWORDS = [
    "节假日", "市场节假日", "暂停申购赎回安排", "暂停赎回安排",
    "申购赎回安排", "开放日安排", "交易日安排", "清明节",
    "劳动节", "端午节", "中秋节", "国庆节", "春节", "元旦",
]


# ============================================================
# 五、基础文字函数
# ============================================================

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
    """从文字中提取日期"""
    text = clean_text(text)
    match = re.search(r"(20\d{2})[-年/.](\d{1,2})[-月/.](\d{1,2})", text)
    if not match:
        return None
    try:
        return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except Exception:
        return None


def format_date(value):
    """格式化日期"""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def is_valid_title(title):
    """判断是否为有效公告标题"""
    title = normalize_title(title)
    if len(title) < 15:
        return False
    invalid_titles = {
        "首页", "上一页", "下一页", "尾页", "返回", "更多",
        "公告", "详情", "查看详情", "点击查看", "基金销售",
        "公告标题", "公告日期",
    }
    if title in invalid_titles:
        return False
    if parse_date(title) is not None:
        return False
    return True


# ============================================================
# 六、严格判断限额公告
# ============================================================

def is_limit_related_title(title):
    """只识别真正的限额、申购调整、定投调整公告"""
    title = normalize_title(title)
    if not title:
        return False
    for keyword in EXCLUDE_TITLE_KEYWORDS:
        if keyword in title:
            return False
    for keyword in STRICT_LIMIT_KEYWORDS:
        if keyword in title:
            return True
    return False


# ============================================================
# 七、拦截无关资源
# ============================================================

def block_unnecessary_resources(route):
    """只拦截图片、字体、视频等资源，不拦截脚本/XHR/Fetch"""
    request = route.request
    resource_type = request.resource_type
    url = request.url.lower()

    if resource_type in {"image", "font", "media"}:
        route.abort()
        return

    blocked_words = [
        "google-analytics", "hm.baidu.com", "sensorsdata",
        "doubleclick", "googlesyndication", "adservice", "advert",
    ]
    if any(keyword in url for keyword in blocked_words):
        route.abort()
        return

    try:
        route.continue_()
    except Exception:
        pass


# ============================================================
# 八、快速提取公告
# ============================================================

def extract_page_records(page, code, fund_name, source_url, start_date, end_date):
    """快速提取公告：标题只读a标签本身，日期从最近公告行读取"""
    try:
        # 使用 r""" 原始字符串，保证正则原样传给 JS
        raw_items = page.evaluate(
            r"""
            () => {
                const result = [];
                const links = Array.from(document.querySelectorAll("a"));
                const clean = (value) => {
                    return String(value || "")
                        .replace(/[\s\u3000]+/g, "")
                        .trim();
                };
                for (const link of links) {
                    const title = clean(link.innerText || link.textContent);
                    if (!title) continue;
                    let row = link.closest("tr");
                    if (!row) row = link.closest("li");
                    if (!row) row = link.closest(".list-item, .notice-item, .item");
                    if (!row) row = link.parentElement;
                    if (!row) row = link;
                    const rowText = clean(row.innerText || row.textContent);
                    result.push({ title: title, rowText: rowText });
                }
                return result;
            }
            """
        )
    except Exception as error:
        print(f"      页面公告提取失败：{error}")
        return []

    records = []
    for item in raw_items:
        title = normalize_title(item.get("title", ""))
        row_text = clean_text(item.get("rowText", ""))

        if not is_valid_title(title):
            continue
        if not is_limit_related_title(title):
            continue

        announcement_date = parse_date(row_text)
        if announcement_date is None:
            announcement_date = parse_date(title)
        if announcement_date is None:
            continue
        if announcement_date < start_date:
            continue
        if announcement_date > end_date:
            continue

        records.append({
            "基金代码": code,
            "基金名称": fund_name,
            "公告标题": title,
            "公告日期": format_date(announcement_date),
            "公告日期对象": announcement_date,
            "公告页面": source_url,
        })

    # 精确去重
    unique_records = {}
    for record in records:
        key = (record["基金代码"], record["公告日期"], record["公告标题"])
        unique_records[key] = record
    records = list(unique_records.values())

    records.sort(key=lambda item: (item["公告日期对象"], item["公告标题"]), reverse=True)
    return records


# ============================================================
# 九、查询单只基金
# ★ v2修复：带重试机制。页面没加载完或被风控导致抓取失败时，
#   自动重试最多 MAX_ATTEMPTS 次，而不是静默当"没有公告"
# ============================================================

def query_one_fund(page, index, total, code, fund_name, start_date, end_date):
    """查询一只基金（带重试机制）"""
    code = str(code).zfill(6)
    source_url = f"https://fundf10.eastmoney.com/jjgg_{code}.html"

    print(f"[{index:02d}/{total}] 正在查询：{code} {fund_name}")

    records = []

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            page.goto(source_url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)

            # 智能等待：等页面 a 标签出现（最多再等3秒）
            try:
                page.wait_for_selector("a", state="attached", timeout=3000)
            except Exception:
                pass

            # 额外固定等待，让 Ajax 有时间注入公告列表
            page.wait_for_timeout(WAIT_AFTER_LOAD)

            records = extract_page_records(
                page=page, code=code, fund_name=fund_name,
                source_url=source_url, start_date=start_date, end_date=end_date,
            )

            if records:
                print(f"找到额度相关公告：{code} {fund_name}")
                for record in records:
                    print(f"- {record['公告日期']} {record['公告标题']}")
                return records

            # 没抓到 → 重试
            if attempt < MAX_ATTEMPTS:
                print(f"  第{attempt}次未抓到，{RETRY_WAIT_SECONDS}秒后重试...")
                time.sleep(RETRY_WAIT_SECONDS)
            else:
                print(f"没有找到额度相关公告：{code} {fund_name}")

        except Exception as error:
            if attempt < MAX_ATTEMPTS:
                print(f"  查询异常({error})，{RETRY_WAIT_SECONDS}秒后重试...")
                time.sleep(RETRY_WAIT_SECONDS)
            else:
                print(f"查询失败：{code} {error}")

    return records


# ============================================================
# 十、查询全部基金
# ★ v2修复：基金之间加随机延迟，降低被东财风控的概率
# ============================================================

def query_all_funds(browser):
    """查询全部基金公告"""
    today = date.today()
    start_date = today - timedelta(days=RECENT_DAYS - 1)
    end_date = today + timedelta(days=ALLOW_FUTURE_DAYS)

    print()
    print("=" * 80)
    print(f"系统日期：{format_date(today)}")
    print(f"公告筛选范围：{format_date(start_date)} 至 {format_date(end_date)}")
    print("=" * 80)

    context = browser.new_context(
        viewport={"width": 1366, "height": 900},
        locale="zh-CN",
        timezone_id="Asia/Shanghai",
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0 Safari/537.36"
        ),
    )
    context.set_default_timeout(1000)
    context.route("**/*", block_unnecessary_resources)
    page = context.new_page()

    results = []
    total = len(fund_dict)

    try:
        for index, (code, fund_name) in enumerate(fund_dict.items(), start=1):
            records = query_one_fund(
                page=page, index=index, total=total, code=code,
                fund_name=fund_name, start_date=start_date, end_date=end_date,
            )
            results.extend(records)
            # 基金之间随机延迟 0.8~1.2 秒，模拟人工浏览节奏
            time.sleep(0.8 + (index % 3) * 0.2)
    finally:
        try:
            page.close()
        except Exception:
            pass
        try:
            context.close()
        except Exception:
            pass

    # 最终去重
    unique_results = {}
    for record in results:
        key = (record["基金代码"], record["公告日期"], record["公告标题"])
        unique_results[key] = record
    return list(unique_results.values())


# ============================================================
# 十一、导出Excel
# ============================================================

def export_excel(records):
    """导出当前查询结果"""
    columns = ["基金代码", "基金名称", "公告标题", "公告日期"]
    rows = []
    for record in records:
        rows.append({
            "基金代码": record["基金代码"],
            "基金名称": record["基金名称"],
            "公告标题": record["公告标题"],
            "公告日期": record["公告日期"],
        })

    result_df = pd.DataFrame(rows, columns=columns)

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="近5天限额公告")
        worksheet = writer.sheets["近5天限额公告"]
        worksheet.column_dimensions["A"].width = 14
        worksheet.column_dimensions["B"].width = 34
        worksheet.column_dimensions["C"].width = 110
        worksheet.column_dimensions["D"].width = 15
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        for cell in worksheet[1]:
            font = copy(cell.font)
            font.bold = True
            cell.font = font
            alignment = copy(cell.alignment)
            alignment.horizontal = "center"
            alignment.vertical = "center"
            alignment.wrap_text = True
            cell.alignment = alignment

        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                alignment = copy(cell.alignment)
                alignment.vertical = "center"
                alignment.wrap_text = True
                cell.alignment = alignment

    print(f"Excel已保存：{os.path.abspath(EXCEL_FILE)}")


# ============================================================
# 十二、历史公告
# ============================================================

def make_key(record):
    """生成公告唯一编号"""
    return "|".join([record["基金代码"], record["公告日期"], record["公告标题"]])


def load_seen_records():
    """读取历史公告"""
    if not os.path.exists(SEEN_FILE):
        return set()
    try:
        with open(SEEN_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
        if not isinstance(data, list):
            return set()
        return set(str(item) for item in data)
    except Exception as error:
        print(f"读取历史记录失败：{error}")
        return set()


def save_seen_records(seen_records):
    """保存历史公告"""
    with open(SEEN_FILE, "w", encoding="utf-8") as file:
        json.dump(sorted(list(seen_records)), file, ensure_ascii=False, indent=2)


# ============================================================
# 十三、飞书通知
# ============================================================

def send_feishu_message(records):
    """发送飞书文本消息"""
    if not records:
        return True

    if not FEISHU_WEBHOOK_URL:
        print("没有配置飞书Webhook环境变量，本次不发送消息")
        return False

    message_lines = [
        "📢 基金限额公告提醒",
        "",
        f"发现新公告：{len(records)}条",
        f"通知时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]

    for index, record in enumerate(records, start=1):
        message_lines.append(f"{index}. {record['基金名称']}")
        message_lines.append(f"基金代码：{record['基金代码']}")
        message_lines.append(f"公告日期：{record['公告日期']}")
        message_lines.append(f"公告标题：{record['公告标题']}")
        message_lines.append(f"公告页面：{record['公告页面']}")
        message_lines.append("")

    payload = {"msg_type": "text", "content": {"text": "\n".join(message_lines)}}

    try:
        response = requests.post(
            FEISHU_WEBHOOK_URL,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=20,
        )
        response.raise_for_status()
        result = response.json()
        if result.get("code", 0) != 0:
            print(f"飞书返回失败：{result}")
            return False
        print("飞书消息发送成功")
        return True
    except Exception as error:
        print(f"飞书发送失败：{error}")
        return False


# ============================================================
# 十四、执行一轮
# ============================================================

def run_one_round(browser, seen_records, first_run):
    """执行一轮抓取和通知"""
    print()
    print("=" * 80)
    print(f"开始查询：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    records = query_all_funds(browser)

    if records:
        export_excel(records)

    current_keys = {make_key(record) for record in records}

    # 第一次运行只建立基准
    if first_run and not FIRST_RUN_NOTIFY:
        print("第一次运行：当前公告只写入历史记录，不发送飞书")
        seen_records.update(current_keys)
        save_seen_records(seen_records)
        return seen_records

    new_records = [record for record in records if make_key(record) not in seen_records]

    if not new_records:
        print("本轮没有新的限额公告")
        seen_records.update(current_keys)
        save_seen_records(seen_records)
        return seen_records

    print()
    print(f"发现新的限额公告：{len(new_records)}条")
    for record in new_records:
        print(f"- {record['基金代码']} {record['公告日期']} {record['公告标题']}")

    success = send_feishu_message(new_records)

    if success:
        for record in new_records:
            seen_records.add(make_key(record))
        save_seen_records(seen_records)
    else:
        print("飞书发送失败，本次不写入历史，下轮将继续重试")

    return seen_records


# ============================================================
# 十五、主程序
# 单轮模式：GitHub Actions 的 cron 定时器负责每10分钟触发一次，
# 每次运行只执行一轮，跑完即退出
# ============================================================

def main():
    seen_records = load_seen_records()
    first_run = not os.path.exists(SEEN_FILE)

    print("=" * 80)
    print("基金限额公告监控（单轮模式 + 重试机制，由 GitHub Actions 定时触发）")
    print("=" * 80)
    print(f"基金数量：{len(fund_dict)}只")
    print(f"查询范围：最近{RECENT_DAYS}个自然日")
    print(f"未来日期容许：{ALLOW_FUTURE_DAYS}天")
    print(f"单基金最大尝试：{MAX_ATTEMPTS}次")
    print(f"历史公告数量：{len(seen_records)}条")
    print(f"是否首次运行：{first_run}")
    print("=" * 80)

    with sync_playwright() as playwright:
        browser = None
        try:
            browser = playwright.chromium.launch(
                headless=HEADLESS,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-gpu",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )
            run_one_round(browser, seen_records, first_run)
        finally:
            if browser is not None:
                try:
                    browser.close()
                except Exception:
                    pass


# ============================================================
# 十六、程序入口
# ============================================================

if __name__ == "__main__":
    main()
