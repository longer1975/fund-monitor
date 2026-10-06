# -*- coding: utf-8 -*-
"""
QDII 基金定期报告监控
- 监控：一季报、半年报/中报、三季报、年报
- 数据源：AKShare -> 东方财富基金公告
- 首次见到某只基金时只建立历史基线，不推送旧报告
- 后续出现新定期报告时，通过飞书机器人立即通知
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import akshare as ak
import pandas as pd
import requests

APP_DIR = Path(__file__).resolve().parent
FUNDS_FILE = APP_DIR / "report_funds.json"
STATE_FILE = APP_DIR / "report_state.json"

FEISHU_WEBHOOK_URL = os.environ.get("FEISHU_WEBHOOK_URL", "").strip()
FEISHU_KEYWORD = os.environ.get("REPORT_FEISHU_KEYWORD", "基金公告").strip()
NOTIFY_ON_FIRST_RUN = os.environ.get("REPORT_NOTIFY_ON_FIRST_RUN", "false").lower() in {
    "1", "true", "yes", "y"
}

REQUEST_RETRIES = 3
REQUEST_TIMEOUT = 20
PER_FUND_SLEEP = 0.25
MAX_STATE_KEYS = 5000


@dataclass(frozen=True)
class Report:
    fund_code: str
    fund_name: str
    report_type: str
    title: str
    date: str

    @property
    def unique_key(self) -> str:
        raw = f"{self.fund_code}|{self.title}|{self.date}"
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()


REPORT_RULES: list[tuple[str, re.Pattern[str]]] = [
    ("一季报", re.compile(r"(?:第一季度报告|第1季度报告|1季度报告|一季度报告|一季报)", re.I)),
    ("中报", re.compile(r"(?:半年度报告|中期报告|半年报|中报)", re.I)),
    ("三季报", re.compile(r"(?:第三季度报告|第3季度报告|3季度报告|三季度报告|三季报)", re.I)),
    ("年报", re.compile(r"(?:年度报告|年报)", re.I)),
]

EXCLUDE_RE = re.compile(
    r"摘要|提示性公告|披露提示|更正公告|更正说明|关于.*(?:季度报告|半年度报告|中期报告|年度报告).*更正",
    re.I,
)


def classify_report(title: str) -> str | None:
    text = re.sub(r"\s+", "", str(title or ""))
    if not text or EXCLUDE_RE.search(text):
        return None
    for report_type, pattern in REPORT_RULES:
        if pattern.search(text):
            return report_type
    return None


def load_funds() -> dict[str, str]:
    with FUNDS_FILE.open("r", encoding="utf-8") as f:
        data = json.load(f)
    return {str(code).zfill(6): str(name) for code, name in data.items()}


def load_state() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return {"initialized_funds": [], "seen": {}, "updated_at": ""}
    try:
        with STATE_FILE.open("r", encoding="utf-8") as f:
            state = json.load(f)
        state.setdefault("initialized_funds", [])
        state.setdefault("seen", {})
        state.setdefault("updated_at", "")
        return state
    except Exception as exc:
        print(f"state.json 读取失败，将按空状态处理：{exc}", file=sys.stderr)
        return {"initialized_funds": [], "seen": {}, "updated_at": ""}


def save_state(state: dict[str, Any]) -> None:
    seen = state.get("seen", {})
    if len(seen) > MAX_STATE_KEYS:
        newest = sorted(
            seen.items(),
            key=lambda item: item[1].get("seen_at", ""),
            reverse=True,
        )[:MAX_STATE_KEYS]
        state["seen"] = dict(newest)

    state["initialized_funds"] = sorted(set(state.get("initialized_funds", [])))
    state["updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")

    tmp = STATE_FILE.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
    tmp.replace(STATE_FILE)


def fetch_announcements(code: str) -> pd.DataFrame:
    last_error: Exception | None = None
    for attempt in range(1, REQUEST_RETRIES + 1):
        try:
            df = ak.fund_announcement_report_em(symbol=code)
            if df is None or df.empty:
                raise RuntimeError("公告接口返回空数据")
            if "公告标题" not in df.columns:
                raise RuntimeError(f"公告接口字段异常：{df.columns.tolist()}")
            return df
        except Exception as exc:
            last_error = exc
            if attempt < REQUEST_RETRIES:
                wait = min(2 ** (attempt - 1), 8)
                print(f"  第 {attempt} 次失败，{wait}s 后重试：{exc}")
                time.sleep(wait)
    raise RuntimeError(f"查询失败，已重试 {REQUEST_RETRIES} 次：{last_error}")


def normalize_date(value: Any) -> str:
    dt = pd.to_datetime(value, errors="coerce")
    if pd.notna(dt):
        return dt.strftime("%Y-%m-%d")
    return str(value or "").strip()


def extract_reports(code: str, name: str, df: pd.DataFrame) -> list[Report]:
    reports: list[Report] = []
    for _, row in df.iterrows():
        title = str(row.get("公告标题", "")).strip()
        report_type = classify_report(title)
        if not report_type:
            continue
        reports.append(
            Report(
                fund_code=code,
                fund_name=name,
                report_type=report_type,
                title=title,
                date=normalize_date(row.get("公告日期", "")),
            )
        )
    reports.sort(key=lambda x: (x.date, x.title), reverse=True)
    return reports


def send_feishu(text: str) -> None:
    if not FEISHU_WEBHOOK_URL:
        raise RuntimeError("未配置 FEISHU_WEBHOOK_URL")

    prefix = f"【{FEISHU_KEYWORD}】\n" if FEISHU_KEYWORD else ""
    payload = {
        "msg_type": "text",
        "content": {"text": prefix + text},
    }
    resp = requests.post(FEISHU_WEBHOOK_URL, json=payload, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()

    try:
        data = resp.json()
    except ValueError:
        return

    code = data.get("code", data.get("StatusCode", 0))
    if code not in (0, "0", None):
        raise RuntimeError(f"飞书返回失败：{data}")


def format_notice(report: Report) -> str:
    return "\n".join(
        [
            "发现新的基金定期报告",
            f"基金：{report.fund_name}（{report.fund_code}）",
            f"类型：{report.report_type}",
            f"日期：{report.date or '未知'}",
            f"公告：{report.title}",
        ]
    )


def scan_once(test_feishu: bool = False) -> int:
    if test_feishu:
        send_feishu("测试成功：基金季报/中报/年报监控已连接飞书。")
        print("飞书测试消息发送成功")
        return 0

    funds = load_funds()
    state = load_state()
    initialized = {str(x).zfill(6) for x in state.get("initialized_funds", [])}
    seen: dict[str, Any] = state.setdefault("seen", {})

    print(f"开始扫描 {len(funds)} 只基金；已建立基线 {len(initialized)} 只")

    failed: list[str] = []
    unseen: list[Report] = []

    for i, (code, name) in enumerate(funds.items(), start=1):
        print(f"[{i:02d}/{len(funds):02d}] {code} {name}")
        try:
            reports = extract_reports(code, name, fetch_announcements(code))
            print(f"  匹配到 {len(reports)} 条定期报告")

            if code not in initialized and not NOTIFY_ON_FIRST_RUN:
                now = datetime.now().astimezone().isoformat(timespec="seconds")
                for report in reports:
                    seen[report.unique_key] = {**asdict(report), "seen_at": now}
                initialized.add(code)
                state["initialized_funds"] = sorted(initialized)
                save_state(state)
                print(f"  首次建立基线：记录 {len(reports)} 条历史报告，不通知")
                continue

            if code not in initialized:
                initialized.add(code)
                state["initialized_funds"] = sorted(initialized)

            for report in reports:
                if report.unique_key not in seen:
                    unseen.append(report)

        except Exception as exc:
            failed.append(f"{code} {name}: {exc}")
            print(f"  查询失败：{exc}", file=sys.stderr)

        time.sleep(PER_FUND_SLEEP)

    if failed and len(failed) >= max(3, len(funds) // 2):
        state["initialized_funds"] = sorted(initialized)
        save_state(state)
        print("本次大量基金查询失败，为避免误报，本轮停止发送。", file=sys.stderr)
        for item in failed:
            print(" -", item, file=sys.stderr)
        return 2

    unseen.sort(key=lambda x: (x.date, x.fund_code, x.title))
    print(f"发现未通知的新报告：{len(unseen)} 条")

    sent = 0
    for report in unseen:
        try:
            send_feishu(format_notice(report))
            sent += 1
            seen[report.unique_key] = {
                **asdict(report),
                "seen_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            }
            state["initialized_funds"] = sorted(initialized)
            save_state(state)
            print(f"  已通知：{report.fund_code} | {report.report_type} | {report.title}")
        except Exception as exc:
            print(f"  飞书发送失败，下次继续重试：{exc}", file=sys.stderr)

    state["initialized_funds"] = sorted(initialized)
    save_state(state)

    if failed:
        print(f"本次有 {len(failed)} 只基金查询失败，其余已完成。")
        for item in failed:
            print(" -", item)

    print(f"扫描完成：新报告 {len(unseen)}，成功通知 {sent}")
    return 0


if __name__ == "__main__":
    test = "--test-feishu" in sys.argv
    raise SystemExit(scan_once(test_feishu=test))
