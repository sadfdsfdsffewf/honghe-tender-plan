# -*- coding: utf-8 -*-
"""红河州招标计划每日工作流。

功能：
1. 从云南省公共资源交易平台增量抓取红河州招标计划；
2. 维护有效项目库，自动删除已超过“预计招标公告发布时间”的项目；
3. 生成 Excel、CSV、离线单文件查询页及日报；
4. 适合由 Codex 每日 15:00 自动化任务调用。
"""
from __future__ import annotations

import argparse
import csv
import html as htmllib
import json
import logging
import os
import re
import time
import urllib.request
from urllib.parse import quote
import zipfile
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
CACHE_DIR = BASE_DIR / "_cache_plan"
ZBGG_CACHE_DIR = BASE_DIR / "_cache_zbgg"
LOG_DIR = BASE_DIR / "logs"
for _d in (DATA_DIR, OUTPUT_DIR, CACHE_DIR, ZBGG_CACHE_DIR, LOG_DIR):
    _d.mkdir(parents=True, exist_ok=True)

API_BASE = "https://ggzy.yn.gov.cn/ynggfwpt-home-api"
PORTAL_URL = "https://ggzy.yn.gov.cn/ynggfwpt-home-web/"
PORTAL_DETAIL_URL = "https://ggzy.yn.gov.cn/ynggfwpt-home-web/#/tradeHall/tradeDetail"
CITY_ID = "015"                  # 红河州在云南站内的筛选码
REGION_PREFIX = "5325"           # 红河州行政区划代码前缀
BOOTSTRAP_DAYS = 180             # 首次运行回看近半年
OVERLAP_DAYS = 2                 # 每次增量多回看 2 天，捕获更正
ANNOUNCEMENT_RETENTION_DAYS = 30  # 微信公告页保留最近 30 天
REQUEST_INTERVAL = 0.38
HTTP_TIMEOUT = 30
TZ_CN = timezone(timedelta(hours=8))

STATE_FILE = DATA_DIR / "active_plans.json"
ANNOUNCEMENT_STATE_FILE = DATA_DIR / "active_announcements.json"
SUMMARY_FILE = DATA_DIR / "latest_summary.json"
ANNOUNCEMENT_SUMMARY_FILE = DATA_DIR / "latest_announcement_summary.json"
ACTIVE_XLSX = OUTPUT_DIR / "红河州招标计划_有效项目.xlsx"
TODAY_XLSX = OUTPUT_DIR / "红河州招标计划_今日更新.xlsx"
ACTIVE_CSV = OUTPUT_DIR / "红河州招标计划_有效项目.csv"
QUERY_HTML = OUTPUT_DIR / "红河州招标计划_查询.html"
WECHAT_HTML = OUTPUT_DIR / "红河州招标计划_微信版.html"
PUBLISH_DIR = OUTPUT_DIR / "publish"
PUBLISH_INDEX = PUBLISH_DIR / "index.html"
REPORT_MD = OUTPUT_DIR / "最新日报.md"
PORTABLE_ZIP = OUTPUT_DIR / "红河州招标计划_便携查询包.zip"
PUBLIC_ZIP = OUTPUT_DIR / "红河州招标计划_公网部署包.zip"
PUBLISH_DIR.mkdir(parents=True, exist_ok=True)

HDRS = {
    "Content-Type": "application/json;charset=UTF-8",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Referer": PORTAL_URL,
    "Origin": "https://ggzy.yn.gov.cn",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9",
}

COLUMNS = [
    ("招标项目名称", 46),
    ("标段合同估算价（万元）", 20),
    ("招标人", 32),
    ("项目行政主管地区", 22),
    ("入场登记部门", 28),
    ("预计招标公告发布时间", 21),
    ("招标人联系人", 14),
    ("招标人联系电话", 19),
    ("计划性质", 10),
    ("发布时间", 20),
    ("交易平台", 14),
    ("招标内容", 60),
    ("备注", 72),
]
COLUMN_NAMES = [x[0] for x in COLUMNS]
HONGHE_KEYWORDS = [
    "红河", "蒙自", "个旧", "开远", "弥勒", "建水", "石屏", "泸西",
    "元阳", "绿春", "屏边", "金平", "河口",
]


def now_cn() -> datetime:
    return datetime.now(TZ_CN)


def log_setup() -> logging.Logger:
    logger = logging.getLogger("honghe_plan")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    stream = logging.StreamHandler()
    stream.setFormatter(fmt)
    logger.addHandler(stream)
    file_handler = logging.FileHandler(
        LOG_DIR / f"run_{now_cn():%Y%m%d}.log", encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    logger.addHandler(file_handler)
    return logger


LOG = log_setup()


def load_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        LOG.warning("读取 %s 失败，将使用默认值: %s", path.name, exc)
        return default


def atomic_write_json(path: Path, value: Any) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def atomic_write_text(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def api_post(path: str, payload: dict[str, Any], retries: int = 4) -> dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    for attempt in range(retries):
        try:
            req = urllib.request.Request(
                API_BASE + path, data=body, headers=HDRS, method="POST"
            )
            with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
                return json.loads(resp.read().decode("utf-8", "replace"))
        except Exception:
            if attempt == retries - 1:
                raise
            time.sleep(2.5 * (attempt + 1))
    raise RuntimeError("unreachable")


def api_get(path: str, retries: int = 4) -> dict[str, Any]:
    for attempt in range(retries):
        try:
            req = urllib.request.Request(API_BASE + path, headers=HDRS)
            with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
                return json.loads(resp.read().decode("utf-8", "replace"))
        except Exception:
            if attempt == retries - 1:
                raise
            time.sleep(2.5 * (attempt + 1))
    raise RuntimeError("unreachable")


def trade_url(guid: Any, row_code: str) -> str:
    guid_text = str(guid or "").strip()
    if not guid_text:
        return ""
    return (
        f"{PORTAL_DETAIL_URL}?guid={quote(guid_text)}"
        f"&colCode=1&rowCode={quote(row_code)}&from=home"
    )


def project_url(guid: Any) -> str:
    return trade_url(guid, "招标计划")


def announcement_url(guid: Any) -> str:
    return trade_url(guid, "招标公告")


def fmt_api_time(dt: datetime) -> str:
    return dt.astimezone(TZ_CN).strftime("%Y%m%d%H%M%S")


def fetch_list_window(start: datetime, end: datetime) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    page = 1
    pages = 1
    while page <= pages:
        payload = {
            "pageNum": page,
            "pageSize": 200,
            "cityId": CITY_ID,
            "tradeType": "gcjs",
            "bulletinNature": "zbjh",
            "childType": "",
            "title": "",
            "startTime": fmt_api_time(start),
            "endTime": fmt_api_time(end),
        }
        data = api_post("/jyzyCenter/jyInfo/gcjs/getTenserPlanList", payload)
        if str(data.get("code")) != "1":
            raise RuntimeError(f"列表接口异常: {data}")
        value = data.get("value") or {}
        pages = int(value.get("pages") or 0)
        batch = value.get("list") or []
        items.extend(batch)
        LOG.info("列表第 %s/%s 页：%s 条", page, pages or 1, len(batch))
        page += 1
        if page <= pages:
            time.sleep(REQUEST_INTERVAL)
    return items


def fetch_detail(guid: str, max_cache_hours: float = 12.0) -> dict[str, Any]:
    cache_file = CACHE_DIR / f"{guid}.json"
    if cache_file.exists():
        age_hours = (time.time() - cache_file.stat().st_mtime) / 3600
        if age_hours <= max_cache_hours:
            try:
                return json.loads(cache_file.read_text(encoding="utf-8"))
            except Exception:
                pass
    data = api_get(f"/jyzyCenter/jyInfo/gcjs/getTenserPlanDetail?guid={guid}")
    value = data.get("value") or {}
    if not value:
        raise RuntimeError(f"详情为空: {guid}")
    tmp = cache_file.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, cache_file)
    return value


def fetch_announcement_list_window(start: datetime, end: datetime) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    page = 1
    pages = 1
    while page <= pages:
        payload = {
            "pageNum": page,
            "pageSize": 200,
            "cityId": CITY_ID,
            "tradeType": "gcjs",
            "bulletinNature": "招标公告",
            "childType": "",
            "title": "",
            "startTime": fmt_api_time(start),
            "endTime": fmt_api_time(end),
        }
        data = api_post("/jyzyCenter/jyInfo/gcjs/getZbggList", payload)
        if str(data.get("code")) != "1":
            raise RuntimeError(f"招标公告列表接口异常: {data}")
        value = data.get("value") or {}
        pages = int(value.get("pages") or 0)
        batch = value.get("list") or []
        items.extend(batch)
        LOG.info("公告列表第 %s/%s 页：%s 条", page, pages or 1, len(batch))
        page += 1
        if page <= pages:
            time.sleep(REQUEST_INTERVAL)
    return items


def fetch_announcement_detail(guid: str, max_cache_hours: float = 12.0) -> dict[str, Any]:
    cache_file = ZBGG_CACHE_DIR / f"{guid}.json"
    if cache_file.exists():
        age_hours = (time.time() - cache_file.stat().st_mtime) / 3600
        if age_hours <= max_cache_hours:
            try:
                return json.loads(cache_file.read_text(encoding="utf-8"))
            except Exception:
                pass
    data = api_get(f"/jyzyCenter/jyInfo/gcjs/findZbggByGuid?guid={guid}")
    value = data.get("value") or {}
    if not value:
        raise RuntimeError(f"招标公告详情为空: {guid}")
    tmp = cache_file.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(value, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, cache_file)
    return value


def strip_hidden(value: str) -> str:
    if not value:
        return ""
    for _ in range(3):
        changed = re.sub(
            r"(?is)<span[^>]*style\s*=\s*['\"][^'\"]*display\s*:\s*none"
            r"[^'\"]*['\"][^>]*>.*?</span>",
            " ",
            value,
        )
        if changed == value:
            break
        value = changed
    return value


def clean_html(value: Any) -> str:
    if value is None:
        return ""
    text = strip_hidden(str(value))
    text = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(p|div|h[1-6]|tr|li)>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = htmllib.unescape(text).replace("\u00a0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text.strip()


def split_sections(source: str) -> dict[str, str]:
    out: dict[str, str] = {}
    parts = re.split(r"(?is)<h3[^>]*>(.*?)</h3>", source or "")
    for idx in range(1, len(parts) - 1, 2):
        out[clean_html(parts[idx])] = parts[idx + 1]
    return out


def parse_kv(body: str) -> dict[str, str]:
    pairs: dict[str, str] = {}
    for tr in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", body or ""):
        cells = [
            clean_html(x)
            for x in re.findall(r"(?is)<t[dh][^>]*>(.*?)</t[dh]>", tr)
        ]
        idx = 0
        while idx < len(cells) - 1:
            label = cells[idx]
            if label.endswith("：") or label.endswith(":"):
                pairs[label.rstrip("：:")] = cells[idx + 1]
                idx += 2
            else:
                idx += 1
    return pairs

def parse_kv_pairs(body: str) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    for tr in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", body or ""):
        cells = [
            clean_html(x)
            for x in re.findall(r"(?is)<t[dh][^>]*>(.*?)</t[dh]>", tr)
        ]
        idx = 0
        while idx < len(cells) - 1:
            label = cells[idx]
            if label.endswith("：") or label.endswith(":"):
                pairs.append((label.rstrip("：:"), cells[idx + 1]))
                idx += 2
            else:
                idx += 1
    return pairs


def parse_table_rows(body: str) -> list[list[str]]:
    rows = []
    for tr in re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", body or ""):
        cells = [
            clean_html(x)
            for x in re.findall(r"(?is)<t[dh][^>]*>(.*?)</t[dh]>", tr)
        ]
        if any(cells):
            rows.append(cells)
    return rows


def parse_segments(body: str, project_name: str, fallback_amount: Any) -> list[dict[str, Any]]:
    rows = parse_table_rows(body)
    if len(rows) <= 1:
        return [{
            "序号": "",
            "标段名称": project_name,
            "标段合同估算价（万元）": fallback_amount,
        }]
    header = rows[0]
    result = []
    for row in rows[1:]:
        record = {}
        for idx, value in enumerate(row):
            key = header[idx] if idx < len(header) else f"列{idx + 1}"
            record[key] = value
        if record:
            result.append(record)
    return result


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    text = str(value).replace(",", "").replace("万元", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def segment_amount(segment: dict[str, Any]) -> float | None:
    for key, value in segment.items():
        if "估算价" in key or "金额" in key:
            return to_number(value)
    return None


def fmt_compact_time(value: Any) -> str:
    text = str(value or "").strip()
    if re.fullmatch(r"\d{14}", text):
        return f"{text[:4]}-{text[4:6]}-{text[6:8]} {text[8:10]}:{text[10:12]}:{text[12:14]}"
    if re.fullmatch(r"\d{8}", text):
        return f"{text[:4]}-{text[4:6]}-{text[6:8]}"
    return text


def parse_deadline(raw: Any) -> datetime | None:
    """解析预计发布时间。日期无时间或时间为 00:00 时，按当天结束处理。"""
    text = str(raw or "").strip()
    if not text:
        return None
    text = text.replace("年", "-").replace("月", "-").replace("日", " ")
    compact = re.fullmatch(r"(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?(\d{2})?", text)
    if compact:
        y, m, d, hh, mm, ss = compact.groups()
        try:
            dt = datetime(
                int(y), int(m), int(d), int(hh or 0), int(mm or 0), int(ss or 0),
                tzinfo=TZ_CN,
            )
        except ValueError:
            return None
        if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
            return dt.replace(hour=23, minute=59, second=59)
        return dt
    match = re.search(
        r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})"
        r"(?:[T\s]+(\d{1,2})[:时](\d{1,2})(?:[:分](\d{1,2}))?)?",
        text,
    )
    if not match:
        return None
    y, m, d, hh, mm, ss = match.groups()
    try:
        dt = datetime(
            int(y), int(m), int(d), int(hh or 0), int(mm or 0), int(ss or 0),
            tzinfo=TZ_CN,
        )
    except ValueError:
        return None
    if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
        return dt.replace(hour=23, minute=59, second=59)
    return dt


def iso_or_empty(dt: datetime | None) -> str:
    return dt.astimezone(TZ_CN).isoformat(timespec="seconds") if dt else ""


def is_honghe(detail: dict[str, Any], base: dict[str, Any], item: dict[str, Any]) -> bool:
    region_code = str(detail.get("regionCode") or "")
    if region_code.startswith(REGION_PREFIX):
        return True
    blob = " ".join([
        str(base.get("项目行政主管地区") or ""),
        str(item.get("jyptid") or ""),
    ])
    return any(keyword in blob for keyword in HONGHE_KEYWORDS)


def parse_record(detail: dict[str, Any], item: dict[str, Any]) -> dict[str, Any]:
    sections = split_sections(str(detail.get("content") or ""))
    base = parse_kv(sections.get("基本信息", ""))
    project_name = (
        base.get("招标项目名称")
        or detail.get("tenderProjectName")
        or item.get("tenderProjectName")
        or ""
    )
    fallback_amount = base.get("合同估算价（万元）")
    segments = parse_segments(
        sections.get("标段信息", ""), project_name, fallback_amount
    )
    amounts = [x for x in (segment_amount(s) for s in segments) if x is not None]
    total_amount = round(sum(amounts), 6) if amounts else to_number(fallback_amount)
    deadline_raw = base.get("预计招标公告发布时间", "")
    deadline = parse_deadline(deadline_raw)
    published_raw = detail.get("publishTime") or item.get("publishTime") or ""
    published_text = fmt_compact_time(published_raw)
    published_dt = None
    if re.fullmatch(r"\d{14}", str(published_raw)):
        published_dt = datetime.strptime(str(published_raw), "%Y%m%d%H%M%S").replace(tzinfo=TZ_CN)
    elif re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(published_raw)):
        published_dt = datetime.strptime(str(published_raw), "%Y-%m-%d").replace(tzinfo=TZ_CN)
    record = {
        "guid": detail.get("guid") or item.get("guid") or "",
        "plan_code": (
            detail.get("tenderProjectCode")
            or detail.get("projectCode")
            or item.get("tenderProjectCode")
            or ""
        ),
        "region_code": str(detail.get("regionCode") or ""),
        "招标项目名称": project_name,
        "标段合同估算价（万元）": total_amount,
        "招标人": base.get("招标人", ""),
        "项目行政主管地区": base.get("项目行政主管地区", ""),
        "入场登记部门": base.get("入场登记部门", ""),
        "预计招标公告发布时间": deadline_raw,
        "招标人联系人": base.get("招标人联系人", ""),
        "招标人联系电话": base.get("招标人联系电话", ""),
        "计划性质": item.get("planNatureDesc") or detail.get("planNatureDesc") or "",
        "发布时间": published_text,
        "交易平台": item.get("jyptid") or detail.get("jyptid") or "",
        "招标内容": base.get("招标内容", ""),
        "备注": project_url(detail.get("guid") or item.get("guid")),
        "deadline": iso_or_empty(deadline),
        "published_at": iso_or_empty(published_dt),
        "version": str(detail.get("version") or ""),
    }
    return record


def migrate_links(records: dict[str, dict[str, Any]]) -> int:
    changed = 0
    for record in records.values():
        url = project_url(record.get("guid"))
        if url and record.get("备注") != url:
            record["备注"] = url
            changed += 1
    return changed


def parse_announcement_record(detail: dict[str, Any], item: dict[str, Any]) -> dict[str, Any]:
    sections = split_sections(str(detail.get("bulletincontent") or ""))
    values: dict[str, str] = {}
    for key, value in parse_kv_pairs(sections.get("招标概况", "")):
        values[key] = value

    side = ""
    for key, value in parse_kv_pairs(sections.get("招标人与招标代理", "")):
        if key == "建设单位":
            side = "build"
            values["建设单位"] = value
        elif key == "招标代理机构":
            side = "agency"
            values["招标代理机构"] = value
        elif key in ("经办人", "办公电话", "移动电话"):
            prefix = "建设单位-" if side == "build" else ("代理机构-" if side == "agency" else "")
            values[prefix + key] = value
        else:
            values[key] = value

    for key, value in parse_kv_pairs(sections.get("详细公告内容", "")):
        values[key] = value

    att_section = sections.get("附件信息", "")
    skip_labels = {"序号", "附件", "附件名称", "文件名", "文件名称", "操作", "下载"}
    attachments = []
    for anchor in re.findall(r"(?is)<a[^>]*>(.*?)</a>", att_section):
        text = clean_html(anchor)
        if text and text not in skip_labels and len(text) <= 120:
            attachments.append(text)
    if not attachments:
        for key, value in parse_kv(att_section).items():
            if key not in skip_labels and value and value not in skip_labels and len(value) <= 120:
                attachments.append(value)

    guid = str(detail.get("guid") or item.get("guid") or "")
    name = (
        values.get("招标项目名称")
        or detail.get("bulletinname")
        or item.get("bulletinname")
        or ""
    )
    amount_raw = (
        values.get("标段合同估算价")
        or values.get("合同估算价（万元）")
        or values.get("合同估算价")
    )
    deadline_raw = (
        values.get("公告结束时间")
        or detail.get("bulletinendtime")
        or item.get("bulletinendtime")
        or ""
    )
    published_raw = (
        detail.get("bulletinissuetime")
        or item.get("bulletinissuetime")
        or values.get("发布时间")
        or ""
    )
    published_dt = parse_deadline(published_raw)
    doc_deadline_raw = values.get("招标文件获取截止时间") or ""
    bid_deadline_raw = values.get("递交投标文件截止时间") or ""
    bond_deadline_raw = values.get("投标保证金缴纳截止时间") or ""
    cutoff_candidates = [
        (parse_deadline(doc_deadline_raw), doc_deadline_raw),
        (parse_deadline(bid_deadline_raw), bid_deadline_raw),
        (parse_deadline(bond_deadline_raw), bond_deadline_raw),
    ]
    cutoff_candidates = [(dt, raw) for dt, raw in cutoff_candidates if dt is not None]
    cutoff_dt, cutoff_raw = min(cutoff_candidates, key=lambda pair: pair[0]) if cutoff_candidates else (None, "")
    bid_open_raw = values.get("开标时间") or values.get("开标时间：") or ""
    return {
        "guid": guid,
        "plan_code": str(detail.get("tenderprojectcode") or item.get("tenderprojectcode") or ""),
        "招标项目名称": name,
        "标段合同估算价": to_number(amount_raw),
        "资金来源": values.get("资金来源", ""),
        "建设规模": values.get("建设规模", ""),
        "工程类型": values.get("工程类型", ""),
        "资格审查方式": values.get("资格审查方式", ""),
        "建设单位": values.get("建设单位", ""),
        "建设单位办公电话": values.get("建设单位-办公电话", ""),
        "招标代理机构": values.get("招标代理机构", ""),
        "代理机构经办人": values.get("代理机构-经办人") or values.get("项目负责人") or "",
        "代理机构办公电话": values.get("代理机构-办公电话", ""),
        "代理机构移动电话": values.get("代理机构-移动电话", ""),
        "项目行政主管地区": item.get("jyptid") or detail.get("jyptid") or "",
        "交易地点": values.get("交易地点", ""),
        "发布时间": fmt_compact_time(published_raw),
        "公告结束时间": fmt_compact_time(deadline_raw),
        "公告性质": detail.get("noticeNatureDesc") or values.get("公告性质") or "",
        "联系人": values.get("建设单位-经办人") or values.get("代理机构-经办人") or "",
        "联系电话": values.get("建设单位-办公电话") or values.get("代理机构-移动电话") or values.get("代理机构-办公电话") or "",
        "招标内容": values.get("本次招标内容") or "",
        "资质要求": values.get("资质要求", ""),
        "开标时间": fmt_compact_time(bid_open_raw),
        "招标文件获取截止时间": fmt_compact_time(doc_deadline_raw),
        "递交投标文件截止时间": fmt_compact_time(bid_deadline_raw),
        "投标保证金缴纳截止时间": fmt_compact_time(bond_deadline_raw),
        "投标保证金": values.get("投标保证金金额", ""),
        "附件": "; ".join(dict.fromkeys(attachments)),
        "备注": announcement_url(guid),
        "deadline": iso_or_empty(cutoff_dt),
        "公告结束": iso_or_empty(parse_deadline(deadline_raw)),
        "cutoff_raw": fmt_compact_time(cutoff_raw),
        "published_at": iso_or_empty(published_dt),
        "version": str(detail.get("version") or ""),
    }


def record_key(record: dict[str, Any]) -> str:
    return str(record.get("plan_code") or record.get("guid") or "")


def record_signature(record: dict[str, Any]) -> str:
    important = [
        "招标项目名称", "标段合同估算价（万元）", "招标人", "项目行政主管地区",
        "入场登记部门", "预计招标公告发布时间", "招标人联系人", "招标人联系电话",
        "计划性质", "发布时间", "交易平台", "招标内容", "备注",
    ]
    return json.dumps({k: record.get(k) for k in important}, ensure_ascii=False, sort_keys=True)


def record_future(record: dict[str, Any], at: datetime) -> bool:
    deadline_text = str(record.get("deadline") or "")
    if not deadline_text:
        return True
    deadline = datetime.fromisoformat(deadline_text)
    if deadline.tzinfo is None:
        deadline = deadline.replace(tzinfo=TZ_CN)
    return deadline > at


def purge_expired(records: dict[str, dict[str, Any]], at: datetime) -> list[dict[str, Any]]:
    expired = []
    for key, record in list(records.items()):
        if not record_future(record, at):
            expired.append(record)
            records.pop(key, None)
    return expired


def upsert_records(
    records: dict[str, dict[str, Any]],
    incoming: list[dict[str, Any]],
    at: datetime,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    added, updated, expired_incoming = [], [], []
    for record in incoming:
        key = record_key(record)
        if not key:
            continue
        previous = records.get(key)
        if not record_future(record, at):
            if previous:
                records.pop(key, None)
                updated.append(record)
            else:
                expired_incoming.append(record)
            continue
        if previous is None:
            records[key] = record
            added.append(record)
        elif record_signature(previous) != record_signature(record):
            records[key] = record
            updated.append(record)
    return added, updated, expired_incoming


def money(value: Any) -> str:
    if value is None or value == "":
        return "-"
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value)


def sort_active(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    def key(record: dict[str, Any]) -> tuple[str, str]:
        deadline = str(record.get("deadline") or "9999-12-31T23:59:59+08:00")
        published = str(record.get("published_at") or "")
        return deadline, published
    return sorted(records, key=key)

def announcement_key(record: dict[str, Any]) -> str:
    return str(record.get("guid") or record.get("plan_code") or "")


def announcement_signature(record: dict[str, Any]) -> str:
    important = [
        "招标项目名称", "标段合同估算价", "资金来源", "建设规模", "工程类型",
        "资格审查方式", "建设单位", "建设单位办公电话", "招标代理机构",
        "代理机构经办人", "代理机构办公电话", "代理机构移动电话",
        "项目行政主管地区", "交易地点", "发布时间", "公告结束时间", "公告性质",
        "联系人", "联系电话", "招标内容", "资质要求", "开标时间",
        "招标文件获取截止时间", "递交投标文件截止时间", "投标保证金缴纳截止时间",
        "投标保证金", "附件", "备注",
    ]
    return json.dumps({k: record.get(k) for k in important}, ensure_ascii=False, sort_keys=True)


def announcement_cutoff(record: dict[str, Any]) -> datetime | None:
    values = []
    for key in ("招标文件获取截止时间", "递交投标文件截止时间", "投标保证金缴纳截止时间"):
        deadline = parse_deadline(record.get(key))
        if deadline is not None:
            values.append(deadline)
    return min(values) if values else None


def announcement_is_expired(record: dict[str, Any], at: datetime) -> bool:
    cutoff = announcement_cutoff(record)
    return cutoff is not None and cutoff <= at


def purge_announcements(
    records: dict[str, dict[str, Any]], at: datetime
) -> list[dict[str, Any]]:
    retention_cutoff = at - timedelta(days=ANNOUNCEMENT_RETENTION_DAYS)
    expired = []
    for key, record in list(records.items()):
        anchor = parse_deadline(record.get("published_at")) or announcement_cutoff(record)
        if announcement_is_expired(record, at) or (anchor and anchor < retention_cutoff):
            expired.append(record)
            records.pop(key, None)
    return expired


def upsert_announcements(
    records: dict[str, dict[str, Any]], incoming: list[dict[str, Any]]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    added, updated = [], []
    for record in incoming:
        key = announcement_key(record)
        if not key:
            continue
        previous = records.get(key)
        if previous is None:
            records[key] = record
            added.append(record)
        elif announcement_signature(previous) != announcement_signature(record):
            records[key] = record
            updated.append(record)
    return added, updated


def migrate_announcement_links(records: dict[str, dict[str, Any]]) -> int:
    changed = 0
    for record in records.values():
        url = announcement_url(record.get("guid"))
        if url and record.get("备注") != url:
            record["备注"] = url
            changed += 1
    return changed


def sort_announcements(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        records,
        key=lambda record: (
            str(record.get("published_at") or ""),
            str(record.get("deadline") or ""),
        ),
        reverse=True,
    )


def build_announcement_summary(
    run_at: datetime,
    active: list[dict[str, Any]],
    added: list[dict[str, Any]],
    updated: list[dict[str, Any]],
    expired: list[dict[str, Any]],
) -> dict[str, Any]:
    regions = Counter(str(x.get("项目行政主管地区") or "未注明") for x in active)
    total_amount = sum(
        float(x["标段合同估算价"])
        for x in active
        if x.get("标段合同估算价") not in (None, "")
    )
    ongoing = []
    ending_7d = []
    end_limit = run_at + timedelta(days=7)
    for record in active:
        deadline = announcement_cutoff(record)
        if deadline is None or deadline > run_at:
            ongoing.append(record)
            if deadline is not None and deadline <= end_limit:
                ending_7d.append(record)
    return {
        "run_at": iso_or_empty(run_at),
        "retention_days": ANNOUNCEMENT_RETENTION_DAYS,
        "active_count": len(active),
        "ongoing_count": len(ongoing),
        "ending_7d_count": len(ending_7d),
        "total_amount_wan": round(total_amount, 2),
        "added_count": len(added),
        "updated_count": len(updated),
        "expired_count": len(expired),
        "region_counts": dict(regions.most_common()),
        "added_names": [x.get("招标项目名称", "") for x in added[:20]],
        "updated_names": [x.get("招标项目名称", "") for x in updated[:20]],
        "expired_names": [x.get("招标项目名称", "") for x in expired[:20]],
        "ending_7d": [
            {
                "name": x.get("招标项目名称", ""),
                "deadline": x.get("cutoff_raw", ""),
                "area": x.get("项目行政主管地区", ""),
                "amount": x.get("标段合同估算价"),
                "tenderer": x.get("建设单位", ""),
            }
            for x in sorted(ending_7d, key=lambda y: str(y.get("deadline") or ""))[:12]
        ],
    }


def load_announcement_state() -> dict[str, Any]:
    state = load_json(ANNOUNCEMENT_STATE_FILE, {})
    if not isinstance(state, dict) or not isinstance(state.get("records"), dict):
        return {"schema_version": 1, "last_success_at": "", "records": {}}
    return state


def run_announcement_update(run_at: datetime, args: argparse.Namespace) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    state = load_announcement_state()
    records: dict[str, dict[str, Any]] = state.get("records") or {}
    migrated = migrate_announcement_links(records)
    if migrated:
        LOG.info("迁移公告链接：更新 %s 条历史记录", migrated)
    expired_before = purge_announcements(records, run_at)

    last_success = str(state.get("last_success_at") or "")
    if last_success:
        try:
            window_start = datetime.fromisoformat(last_success) - timedelta(days=OVERLAP_DAYS)
            first_run = False
        except ValueError:
            window_start = run_at - timedelta(days=ANNOUNCEMENT_RETENTION_DAYS)
            first_run = True
    else:
        window_start = run_at - timedelta(days=ANNOUNCEMENT_RETENTION_DAYS)
        first_run = True

    LOG.info(
        "招标公告抓取窗口：%s ~ %s%s",
        window_start.strftime("%Y-%m-%d %H:%M"),
        run_at.strftime("%Y-%m-%d %H:%M"),
        "（首次）" if first_run else "（增量）",
    )
    raw_items = fetch_announcement_list_window(window_start, run_at)
    seen = set()
    listed = []
    for item in raw_items:
        guid = str(item.get("guid") or "")
        if not guid or guid in seen:
            continue
        seen.add(guid)
        listed.append(item)
    LOG.info("公告窗口内去重后待取详情：%s 个", len(listed))

    failures = []
    incoming = []
    excluded_incoming = []
    for idx, item in enumerate(listed, 1):
        guid = str(item.get("guid") or "")
        try:
            detail = fetch_announcement_detail(guid, max_cache_hours=args.cache_hours)
            record = parse_announcement_record(detail, item)
            if announcement_is_expired(record, run_at):
                excluded_incoming.append(record)
            else:
                incoming.append(record)
        except Exception as exc:
            LOG.warning("公告详情失败 %s/%s %s: %s", idx, len(listed), guid, exc)
            failures.append({"guid": guid, "name": item.get("bulletinname", ""), "error": str(exc)})
        if idx % 20 == 0 or idx == len(listed):
            LOG.info("公告详情进度 %s/%s", idx, len(listed))
        time.sleep(REQUEST_INTERVAL)

    added, updated = upsert_announcements(records, incoming)
    expired_after = purge_announcements(records, run_at)
    expired_unique = {announcement_key(x): x for x in (expired_before + expired_after + excluded_incoming)}
    active = sort_announcements(list(records.values()))
    summary = build_announcement_summary(
        run_at, active, added, updated, list(expired_unique.values())
    )
    summary["failures"] = failures
    summary["fetched_count"] = len(incoming) + len(excluded_incoming)
    next_success = str(state.get("last_success_at") or "") if failures else iso_or_empty(run_at)
    atomic_write_json(ANNOUNCEMENT_SUMMARY_FILE, summary)
    atomic_write_json(
        ANNOUNCEMENT_STATE_FILE,
        {
            "schema_version": 1,
            "last_success_at": next_success,
            "records": records,
        },
    )
    LOG.info(
        "公告完成：最近 %s 天 %s 条，进行中 %s，新增 %s，更新 %s，清理 %s，失败 %s",
        ANNOUNCEMENT_RETENTION_DAYS,
        len(active),
        summary["ongoing_count"],
        len(added),
        len(updated),
        len(expired_unique),
        len(failures),
    )
    return active, summary


def write_xlsx(rows: list[dict[str, Any]], path: Path, title: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.append(COLUMN_NAMES)
    for record in rows:
        ws.append([record.get(name, "") for name in COLUMN_NAMES])
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30
    header_fill = PatternFill("solid", fgColor="123B5D")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="D9E2EA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for idx, (_, width) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        row[1].number_format = "#,##0.00"
        link_cell = row[12]
        if str(link_cell.value or "").startswith(("http://", "https://")):
            link_cell.hyperlink = str(link_cell.value)
            link_cell.font = Font(color="0563C1", underline="single")
        ws.row_dimensions[row[0].row].height = 64
    if ws.max_row >= 2:
        ws.conditional_formatting.add(
            f"A2:M{ws.max_row}",
            FormulaRule(
                formula=['$F2<>""'],
                fill=PatternFill("solid", fgColor="FFF4D6"),
            ),
        )
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.properties.title = "红河州招标计划"
    wb.properties.subject = "云南省公共资源交易平台招标计划"
    wb.properties.creator = "红河州招标计划每日工作流"
    wb.save(path)


def write_csv(rows: list[dict[str, Any]], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMN_NAMES, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def build_summary(
    run_at: datetime,
    active: list[dict[str, Any]],
    added: list[dict[str, Any]],
    updated: list[dict[str, Any]],
    expired: list[dict[str, Any]],
    fetched_count: int,
    window_start: datetime,
    window_end: datetime,
) -> dict[str, Any]:
    regions = Counter(str(x.get("项目行政主管地区") or "未注明") for x in active)
    total_amount = sum(
        float(x["标段合同估算价（万元）"])
        for x in active
        if x.get("标段合同估算价（万元）") not in (None, "")
    )
    seven_days = run_at + timedelta(days=7)
    soon = []
    for record in active:
        deadline = parse_deadline(record.get("deadline"))
        if deadline and run_at < deadline <= seven_days:
            soon.append(record)
    soon = sort_active(soon)
    return {
        "run_at": iso_or_empty(run_at),
        "window_start": iso_or_empty(window_start),
        "window_end": iso_or_empty(window_end),
        "fetched_count": fetched_count,
        "active_count": len(active),
        "active_total_amount_wan": round(total_amount, 2),
        "added_count": len(added),
        "updated_count": len(updated),
        "expired_count": len(expired),
        "soon_7d_count": len(soon),
        "region_counts": dict(regions.most_common()),
        "added_names": [x.get("招标项目名称", "") for x in added[:20]],
        "updated_names": [x.get("招标项目名称", "") for x in updated[:20]],
        "expired_names": [x.get("招标项目名称", "") for x in expired[:20]],
        "soon_7d": [
            {
                "name": x.get("招标项目名称", ""),
                "deadline": x.get("预计招标公告发布时间", ""),
                "area": x.get("项目行政主管地区", ""),
                "amount": x.get("标段合同估算价（万元）"),
                "contact": x.get("招标人联系人", ""),
                "phone": x.get("招标人联系电话", ""),
            }
            for x in soon[:12]
        ],
    }


def summary_markdown(summary: dict[str, Any]) -> str:
    run_at = datetime.fromisoformat(summary["run_at"])
    lines = [
        f"# 红河州招标计划日报 · {run_at:%Y-%m-%d %H:%M}",
        "",
        f"- 当前有效项目：**{summary['active_count']}** 个",
        f"- 估算总额：**{money(summary['active_total_amount_wan'])} 万元**",
        f"- 本次新增：**{summary['added_count']}** 个",
        f"- 本次更新：**{summary['updated_count']}** 个",
        f"- 自动删除过期：**{summary['expired_count']}** 个",
        f"- 未来 7 天将发布：**{summary['soon_7d_count']}** 个",
    ]
    announcement = summary.get("announcements") or {}
    if announcement:
        lines.extend([
            f"- 有效招标公告：**{announcement.get('active_count', 0)}** 条",
            f"- 其中 7 天内截止：**{announcement.get('ending_7d_count', 0)}** 条",
        ])
    lines.extend([
        "",
        "## 地区分布",
    ])
    if summary["region_counts"]:
        lines.append(" · ".join(f"{k} {v}" for k, v in summary["region_counts"].items()))
    else:
        lines.append("暂无有效项目")
    lines.extend(["", "## 未来 7 天"])
    if summary["soon_7d"]:
        for item in summary["soon_7d"]:
            lines.append(
                f"- {item['deadline'] or '待定'}｜{item['area'] or '-'}｜"
                f"{item['name']}｜{money(item['amount'])} 万元"
            )
    else:
        lines.append("- 暂无")
    lines.extend(["", "## 本次新增"])
    if summary["added_names"]:
        lines.extend(f"- {name}" for name in summary["added_names"])
    else:
        lines.append("- 无")
    lines.extend(
        [
            "",
            "> 数据来源：云南省公共资源交易平台。"
            "“预计招标公告发布时间”为计划时间，最终以实际公告为准。",
            "",
        ]
    )
    return "\n".join(lines)

HTML_TEMPLATE = r'''<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>红河州招标计划 · 离线查询</title>
<style>
:root{--navy:#123b5d;--blue:#1677a8;--teal:#1d8f83;--bg:#eef4f7;--card:#fff;--ink:#173042;--muted:#718493;--line:#dce7ed;--amber:#c77800;--danger:#b53b3b}
*{box-sizing:border-box}body{margin:0;font-family:"Microsoft YaHei UI","Microsoft YaHei","PingFang SC",Arial,sans-serif;background:linear-gradient(145deg,#eaf3f7,#f5f7f4 55%,#edf5f1);color:var(--ink);min-height:100vh;overflow-x:hidden}
.hero{padding:28px clamp(18px,4vw,52px) 72px;background:radial-gradient(circle at 88% 15%,rgba(255,255,255,.22),transparent 25%),linear-gradient(120deg,#0d334f,#126080 58%,#22877b);color:#fff;position:relative;overflow:hidden}
.hero:after{content:"";position:absolute;width:280px;height:280px;border:1px solid rgba(255,255,255,.17);border-radius:50%;right:-80px;top:-110px;box-shadow:0 0 0 45px rgba(255,255,255,.04),0 0 0 95px rgba(255,255,255,.03)}
.kicker{font-size:12px;letter-spacing:.18em;text-transform:uppercase;opacity:.72}.hero h1{font-size:clamp(26px,4vw,44px);margin:10px 0 8px;letter-spacing:.02em}.hero p{margin:0;opacity:.78;font-size:14px}.badge{display:inline-flex;align-items:center;gap:7px;margin-top:18px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);padding:7px 12px;border-radius:999px;font-size:12px;backdrop-filter:blur(6px)}
main{max-width:1560px;margin:-42px auto 0;padding:0 clamp(12px,3vw,34px) 50px;position:relative}.tabs{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:9px;max-width:520px;margin:0 0 12px}.tab{height:48px;border:1px solid rgba(220,231,237,.92);background:rgba(255,255,255,.95);border-radius:13px;color:#526a79;font:inherit;font-weight:700;cursor:pointer;box-shadow:0 8px 22px rgba(23,48,66,.07);display:flex;align-items:center;justify-content:center;gap:8px}.tab span{font-size:11px;padding:2px 7px;border-radius:999px;background:#e7f0f4;color:#526a79}.tab.active{background:linear-gradient(120deg,#123b5d,#176d82);border-color:transparent;color:#fff}.tab.active span{background:rgba(255,255,255,.16);color:#fff}.stats{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:13px}.stat{background:rgba(255,255,255,.97);border:1px solid rgba(220,231,237,.9);border-radius:16px;padding:18px 19px;box-shadow:0 12px 30px rgba(23,48,66,.08)}.stat b{display:block;font-size:27px;letter-spacing:-.03em}.stat span{font-size:12px;color:var(--muted)}.stat.alert b{color:var(--amber)}
.panel{margin-top:16px;background:rgba(255,255,255,.96);border:1px solid var(--line);border-radius:16px;box-shadow:0 12px 32px rgba(23,48,66,.07);overflow:hidden}.toolbar{display:grid;grid-template-columns:minmax(240px,2fr) repeat(4,minmax(120px,1fr));gap:10px;padding:15px}.toolbar input,.toolbar select{border:1px solid #cfdde5;background:#f9fbfc;border-radius:10px;height:40px;padding:0 12px;font:inherit;color:var(--ink);outline:none}.toolbar input:focus,.toolbar select:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(22,119,168,.12)}
.notice{padding:0 16px 13px;color:var(--muted);font-size:12px}.table-wrap{overflow:auto;max-height:66vh;border-top:1px solid var(--line)}table{border-collapse:separate;border-spacing:0;width:100%;min-width:1180px;font-size:13px}thead th{position:sticky;top:0;z-index:2;background:#f2f7f9;color:#526a79;font-weight:700;text-align:left;padding:12px 11px;border-bottom:1px solid var(--line);white-space:nowrap}tbody td{padding:12px 11px;border-bottom:1px solid #edf2f5;vertical-align:top}tbody tr{cursor:pointer;transition:background .15s}tbody tr:hover{background:#f2f9fa}.name{font-weight:700;line-height:1.5;min-width:280px}.amount{font-variant-numeric:tabular-nums;white-space:nowrap}.deadline{white-space:nowrap;font-weight:700}.soon{color:var(--amber)}.region{display:inline-block;padding:3px 8px;border-radius:999px;background:#e7f4f1;color:#176e64;white-space:nowrap}.empty{padding:64px 20px;text-align:center;color:var(--muted)}.footer{padding:14px 18px;border-top:1px solid var(--line);font-size:12px;color:var(--muted);display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}.footer a{color:var(--blue)}
.cards{display:none}.card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;box-shadow:0 5px 18px rgba(23,48,66,.06);min-width:0;width:100%;overflow:hidden}.card-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px;min-width:0}.card h3{font-size:15px;line-height:1.55;margin:0 0 12px;font-weight:700;overflow-wrap:anywhere;word-break:break-word}.card-deadline{font-size:11px;color:var(--muted);white-space:nowrap;max-width:58%;overflow:hidden;text-overflow:ellipsis}.card-deadline.soon{color:var(--amber);font-weight:700}.card-main{display:flex;justify-content:space-between;align-items:flex-end;gap:12px;margin-bottom:8px}.card-price{font-size:21px;font-weight:800;letter-spacing:-.03em}.card-price small{font-size:11px;color:var(--muted);margin-left:3px;font-weight:400}.card-nature{font-size:11px;color:#526a79;background:#f1f6f8;border-radius:999px;padding:3px 8px}.card-nature.ended{background:#eef0f2;color:#77838c}.card-tenderer{font-size:12px;color:#526a79;line-height:1.55;margin-bottom:10px}.card-foot{display:flex;justify-content:space-between;gap:10px;align-items:center;padding-top:10px;border-top:1px solid #edf2f5;font-size:11px;color:var(--muted)}.card-link{color:var(--blue);font-weight:700;text-decoration:none;white-space:nowrap}.card-empty{grid-column:1/-1;padding:48px 12px;text-align:center;color:var(--muted)}
.modal{position:fixed;inset:0;background:rgba(10,29,42,.58);display:none;align-items:center;justify-content:center;padding:18px;z-index:20}.modal.open{display:flex}.modal-card{background:#fff;border-radius:18px;width:min(860px,96vw);max-height:90vh;overflow:auto;box-shadow:0 24px 80px rgba(0,0,0,.25)}.modal-head{position:sticky;top:0;background:linear-gradient(120deg,#123b5d,#1a6f87);color:#fff;padding:20px 24px;display:flex;justify-content:space-between;gap:16px}.modal-head h2{font-size:19px;line-height:1.5;margin:0}.close{border:0;background:rgba(255,255,255,.14);color:#fff;border-radius:50%;width:34px;height:34px;font-size:22px;cursor:pointer;flex:none}.detail-grid{display:grid;grid-template-columns:1fr 1fr;gap:0;padding:10px 24px 24px}.detail-item{padding:14px 0;border-bottom:1px solid #edf2f5}.detail-item.full{grid-column:1/-1}.detail-item label{display:block;color:var(--muted);font-size:11px;margin-bottom:6px}.detail-item div{line-height:1.75;white-space:pre-wrap;word-break:break-word}.phone{color:var(--blue);font-weight:700}.hidden{display:none!important}
@media(max-width:900px){.stats{grid-template-columns:repeat(2,1fr)}.toolbar{grid-template-columns:1fr 1fr}.toolbar input{grid-column:1/-1}.modal-card{max-height:94vh}.detail-grid{grid-template-columns:1fr}.detail-item.full{grid-column:auto}}
@media(max-width:720px){body{background:#eef4f7}.hero{padding:20px 16px 52px}.hero:after{width:180px;height:180px;right:-90px;top:-90px}.kicker{display:none}.hero h1{font-size:26px;margin:0 0 6px}.hero p{font-size:12px}.badge{margin-top:12px;padding:6px 10px;font-size:11px}main{margin:-30px auto 0;padding:0 10px 30px}.tabs{max-width:none;margin-bottom:9px}.tab{height:44px;font-size:13px}.stats{gap:8px}.stat{padding:12px 13px;border-radius:12px}.stat b{font-size:21px}.stat span{font-size:10.5px}.panel{border-radius:14px}.toolbar{grid-template-columns:1fr 1fr;gap:8px;padding:11px}.toolbar input{grid-column:1/-1;font-size:16px}.toolbar select{height:38px;font-size:13px;padding:0 8px}.notice{padding:0 11px 10px;font-size:11px}.table-wrap{display:none}.cards{display:grid;grid-template-columns:minmax(0,1fr);gap:10px;padding:11px;border-top:1px solid var(--line);background:#f5f9fa;min-width:0}.footer{padding:12px;font-size:11px}.modal{padding:0;align-items:flex-end}.modal-card{width:100%;max-height:92vh;border-radius:18px 18px 0 0}.modal-head{padding:16px}.modal-head h2{font-size:16px}.detail-grid{padding:0 16px 22px}.detail-item{padding:12px 0}}
@media(max-width:520px){.stats{grid-template-columns:1fr 1fr}.toolbar{grid-template-columns:repeat(2,minmax(0,1fr))}.toolbar input{grid-column:1/-1}.footer{display:block}.footer span{display:block;margin-top:6px}}
@media print{.hero{background:#123b5d!important;-webkit-print-color-adjust:exact}.toolbar,.notice,.footer,.modal{display:none!important}.table-wrap{max-height:none;overflow:visible}.panel{box-shadow:none}table{min-width:0;font-size:10px}thead th{position:static}}
</style>
</head>
<body>
<header class="hero">
  <div class="kicker">Honghe · Tender Plan Monitor</div>
  <h1>红河州招标计划</h1>
  <p>每天自动更新 · 到期项目自动清理 · 离线可查</p>
  <div class="badge"><span>●</span> 数据更新时间：<strong id="runAt"></strong></div>
</header>
<main>
  <nav class="tabs" aria-label="信息类型切换">
    <button class="tab active" id="tabPlan" type="button">招标计划 <span id="tabPlanCount">0</span></button>
    <button class="tab" id="tabAnnouncement" type="button">招标公告 <span id="tabAnnouncementCount">0</span></button>
  </nav>
  <section class="stats">
    <div class="stat"><b id="activeCount">0</b><span id="statLabel1">当前有效项目</span></div>
    <div class="stat"><b id="knownCount">0</b><span id="statLabel2">本次更新/新增</span></div>
    <div class="stat alert"><b id="soonCount">0</b><span id="statLabel3">未来 7 天将发布</span></div>
    <div class="stat"><b id="totalAmount">0</b><span id="statLabel4">估算总额（万元）</span></div>
  </section>
  <section class="panel">
    <div class="toolbar">
      <input id="search" type="search" placeholder="搜索项目、招标人、联系人、内容…">
      <select id="region"><option value="">全部地区</option></select>
      <select id="window">
        <option value="all">全部有效期</option>
        <option value="3">3 天内</option>
        <option value="7">7 天内</option>
        <option value="30">30 天内</option>
      </select>
      <select id="amount">
        <option value="0">全部金额</option>
        <option value="100">≥ 100 万</option>
        <option value="500">≥ 500 万</option>
        <option value="1000">≥ 1000 万</option>
        <option value="5000">≥ 5000 万</option>
      </select>
      <select id="sort">
        <option value="deadline">按预计时间升序</option>
        <option value="publish">按发布时间降序</option>
        <option value="amount">按金额降序</option>
        <option value="region">按地区排序</option>
      </select>
    </div>
    <div class="notice" id="notice"></div>
    <div class="table-wrap">
      <table>
        <thead><tr><th id="thProject">招标项目</th><th>地区</th><th id="thDeadline">预计发布时间</th><th>金额（万元）</th><th id="thParty">招标人 / 联系人</th><th id="thNature">计划性质</th><th>发布时间</th></tr></thead>
        <tbody id="rows"></tbody>
      </table>
    </div>
    <div class="cards" id="cards"></div>
    <div class="footer"><span id="shown"></span><span>来源：<a href="https://ggzy.yn.gov.cn" target="_blank" rel="noopener">云南省公共资源交易平台</a> · 最终以实际公告为准</span></div>
  </section>
</main>
<div class="modal" id="modal" role="dialog" aria-modal="true">
  <div class="modal-card">
    <div class="modal-head"><h2 id="modalTitle"></h2><button class="close" id="closeModal" aria-label="关闭">×</button></div>
    <div class="detail-grid" id="modalBody"></div>
  </div>
</div>
<script>const PLAN_DATA = __DATA_JSON__;
const ANNOUNCEMENT_DATA = __ANNOUNCEMENT_DATA_JSON__;
const SUMMARY = __SUMMARY_JSON__;
const ANNOUNCEMENT_SUMMARY = __ANNOUNCEMENT_SUMMARY_JSON__;
let NOW = Date.now();
let hiddenExpired = [];
let currentType = 'plan';
let data = [];
const $ = id => document.getElementById(id);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const nf = new Intl.NumberFormat('zh-CN',{maximumFractionDigits:2});
const dtf = new Intl.DateTimeFormat('zh-CN',{year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hour12:false,timeZone:'Asia/Shanghai'});
function dateText(ms){ if(!ms) return '待定'; return dtf.format(new Date(ms)).replace(/\//g,'-'); }
function deadlineText(x){ let raw=String(x.deadline||'').trim(); if(!raw) return dateText(x.deadlineMs); if(/^\d{14}$/.test(raw)) raw=`${raw.slice(0,4)}-${raw.slice(4,6)}-${raw.slice(6,8)} ${raw.slice(8,10)}:${raw.slice(10,12)}`; else if(/^\d{8}$/.test(raw)) raw=`${raw.slice(0,4)}-${raw.slice(4,6)}-${raw.slice(6,8)}`; return /(?:^|[ T])00:00(?::00)?$/.test(raw)?raw.replace(/00:00(?::00)?$/,'00:00（当日有效）'):raw; }
function deadlineClass(ms){ return ms && ms>NOW && ms-NOW <= 7*86400000 ? 'soon' : ''; }
function amountValue(x){ const n=Number(x.amount); return Number.isFinite(n)?n:0; }
function statusText(x){ return x.deadlineMs && x.deadlineMs<=NOW ? '已结束' : '进行中'; }
function fillRegions(){
  while($('region').options.length>1) $('region').remove(1);
  const vals=[...new Set(data.map(x=>x.area).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'zh-CN'));
  for(const v of vals){ const o=document.createElement('option'); o.value=v;o.textContent=v;$('region').appendChild(o); }
}
function updateStats(){
  const raw=currentType==='plan'?PLAN_DATA:ANNOUNCEMENT_DATA;
  const summary=currentType==='plan'?SUMMARY:ANNOUNCEMENT_SUMMARY;
  $('tabPlanCount').textContent=PLAN_DATA.filter(x=>!x.deadlineMs||x.deadlineMs>NOW).length;
  $('tabAnnouncementCount').textContent=ANNOUNCEMENT_DATA.filter(x=>!x.deadlineMs||x.deadlineMs>NOW).length;
  $('activeCount').textContent=data.length;
  $('knownCount').textContent=(summary.added_count||0)+(summary.updated_count||0);
  $('soonCount').textContent=currentType==='plan'?(summary.soon_7d_count||0):(summary.ending_7d_count||0);
  $('totalAmount').textContent=nf.format(currentType==='plan'?(summary.active_total_amount_wan||0):(summary.total_amount_wan||0));
  $('statLabel1').textContent=currentType==='plan'?'当前有效项目':'最近30天公告';
  $('statLabel2').textContent='本次更新/新增';
  $('statLabel3').textContent=currentType==='plan'?'未来 7 天将发布':'7天内截止';
  $('statLabel4').textContent='估算总额（万元）';
}
function applyType(type){
  currentType=type; NOW=Date.now(); hiddenExpired=[];
  const raw=type==='plan'?PLAN_DATA:ANNOUNCEMENT_DATA;
  data=raw.filter(x=>{ if(!x.deadlineMs||x.deadlineMs>NOW) return true; hiddenExpired.push(x); return false; });
  $('tabPlan').classList.toggle('active',type==='plan');
  $('tabAnnouncement').classList.toggle('active',type==='announcement');
  $('thProject').textContent=type==='plan'?'招标项目':'公告项目';
  $('thDeadline').textContent=type==='plan'?'预计发布时间':'最近截止';
  $('thParty').textContent=type==='plan'?'招标人 / 联系人':'建设单位 / 联系人';
  $('thNature').textContent=type==='plan'?'计划性质':'公告状态';
  $('search').placeholder=type==='plan'?'搜索项目、招标人、联系人、内容…':'搜索公告、建设单位、代理机构、内容…';
  $('window').innerHTML=type==='plan'
    ?'<option value="all">全部有效期</option><option value="3">3 天内</option><option value="7">7 天内</option><option value="30">30 天内</option>'
    :'<option value="all">全部有效</option><option value="3">3天内截止</option><option value="7">7天内截止</option><option value="30">30天内截止</option>';
  $('sort').value=type==='plan'?'deadline':'publish';
  updateStats();
  fillRegions();
  render();
}
function render(){
  const q=$('search').value.trim().toLowerCase(), region=$('region').value, win=$('window').value, min=Number($('amount').value), sort=$('sort').value;
  let rows=data.filter(x=>{
    if(region && x.area!==region) return false;
    if(amountValue(x)<min) return false;
    if(win!=='all'){
      if(!x.deadlineMs) return false;
      if(x.deadlineMs<=NOW || x.deadlineMs-NOW > Number(win)*86400000) return false;
    }
    if(q){
      const blob=[x.name,x.tenderer,x.area,x.contact,x.phone,x.content,x.remark,x.platform,x.agency,x.agencyContact,x.tendererPhone,x.qualification,x.source,x.scale,x.projectType,x.reviewMethod,x.docDeadline,x.bidDeadline,x.bondDeadline].join(' ').toLowerCase();
      if(!blob.includes(q)) return false;
    }
    return true;
  });
  rows.sort((a,b)=>{
    if(sort==='publish') return String(b.publishedAt).localeCompare(String(a.publishedAt));
    if(sort==='amount') return amountValue(b)-amountValue(a);
    if(sort==='region') return String(a.area).localeCompare(String(b.area),'zh-CN') || (a.deadlineMs||Infinity)-(b.deadlineMs||Infinity);
    return (a.deadlineMs||Infinity)-(b.deadlineMs||Infinity);
  });
  $('rows').innerHTML=rows.length?rows.map(x=>{const status=currentType==='announcement'?statusText(x):(x.nature||'-');return `<tr data-i="${data.indexOf(x)}"><td class="name">${esc(x.name)}</td><td><span class="region">${esc(x.area||'-')}</span></td><td class="deadline ${deadlineClass(x.deadlineMs)}">${esc(deadlineText(x))}</td><td class="amount">${x.amount==null?'-':nf.format(amountValue(x))}</td><td>${esc(x.tenderer||'-')}<br><span style="color:var(--muted)">${esc(x.contact||'')}${x.phone?' · '+esc(x.phone):''}</span></td><td>${esc(status)}</td><td style="white-space:nowrap">${esc(x.published||'-')}</td></tr>`}).join(''):`<tr><td class="empty" colspan="7">没有符合条件的项目</td></tr>`;
  $('cards').innerHTML=rows.length?rows.map(x=>{const status=currentType==='announcement'?statusText(x):(x.nature||'-'); const ended=status==='已结束'?' ended':''; return `<article class="card" data-i="${data.indexOf(x)}"><div class="card-head"><span class="region">${esc(x.area||'-')}</span><span class="card-deadline ${deadlineClass(x.deadlineMs)}">${esc(deadlineText(x))}</span></div><h3>${esc(x.name)}</h3><div class="card-main"><div class="card-price">${x.amount==null?'-':nf.format(amountValue(x))}<small>万元</small></div><span class="card-nature${ended}">${esc(status)}</span></div><div class="card-tenderer">${esc(x.tenderer||'-')}${x.contact?' · '+esc(x.contact):''}${x.phone?' · '+esc(x.phone):''}</div><div class="card-foot"><span>发布 ${esc(x.published||'-')}</span>${x.remark?`<a class="card-link" href="${esc(x.remark)}" target="_blank" rel="noopener">查看原文 ›</a>`:''}</div></article>`}).join(''):`<div class="card-empty">没有符合条件的项目</div>`;
  $('shown').textContent=`显示 ${rows.length} / ${data.length} 个${currentType==='plan'?'有效项目':'公告'}`;
  const notice=[];
  if(currentType==='plan'){
    if(hiddenExpired.length) notice.push(`已按当前时间自动隐藏 ${hiddenExpired.length} 个过期项目`);
    if(SUMMARY.expired_count) notice.push(`上次更新自动清理 ${SUMMARY.expired_count} 个`);
  }else{
    if(hiddenExpired.length) notice.push(`已自动隐藏 ${hiddenExpired.length} 条超过关键截止时间的公告`);
    if(ANNOUNCEMENT_SUMMARY.expired_count) notice.push(`上次更新已排除 ${ANNOUNCEMENT_SUMMARY.expired_count} 条`);
  }
  $('notice').textContent=notice.join(' · ');
}
function showDetail(index){
  const x=data[index]; if(!x) return;
  $('modalTitle').textContent=x.name||'项目详情';
  const planItems=[['项目行政主管地区',x.area],['入场登记部门',x.department],['招标人',x.tenderer],['预计招标公告发布时间',deadlineText(x)],['标段合同估算价（万元）',x.amount==null?'-':nf.format(amountValue(x))],['计划性质',x.nature],['发布时间',x.published],['交易平台',x.platform],['招标人联系人',x.contact],['联系电话',x.phone],['招标内容',x.content],['项目链接',x.remark]];
  const announcementItems=[['项目行政主管地区',x.area],['建设单位',x.tenderer],['建设单位办公电话',x.tendererPhone],['招标代理机构',x.agency],['代理机构经办人',x.agencyContact],['代理机构办公电话',x.agencyOfficePhone],['代理机构移动电话',x.agencyMobilePhone],['资金来源',x.source],['建设规模',x.scale],['工程类型',x.projectType],['资格审查方式',x.reviewMethod],['招标文件获取截止时间',x.docDeadline],['递交投标文件截止时间',x.bidDeadline],['投标保证金缴纳截止时间',x.bondDeadline],['最近截止时间',deadlineText(x)],['公告结束时间',x.announcementEnd],['标段合同估算价（万元）',x.amount==null?'-':nf.format(amountValue(x))],['公告性质',x.nature],['发布时间',x.published],['开标时间',x.openTime],['投标保证金',x.bidBond],['资质要求',x.qualification],['招标内容',x.content],['项目链接',x.remark]];
  const items=currentType==='plan'?planItems:announcementItems;
  $('modalBody').innerHTML=items.map(([k,v])=>{ const isLink=k==='项目链接'&&/^https?:\/\//.test(String(v||'')); const body=isLink?`<a href="${esc(v)}" target="_blank" rel="noopener">${esc(v)}</a>`:esc(v||'-'); const full=['建设规模','招标内容','资质要求','项目链接'].includes(k)?' full':''; return `<div class="detail-item${full}"><label>${esc(k)}</label><div class="${k==='联系电话'?'phone':''}">${body}</div></div>`; }).join('');
  $('modal').classList.add('open');
}
$('runAt').textContent=SUMMARY.run_at?SUMMARY.run_at.replace('T',' ').slice(0,16):'-';
['search','region','window','amount','sort'].forEach(id=>$(id).addEventListener(id==='search'?'input':'change',render));
$('tabPlan').addEventListener('click',()=>applyType('plan'));
$('tabAnnouncement').addEventListener('click',()=>applyType('announcement'));
$('rows').addEventListener('click',e=>{const tr=e.target.closest('tr[data-i]'); if(tr) showDetail(Number(tr.dataset.i));});
$('cards').addEventListener('click',e=>{if(e.target.closest('a')) return; const card=e.target.closest('article[data-i]'); if(card) showDetail(Number(card.dataset.i));});
$('closeModal').onclick=()=>$('modal').classList.remove('open');
$('modal').onclick=e=>{if(e.target===$('modal'))$('modal').classList.remove('open')};
document.addEventListener('keydown',e=>{if(e.key==='Escape')$('modal').classList.remove('open')});
applyType('plan');
</script>
</body>
</html>'''


def write_html(
    active: list[dict[str, Any]],
    summary: dict[str, Any],
    announcements: list[dict[str, Any]],
    announcement_summary: dict[str, Any],
    path: Path,
) -> None:
    payload = []
    for x in active:
        deadline_text = x.get("deadline") or ""
        deadline_ms = int(datetime.fromisoformat(deadline_text).timestamp() * 1000) if deadline_text else None
        payload.append({
            "name": x.get("招标项目名称", ""),
            "amount": x.get("标段合同估算价（万元）"),
            "tenderer": x.get("招标人", ""),
            "area": x.get("项目行政主管地区", ""),
            "department": x.get("入场登记部门", ""),
            "deadline": x.get("预计招标公告发布时间", ""),
            "deadlineMs": deadline_ms,
            "contact": x.get("招标人联系人", ""),
            "phone": x.get("招标人联系电话", ""),
            "nature": x.get("计划性质", ""),
            "published": x.get("发布时间", ""),
            "publishedAt": x.get("published_at", ""),
            "platform": x.get("交易平台", ""),
            "content": x.get("招标内容", ""),
            "remark": x.get("备注", ""),
        })
    announcement_payload = []
    for x in announcements:
        deadline_text = x.get("deadline") or ""
        deadline_ms = int(datetime.fromisoformat(deadline_text).timestamp() * 1000) if deadline_text else None
        announcement_payload.append({
            "name": x.get("招标项目名称", ""),
            "projectCode": x.get("plan_code", ""),
            "amount": x.get("标段合同估算价"),
            "source": x.get("资金来源", ""),
            "scale": x.get("建设规模", ""),
            "projectType": x.get("工程类型", ""),
            "reviewMethod": x.get("资格审查方式", ""),
            "tenderer": x.get("建设单位", ""),
            "tendererPhone": x.get("建设单位办公电话", ""),
            "agency": x.get("招标代理机构", ""),
            "agencyContact": x.get("代理机构经办人", ""),
            "agencyOfficePhone": x.get("代理机构办公电话", ""),
            "agencyMobilePhone": x.get("代理机构移动电话", ""),
            "area": x.get("项目行政主管地区", ""),
            "department": x.get("交易地点", ""),
            "deadline": x.get("cutoff_raw", ""),
            "deadlineMs": deadline_ms,
            "announcementEnd": x.get("公告结束时间", ""),
            "contact": x.get("联系人", ""),
            "phone": x.get("联系电话", ""),
            "nature": x.get("公告性质", ""),
            "published": x.get("发布时间", ""),
            "publishedAt": x.get("published_at", ""),
            "platform": x.get("项目行政主管地区", ""),
            "content": x.get("招标内容", ""),
            "remark": x.get("备注", ""),
            "qualification": x.get("资质要求", ""),
            "openTime": x.get("开标时间", ""),
            "docDeadline": x.get("招标文件获取截止时间", ""),
            "bidDeadline": x.get("递交投标文件截止时间", ""),
            "bondDeadline": x.get("投标保证金缴纳截止时间", ""),
            "bidBond": x.get("投标保证金", ""),
        })
    data_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\/")
    announcement_data_json = json.dumps(announcement_payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\/")
    summary_json = json.dumps(summary, ensure_ascii=False, separators=(",", ":")).replace("</", "<\/")
    announcement_summary_json = json.dumps(announcement_summary, ensure_ascii=False, separators=(",", ":")).replace("</", "<\/")
    text = (HTML_TEMPLATE
        .replace("__DATA_JSON__", data_json)
        .replace("__ANNOUNCEMENT_DATA_JSON__", announcement_data_json)
        .replace("__SUMMARY_JSON__", summary_json)
        .replace("__ANNOUNCEMENT_SUMMARY_JSON__", announcement_summary_json))
    atomic_write_text(path, text)

def write_portable_zip() -> None:
    tmp = PORTABLE_ZIP.with_suffix(".zip.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(QUERY_HTML, arcname="红河州招标计划_查询.html")
        zf.write(WECHAT_HTML, arcname="红河州招标计划_微信版.html")
        zf.write(ACTIVE_CSV, arcname="红河州招标计划_有效项目.csv")
        zf.write(REPORT_MD, arcname="最新日报.md")
        readme = (
            "红河州招标计划 · 便携查询包\n"
            "==============================\n\n"
            "1. 手机微信建议打开“红河州招标计划_微信版.html”。\n"
            "2. 电脑双击“红河州招标计划_查询.html”即可离线查询。\n"
            "3. 不需要安装 Python、Excel 或任何软件。\n"
            "4. 查询页已嵌入生成时的数据；如需最新数据，请重新运行每日工作流。\n"
            "5. 数据来源：云南省公共资源交易平台。\n"
        )
        zf.writestr("使用说明.txt", readme)
    os.replace(tmp, PORTABLE_ZIP)

def write_public_zip() -> None:
    tmp = PUBLIC_ZIP.with_suffix(".zip.tmp")
    with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(PUBLISH_INDEX, arcname="index.html")
        zf.writestr(".nojekyll", "")
        zf.writestr(
            "部署说明.txt",
            "红河州招标计划 · 公网部署包\n"
            "==============================\n\n"
            "1. 解压后应得到 index.html 和 .nojekyll。\n"
            "2. 将这两个文件上传到静态网站服务的根目录。\n"
            "3. 访问网站首页即可，不需要服务器或数据库。\n"
            "4. 页面已包含生成时的数据，可与每日更新任务配合发布。\n",
        )
    os.replace(tmp, PUBLIC_ZIP)


def load_state() -> dict[str, Any]:
    state = load_json(STATE_FILE, {})
    if not isinstance(state, dict) or not isinstance(state.get("records"), dict):
        return {"schema_version": 1, "last_success_at": "", "records": {}}
    return state


def default_window(last_success: str, run_at: datetime) -> tuple[datetime, datetime, bool]:
    if last_success:
        try:
            start = datetime.fromisoformat(last_success) - timedelta(days=OVERLAP_DAYS)
            return start, run_at, False
        except ValueError:
            pass
    return run_at - timedelta(days=BOOTSTRAP_DAYS), run_at, True


def unique_incoming(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen = set()
    result = []
    for item in items:
        key = item.get("tenderProjectCode") or item.get("guid")
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def run(args: argparse.Namespace) -> dict[str, Any]:
    run_at = now_cn()
    state = load_state()
    records: dict[str, dict[str, Any]] = state.get("records") or {}
    migrated = migrate_links(records)
    if migrated:
        LOG.info("迁移项目链接：更新 %s 条历史记录", migrated)
    expired_before = purge_expired(records, run_at)
    LOG.info("启动清理：删除 %s 个已过期项目", len(expired_before))

    window_start, window_end, bootstrap = default_window(
        str(state.get("last_success_at") or ""), run_at
    )
    if args.bootstrap_days:
        window_start = run_at - timedelta(days=args.bootstrap_days)
        bootstrap = True
    LOG.info(
        "抓取窗口：%s ~ %s%s",
        window_start.strftime("%Y-%m-%d %H:%M"),
        window_end.strftime("%Y-%m-%d %H:%M"),
        "（首次/重建）" if bootstrap else "（增量）",
    )
    listed = unique_incoming(fetch_list_window(window_start, window_end))
    LOG.info("窗口内去重后待取详情：%s 个", len(listed))

    incoming: list[dict[str, Any]] = []
    failures: list[dict[str, str]] = []
    for idx, item in enumerate(listed, 1):
        guid = str(item.get("guid") or "")
        try:
            detail = fetch_detail(guid, max_cache_hours=args.cache_hours)
            record = parse_record(detail, item)
            if is_honghe(detail, record, item):
                incoming.append(record)
        except Exception as exc:
            LOG.warning("详情失败 %s/%s %s: %s", idx, len(listed), guid, exc)
            failures.append({"guid": guid, "name": item.get("tenderProjectName", ""), "error": str(exc)})
        if idx % 20 == 0 or idx == len(listed):
            LOG.info("详情进度 %s/%s", idx, len(listed))
        time.sleep(REQUEST_INTERVAL)

    added, updated, expired_incoming = upsert_records(records, incoming, run_at)
    expired_after = purge_expired(records, run_at)
    expired_all = expired_before + expired_incoming + expired_after
    expired_unique = {record_key(x): x for x in expired_all}
    active = sort_active(list(records.values()))
    today = run_at.strftime("%Y-%m-%d")
    today_rows = [
        x for x in (added + updated)
        if str(x.get("发布时间") or "").startswith(today)
    ]
    summary = build_summary(
        run_at, active, added, updated, list(expired_unique.values()),
        len(incoming), window_start, window_end
    )
    summary["failures"] = failures
    leftover = [x for x in active if not record_future(x, run_at)]
    if leftover:
        summary["internal_error"] = f"清理后仍有 {len(leftover)} 个过期项目"
        raise RuntimeError(summary["internal_error"])

    announcement_active, announcement_summary = run_announcement_update(run_at, args)
    summary["announcements"] = announcement_summary

    write_xlsx(active, ACTIVE_XLSX, "有效项目")
    write_xlsx(sort_active(today_rows), TODAY_XLSX, "今日更新")
    write_csv(active, ACTIVE_CSV)
    write_html(active, summary, announcement_active, announcement_summary, QUERY_HTML)
    html_bytes = QUERY_HTML.read_bytes()
    WECHAT_HTML.write_bytes(html_bytes)
    PUBLISH_INDEX.write_bytes(html_bytes)
    (PUBLISH_DIR / ".nojekyll").write_bytes(b"")
    write_public_zip()
    report = summary_markdown(summary)
    atomic_write_text(REPORT_MD, report)
    write_portable_zip()
    atomic_write_json(SUMMARY_FILE, summary)
    atomic_write_json(
        STATE_FILE,
        {
            "schema_version": 1,
            "last_success_at": iso_or_empty(run_at),
            "records": records,
        },
    )
    print("\n" + report)
    LOG.info(
        "完成：有效 %s，新增 %s，更新 %s，清理 %s，失败 %s",
        len(active), len(added), len(updated), len(expired_unique), len(failures),
    )
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="红河州招标计划每日工作流")
    parser.add_argument("--bootstrap-days", type=int, default=0, help="强制回看指定天数并重建")
    parser.add_argument("--cache-hours", type=float, default=12.0, help="详情缓存有效小时数")
    return parser.parse_args()


if __name__ == "__main__":
    run(parse_args())