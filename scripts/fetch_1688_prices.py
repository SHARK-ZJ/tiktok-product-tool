#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1688 cost fetcher for TikTok product exports.

Input: CSV/XLSX exported from Dianxiaomi/TikTok with a source URL column.
Output: CSV that can be imported into the frontend tool as a 1688 cost table.

The script uses Python standard libraries only. 1688 may require login or block
requests; blocked pages are marked as failed and the task continues.
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import random
import re
import sys
import time
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET


SOURCE_URL_FIELDS = [
    "来源Url",
    "来源URL",
    "来源url",
    "source_url",
    "Source URL",
    "1688 URL",
    "Product URL",
    "1688链接",
]

OUTPUT_FIELDS = [
    "来源Url",
    "offerId",
    "1688商品标题",
    "1688采购价",
    "1688重量(g)",
    "1688运费",
    "抓取状态",
    "备注",
]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


@dataclass
class FetchResult:
    source_url: str
    offer_id: str = ""
    title: str = ""
    price_rmb: str = ""
    weight_g: str = ""
    shipping_rmb: str = ""
    status: str = "失败"
    note: str = ""

    def to_row(self) -> Dict[str, str]:
        return {
            "来源Url": self.source_url,
            "offerId": self.offer_id,
            "1688商品标题": self.title,
            "1688采购价": self.price_rmb,
            "1688重量(g)": self.weight_g,
            "1688运费": self.shipping_rmb,
            "抓取状态": self.status,
            "备注": self.note,
        }


def normalize_header(value: str) -> str:
    return str(value or "").strip().lower()


def extract_offer_id(url: str) -> str:
    text = str(url or "").strip()
    if not text:
        return ""
    match = re.search(r"/offer/(\d+)\.html", text, flags=re.I)
    if match:
        return match.group(1)
    match = re.search(r"[?&](?:offerId|id)=(\d+)", text, flags=re.I)
    if match:
        return match.group(1)
    match = re.search(r"\b\d{10,}\b", text)
    return match.group(0) if match else ""


def canonical_url(url: str) -> str:
    url = str(url or "").strip()
    offer_id = extract_offer_id(url)
    if offer_id:
        return f"https://detail.1688.com/offer/{offer_id}.html"
    return re.split(r"[?#]", url)[0].rstrip("/")


def read_csv_rows(path: Path) -> List[Dict[str, str]]:
    for encoding in ("utf-8-sig", "gb18030", "gbk"):
        try:
            with path.open("r", encoding=encoding, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    with path.open("r", encoding="utf-8", errors="replace", newline="") as f:
        return list(csv.DictReader(f))


def cell_ref_to_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref or "", flags=re.I)
    if not letters:
        return 0
    total = 0
    for char in letters.group(0).upper():
        total = total * 26 + ord(char) - 64
    return total - 1


def read_xlsx_rows(path: Path) -> List[Dict[str, str]]:
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as zf:
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", ns):
                parts = [node.text or "" for node in si.findall(".//a:t", ns)]
                shared_strings.append("".join(parts))

        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in zf.namelist():
            sheets = [name for name in zf.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", name)]
            if not sheets:
                return []
            sheet_name = sheets[0]

        root = ET.fromstring(zf.read(sheet_name))
        table: List[List[str]] = []
        for row_node in root.findall(".//a:row", ns):
            row: List[str] = []
            for cell in row_node.findall("a:c", ns):
                ref = cell.attrib.get("r", "")
                col = cell_ref_to_index(ref)
                value = ""
                cell_type = cell.attrib.get("t", "")
                value_node = cell.find("a:v", ns)
                inline_node = cell.find(".//a:t", ns)
                if value_node is not None and value_node.text is not None:
                    value = value_node.text
                    if cell_type == "s":
                        idx = int(value) if value.isdigit() else -1
                        value = shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
                elif inline_node is not None and inline_node.text is not None:
                    value = inline_node.text
                while len(row) <= col:
                    row.append("")
                row[col] = value.strip()
            table.append(row)

    if not table:
        return []

    headers = [str(x or "").strip() for x in table[0]]
    # TikTok Seller Center templates place real data from row 6 onward. Normal
    # cost tables usually start from row 2. Detect template helper rows gently.
    data_start = 5 if len(table) > 5 and any("product_id" == h for h in headers) else 1
    rows: List[Dict[str, str]] = []
    for cells in table[data_start:]:
        item = {header: (cells[i].strip() if i < len(cells) else "") for i, header in enumerate(headers) if header}
        if any(str(v).strip() for v in item.values()):
            rows.append(item)
    return rows


def read_rows(path: Path) -> List[Dict[str, str]]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return read_xlsx_rows(path)
    return read_csv_rows(path)


def find_source_url(row: Dict[str, str]) -> str:
    normalized = {normalize_header(k): k for k in row.keys()}
    for field in SOURCE_URL_FIELDS:
        key = normalized.get(normalize_header(field))
        if key and str(row.get(key, "")).strip():
            return str(row[key]).strip()
    for key, value in row.items():
        if "1688.com" in str(value):
            return str(value).strip()
        lowered = normalize_header(key)
        if ("url" in lowered or "链接" in key) and str(value).strip():
            return str(value).strip()
    return ""


def extract_urls(rows: Iterable[Dict[str, str]]) -> List[str]:
    seen: set[str] = set()
    urls: List[str] = []
    for row in rows:
        url = find_source_url(row)
        if not url:
            continue
        key = canonical_url(url)
        if key in seen:
            continue
        seen.add(key)
        urls.append(url)
    return urls


def read_done(output_path: Path) -> set[str]:
    if not output_path.exists():
        return set()
    done: set[str] = set()
    for row in read_csv_rows(output_path):
        status = str(row.get("抓取状态", "")).strip()
        url = str(row.get("来源Url", "")).strip()
        if status == "成功" and url:
            done.add(canonical_url(url))
    return done


def append_result(output_path: Path, result: FetchResult) -> None:
    exists = output_path.exists() and output_path.stat().st_size > 0
    with output_path.open("a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDS)
        if not exists:
            writer.writeheader()
        writer.writerow(result.to_row())


def fetch_html(url: str, timeout: int = 25) -> str:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
            "Referer": "https://www.1688.com/",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
        content_type = resp.headers.get("Content-Type", "")
    encoding_match = re.search(r"charset=([\w-]+)", content_type, flags=re.I)
    encoding = encoding_match.group(1) if encoding_match else "utf-8"
    try:
        return raw.decode(encoding, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


def clean_page_text(raw_html: str) -> str:
    raw_html = decode_js_unicode_escapes(raw_html)
    text = re.sub(r"<script\b[\s\S]*?</script>", " ", raw_html, flags=re.I)
    text = re.sub(r"<style\b[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def decode_js_unicode_escapes(text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        try:
            return chr(int(match.group(1), 16))
        except ValueError:
            return match.group(0)

    return re.sub(r"\\u([0-9a-fA-F]{4})", replace, str(text or ""))


def pick_lowest_number(values: Iterable[str]) -> str:
    nums: List[float] = []
    for value in values:
        try:
            nums.append(float(str(value).replace(",", "")))
        except ValueError:
            continue
    if not nums:
        return ""
    lowest = min(nums)
    return str(int(lowest)) if lowest.is_integer() else f"{lowest:.2f}".rstrip("0").rstrip(".")


def extract_title(raw_html: str, page_text: str) -> str:
    raw_html = decode_js_unicode_escapes(raw_html)
    patterns = [
        r'<meta\s+property=["\']og:title["\']\s+content=["\']([^"\']+)["\']',
        r'<meta\s+name=["\']title["\']\s+content=["\']([^"\']+)["\']',
        r"<title>([\s\S]*?)</title>",
        r'"subject"\s*:\s*"([^"]+)"',
        r'"title"\s*:\s*"([^"]+)"',
    ]
    for pattern in patterns:
        match = re.search(pattern, raw_html, flags=re.I)
        if match:
            title = html.unescape(match.group(1)).strip()
            title = re.sub(r"[-_ ]*1688.*$", "", title, flags=re.I).strip()
            if title:
                return title[:200]
    return page_text[:80]


def extract_price(raw_html: str, page_text: str) -> str:
    raw_html = decode_js_unicode_escapes(raw_html)
    combined = raw_html + " " + page_text
    one_piece_patterns = [
        r"1\s*件\s*价(?:格)?[^0-9￥¥]{0,20}[￥¥]?\s*(\d+(?:\.\d+)?)",
        r"1\s*件[^0-9￥¥]{0,20}[￥¥]\s*(\d+(?:\.\d+)?)",
        r'"beginAmount"\s*:\s*"?1"?[\s\S]{0,180}?"price"\s*:\s*"?(\d+(?:\.\d+)?)"?',
    ]
    for pattern in one_piece_patterns:
        matches = re.findall(pattern, combined, flags=re.I)
        price = pick_lowest_number(matches)
        if price:
            return price

    main_price_window = re.search(r"(?:1688\s*)?价格[^0-9￥¥]{0,30}[￥¥]\s*(\d+(?:\.\d+)?)", combined, flags=re.I)
    if main_price_window:
        return pick_lowest_number([main_price_window.group(1)])

    general_patterns = [
        r"[￥¥]\s*(\d+(?:\.\d+)?)(?:\s*[-~至]\s*(\d+(?:\.\d+)?))?\s*起?",
        r'"price"\s*:\s*"?(\d+(?:\.\d+)?)"?',
        r'"priceRange"\s*:\s*"([^"]+)"',
        r'"salePrice"\s*:\s*"?(\d+(?:\.\d+)?)"?',
    ]
    candidates: List[str] = []
    for pattern in general_patterns:
        for match in re.findall(pattern, combined, flags=re.I):
            if isinstance(match, tuple):
                candidates.extend(part for part in match if part)
            else:
                candidates.extend(re.findall(r"\d+(?:\.\d+)?", str(match)))
    return pick_lowest_number(candidates)


def extract_weight_g(raw_html: str, page_text: str) -> str:
    text = decode_js_unicode_escapes(raw_html) + " " + page_text
    patterns = [
        r"(?:重量|包裹重量|包装重量|商品重量)[^0-9]{0,30}(\d+(?:\.\d+)?)\s*(kg|KG|千克|公斤|g|G|克)",
        r"(?:weight|parcel_weight|packageWeight)[^0-9]{0,30}(\d+(?:\.\d+)?)\s*(kg|KG|g|G|gram|grams)?",
    ]
    for pattern in patterns:
        for num, unit in re.findall(pattern, text, flags=re.I):
            value = float(num)
            unit = (unit or "g").lower()
            if unit in ("kg", "千克", "公斤"):
                value *= 1000
            return str(int(value)) if value.is_integer() else f"{value:.2f}".rstrip("0").rstrip(".")
    # Fallback for compact specs such as "500g" when the label is stripped by
    # the page renderer. Ignore very small values that are likely quantities.
    for num, unit in re.findall(r"\b(\d+(?:\.\d+)?)\s*(kg|KG|g|G|克|千克|公斤)\b", text):
        value = float(num)
        if value < 10 and unit.lower() in ("g", "克"):
            continue
        if unit.lower() in ("kg", "千克", "公斤"):
            value *= 1000
        return str(int(value)) if value.is_integer() else f"{value:.2f}".rstrip("0").rstrip(".")
    return ""


def extract_shipping_rmb(raw_html: str, page_text: str) -> str:
    text = decode_js_unicode_escapes(raw_html) + " " + page_text
    patterns = [
        r"(?:运费|物流费|快递费)[^0-9￥¥]{0,30}[￥¥]?\s*(\d+(?:\.\d+)?)",
        r"(?:shipping|freight)[^0-9]{0,30}(\d+(?:\.\d+)?)",
    ]
    candidates: List[str] = []
    for pattern in patterns:
        candidates.extend(re.findall(pattern, text, flags=re.I))
    shipping = pick_lowest_number(candidates)
    if shipping:
        return shipping
    # Fallback: if "运费" survives but the symbol is separated by markup, scan a
    # short window after the label.
    match = re.search(r"(?:运费|物流费|快递费)[\s\S]{0,80}", text)
    if match:
        return pick_lowest_number(re.findall(r"\d+(?:\.\d+)?", match.group(0)))
    return ""


def page_requires_login(raw_html: str, page_text: str) -> bool:
    lowered = (raw_html + " " + page_text).lower()
    markers = ["login.1688.com", "请登录", "登录后", "验证码", "captcha", "anti-bot", "访问验证"]
    return any(marker.lower() in lowered for marker in markers)


def parse_1688_page(source_url: str, raw_html: str) -> FetchResult:
    page_text = clean_page_text(raw_html)
    offer_id = extract_offer_id(source_url) or extract_offer_id(raw_html)
    if page_requires_login(raw_html, page_text):
        return FetchResult(source_url=source_url, offer_id=offer_id, status="失败", note="页面需要登录或验证")

    result = FetchResult(
        source_url=source_url,
        offer_id=offer_id,
        title=extract_title(raw_html, page_text),
        price_rmb=extract_price(raw_html, page_text),
        weight_g=extract_weight_g(raw_html, page_text),
        shipping_rmb=extract_shipping_rmb(raw_html, page_text),
        status="成功",
        note="",
    )
    missing = []
    if not result.price_rmb:
        missing.append("采购价未抓到")
    if not result.weight_g:
        missing.append("重量未抓到")
    if not result.shipping_rmb:
        result.shipping_rmb = ""
        missing.append("运费未抓到")
    if not result.price_rmb:
        result.status = "失败"
    result.note = "；".join(missing)
    return result


def fetch_one(url: str) -> FetchResult:
    canonical = canonical_url(url)
    try:
        raw_html = fetch_html(canonical)
        return parse_1688_page(canonical, raw_html)
    except urllib.error.HTTPError as error:
        return FetchResult(source_url=canonical, offer_id=extract_offer_id(canonical), status="失败", note=f"HTTP {error.code}")
    except Exception as error:
        return FetchResult(source_url=canonical, offer_id=extract_offer_id(canonical), status="失败", note=str(error)[:180])


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch 1688 cost data from TikTok/Dianxiaomi product exports.")
    parser.add_argument("input", help="Input CSV/XLSX containing 来源Url or 1688 URL.")
    parser.add_argument("-o", "--output", default="1688-cost-table.csv", help="Output CSV path.")
    parser.add_argument("--min-wait", type=float, default=1.0, help="Minimum random wait seconds between URLs.")
    parser.add_argument("--max-wait", type=float, default=3.0, help="Maximum random wait seconds between URLs.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of URLs for testing. 0 means all.")
    parser.add_argument("--force", action="store_true", help="Ignore successful rows in existing output and fetch again.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2

    rows = read_rows(input_path)
    urls = extract_urls(rows)
    if args.limit > 0:
        urls = urls[: args.limit]
    done = set() if args.force else read_done(output_path)
    pending = [url for url in urls if canonical_url(url) not in done]

    print(f"Found {len(urls)} unique 1688 URLs. Pending: {len(pending)}. Output: {output_path}")
    for index, url in enumerate(pending, start=1):
        canonical = canonical_url(url)
        print(f"[{index}/{len(pending)}] Fetching {canonical}")
        result = fetch_one(canonical)
        append_result(output_path, result)
        print(f"  -> {result.status}: price={result.price_rmb or '-'} weight={result.weight_g or '-'} shipping={result.shipping_rmb or '-'} note={result.note or '-'}")
        if index < len(pending):
            delay = random.uniform(max(0, args.min_wait), max(args.min_wait, args.max_wait))
            time.sleep(delay)

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
