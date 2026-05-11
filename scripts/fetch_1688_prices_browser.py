#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Browser-login version of the 1688 cost fetcher.

Use this when the standard urllib script sees "login required" or verification
pages. This script opens a real Chromium browser with a persistent profile. You
log in manually once, then it reuses that browser session to fetch offer pages.

It does not solve captchas or bypass verification. If 1688 asks for manual
verification, complete it in the browser and then press Enter in the terminal.
"""

from __future__ import annotations

import argparse
import random
import sys
import time
from pathlib import Path

try:
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
    from playwright.sync_api import sync_playwright
except ImportError:  # pragma: no cover - user environment check
    print(
        "缺少 Playwright。请先执行：\n"
        "  pip install playwright\n"
        "  python -m playwright install chromium",
        file=sys.stderr,
    )
    raise SystemExit(2)

from fetch_1688_prices import (
    append_result,
    canonical_url,
    extract_urls,
    parse_1688_page,
    read_done,
    read_rows,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch 1688 costs with a logged-in browser session.")
    parser.add_argument("input", nargs="?", help="Input CSV/XLSX containing 来源Url or 1688 URL.")
    parser.add_argument("-o", "--output", default="1688-cost-table.csv", help="Output CSV path.")
    parser.add_argument("--profile-dir", default=".1688-browser-profile", help="Persistent browser profile directory.")
    parser.add_argument("--min-wait", type=float, default=1.0, help="Minimum random wait seconds between URLs.")
    parser.add_argument("--max-wait", type=float, default=3.0, help="Maximum random wait seconds between URLs.")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of URLs for testing. 0 means all.")
    parser.add_argument("--force", action="store_true", help="Ignore successful rows in existing output and fetch again.")
    parser.add_argument("--headless", action="store_true", help="Run browser headless. Not recommended for first login.")
    parser.add_argument("--skip-login-prompt", action="store_true", help="Do not wait for manual login before fetching.")
    parser.add_argument("--login-only", action="store_true", help="Only open the 1688 login page and save browser session.")
    parser.add_argument("--channel", choices=["chromium", "chrome", "msedge"], default="chromium", help="Browser channel to use.")
    return parser.parse_args()


def wait_for_login(page) -> None:
    login_url = "https://login.1688.com/member/signin.htm?Done=https%3A%2F%2Fwww.1688.com%2F"
    page.goto(login_url, wait_until="domcontentloaded", timeout=60000)
    try:
        page.bring_to_front()
    except Exception:
        pass
    print("\n浏览器已打开 1688 登录页。")
    print("如果你没看到浏览器窗口，请检查任务栏里是否有 Chromium / Chrome / Edge 图标。")
    print("请在浏览器里完成 1688 登录 / 验证。")
    print("确认已经登录后，回到这个命令行窗口按 Enter。")
    input("按 Enter 继续...")


def launch_context(playwright, args):
    browser_type = playwright.chromium
    launch_kwargs = {
        "user_data_dir": args.profile_dir,
        "headless": args.headless,
        "viewport": None,
        "locale": "zh-CN",
        "args": [
            "--start-maximized",
            "--disable-blink-features=AutomationControlled",
        ],
    }
    if args.channel != "chromium":
        launch_kwargs["channel"] = args.channel
    return browser_type.launch_persistent_context(**launch_kwargs)


def fetch_with_browser(page, url: str):
    canonical = canonical_url(url)
    try:
        page.goto(canonical, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(2500)
        try:
            page.mouse.wheel(0, 900)
            page.wait_for_timeout(1200)
        except Exception:
            pass
        html = page.content()
        try:
            body_text = page.locator("body").inner_text(timeout=5000)
        except Exception:
            body_text = ""
        return parse_1688_page(canonical, html + "\n" + body_text)
    except PlaywrightTimeoutError:
        from fetch_1688_prices import FetchResult, extract_offer_id

        return FetchResult(source_url=canonical, offer_id=extract_offer_id(canonical), status="失败", note="页面加载超时")
    except Exception as error:
        from fetch_1688_prices import FetchResult, extract_offer_id

        return FetchResult(source_url=canonical, offer_id=extract_offer_id(canonical), status="失败", note=str(error)[:180])


def main() -> int:
    args = parse_args()
    if not args.input and not args.login_only:
        print("请提供输入表格文件，或使用 --login-only 先打开浏览器登录 1688。", file=sys.stderr)
        return 2

    if args.login_only:
        with sync_playwright() as p:
            context = launch_context(p, args)
            page = context.pages[0] if context.pages else context.new_page()
            wait_for_login(page)
            print("登录态已保存。之后运行抓取命令会复用这个浏览器登录态。")
            context.close()
        return 0

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
    if not pending:
        return 0

    with sync_playwright() as p:
        context = launch_context(p, args)
        page = context.pages[0] if context.pages else context.new_page()
        if not args.skip_login_prompt and not args.headless:
            wait_for_login(page)

        for index, url in enumerate(pending, start=1):
            canonical = canonical_url(url)
            print(f"[{index}/{len(pending)}] Fetching {canonical}")
            result = fetch_with_browser(page, canonical)
            append_result(output_path, result)
            print(
                f"  -> {result.status}: price={result.price_rmb or '-'} "
                f"weight={result.weight_g or '-'} shipping={result.shipping_rmb or '-'} "
                f"note={result.note or '-'}"
            )

            if result.status == "失败" and ("登录" in result.note or "验证" in result.note):
                print("检测到登录/验证页面。请在浏览器里处理后按 Enter 继续。")
                input("按 Enter 继续...")

            if index < len(pending):
                delay = random.uniform(max(0, args.min_wait), max(args.min_wait, args.max_wait))
                time.sleep(delay)

        context.close()

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
