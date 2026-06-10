#!/usr/bin/env python3
"""Facebook 社團貼文爬蟲:攔截 GraphQL 回應,收集一個月內的貼文(post_id、精確時間、全文)。

用法:
    .venv/bin/python crawler.py            # 爬 config.json 裡所有社團
    .venv/bin/python crawler.py --group buyswitchandps   # 只爬指定社團

第一次執行會開啟瀏覽器視窗,請手動登入 Facebook(登入狀態保存在 profile/,之後免登入)。
每次執行都會把當下抓到的貼文內容以快照形式 append 到 data/posts.jsonl
(同一篇貼文若被編輯,不同執行批次會留下不同快照,可比對價格修改)。
"""
import argparse
import json
import random
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

from playwright.sync_api import sync_playwright

BASE = Path(__file__).parent
DATA = BASE / "data"
PROFILE = BASE / "profile"
CONFIG = json.loads((BASE / "config.json").read_text())

SKIP_KEYS = {"feedback", "comment_rendering_instance", "comments", "comet_ufi_summary_and_actions_renderer"}


def gather(obj, acc, depth=0):
    """在單一 Story 子樹中收集欄位(時間、內文、網址、作者)。"""
    if depth > 60:
        return
    if isinstance(obj, list):
        for x in obj:
            gather(x, acc, depth + 1)
        return
    if not isinstance(obj, dict):
        return
    ct = obj.get("creation_time")
    if isinstance(ct, int) and "creation_time" not in acc:
        acc["creation_time"] = ct
    msg = obj.get("message")
    if isinstance(msg, dict) and isinstance(msg.get("text"), str):
        if len(msg["text"]) > len(acc.get("text", "")):
            acc["text"] = msg["text"]
    www = obj.get("wwwURL")
    if isinstance(www, str) and "url" not in acc:
        acc["url"] = www
    u = obj.get("url")
    if isinstance(u, str) and "url" not in acc and ("/posts/" in u or "/permalink/" in u):
        acc["url"] = u
    actors = obj.get("actors")
    if isinstance(actors, list) and actors and isinstance(actors[0], dict):
        name = actors[0].get("name")
        if isinstance(name, str) and "author" not in acc:
            acc["author"] = name
    for k, v in obj.items():
        if k in SKIP_KEYS:
            continue
        gather(v, acc, depth + 1)


def find_stories(obj, out, depth=0):
    """遞迴尋找 __typename == Story 且帶 post_id 的節點。"""
    if depth > 60:
        return
    if isinstance(obj, list):
        for x in obj:
            find_stories(x, out, depth + 1)
        return
    if not isinstance(obj, dict):
        return
    if obj.get("__typename") == "Story" and obj.get("post_id"):
        acc = {"post_id": str(obj["post_id"])}
        gather(obj, acc)
        if acc.get("creation_time"):
            prev = out.get(acc["post_id"])
            if not prev or len(acc.get("text", "")) > len(prev.get("text", "")):
                out[acc["post_id"]] = acc
    for v in obj.values():
        find_stories(v, out, depth + 1)


def process_payload(text, out):
    for line in text.split("\n"):
        line = line.strip()
        if not line or line[0] != "{":
            continue
        try:
            find_stories(json.loads(line), out)
        except (json.JSONDecodeError, RecursionError):
            pass


def harvest_inline_json(page, out):
    """初始頁面的 SSR 資料藏在 <script type=application/json> 裡,一併解析。"""
    blobs = page.evaluate(
        "[...document.querySelectorAll('script[type=\"application/json\"]')]"
        ".map(s => s.textContent).filter(t => t && t.includes('\"Story\"'))"
    )
    for b in blobs:
        try:
            find_stories(json.loads(b), out)
        except (json.JSONDecodeError, RecursionError):
            pass


def wait_for_login(ctx, page):
    page.goto("https://www.facebook.com/", wait_until="domcontentloaded")
    def logged_in():
        return any(c["name"] == "c_user" for c in ctx.cookies("https://www.facebook.com"))
    if logged_in():
        print("[login] 已是登入狀態")
        return
    print("[login] 尚未登入 — 請在開啟的瀏覽器視窗中登入 Facebook(最多等 10 分鐘)…", flush=True)
    deadline = time.time() + 600
    while time.time() < deadline:
        if logged_in():
            print("[login] 登入成功,3 秒後開始爬取")
            time.sleep(3)
            return
        time.sleep(3)
    sys.exit("[login] 等待登入逾時,請重新執行")


def crawl_group(page, pending, group, cutoff_ts):
    posts = {}
    pending.clear()
    url = f"https://www.facebook.com/groups/{group['id']}?sorting_setting=CHRONOLOGICAL"
    print(f"[crawl] 前往 {url}")
    page.goto(url, wait_until="domcontentloaded")
    page.wait_for_timeout(6000)
    harvest_inline_json(page, posts)

    max_rounds = CONFIG["max_rounds"]
    stall_limit = CONFIG["stall_rounds_limit"]
    stall = 0
    for rnd in range(1, max_rounds + 1):
        page.mouse.wheel(0, random.randint(2200, 3200))
        page.wait_for_timeout(random.randint(1300, 2400))

        before = len(posts)
        drained, retry = list(pending), []
        pending.clear()
        for resp in drained:
            try:
                process_payload(resp.text(), posts)
            except Exception:
                retry.append(resp)
        pending.extend(retry)  # 回應還沒完成的下一輪再試

        in_window = [p for p in posts.values() if p["creation_time"] >= cutoff_ts]
        oldest = min((p["creation_time"] for p in posts.values()), default=None)
        stall = stall + 1 if len(posts) == before else 0

        if rnd % 10 == 0 or stall in (stall_limit,):
            o = datetime.fromtimestamp(oldest, tz=timezone.utc).astimezone() if oldest else None
            print(f"[crawl] round={rnd} posts={len(posts)} (月內 {len(in_window)}) oldest={o:%Y-%m-%d %H:%M} stall={stall}" if o
                  else f"[crawl] round={rnd} posts={len(posts)} stall={stall}", flush=True)
        if oldest and oldest < cutoff_ts - 86400:
            print(f"[crawl] 已爬到 cutoff 之前,停止")
            break
        if stall >= stall_limit:
            print(f"[crawl] 連續 {stall_limit} 輪無新貼文,停止(可能已到動態牆底部)")
            break
    return posts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--group", help="只爬指定社團 id")
    ap.add_argument("--days", type=int, default=CONFIG["days"])
    args = ap.parse_args()

    groups = CONFIG["groups"]
    if args.group:
        groups = [g for g in groups if g["id"] == args.group]
        if not groups:
            sys.exit(f"config.json 中找不到社團 {args.group}")

    cutoff = datetime.now(timezone.utc) - timedelta(days=args.days)
    cutoff_ts = int(cutoff.timestamp())
    scraped_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    DATA.mkdir(exist_ok=True)
    out_path = DATA / "posts.jsonl"

    with sync_playwright() as p:
        launch_kw = dict(
            user_data_dir=str(PROFILE),
            headless=CONFIG.get("headless", False),
            viewport={"width": 1280, "height": 900},
            locale="zh-TW",
            args=["--disable-blink-features=AutomationControlled"],
        )
        try:
            ctx = p.chromium.launch_persistent_context(channel="chrome", **launch_kw)
        except Exception:
            ctx = p.chromium.launch_persistent_context(**launch_kw)
        page = ctx.pages[0] if ctx.pages else ctx.new_page()

        pending = []
        page.on("response", lambda r: pending.append(r) if "/api/graphql" in r.url else None)

        wait_for_login(ctx, page)

        total = 0
        for g in groups:
            posts = crawl_group(page, pending, g, cutoff_ts)
            kept = [p for p in posts.values() if p["creation_time"] >= cutoff_ts]
            kept.sort(key=lambda p: p["creation_time"])
            with out_path.open("a", encoding="utf-8") as f:
                for prec in kept:
                    rec = {
                        "post_id": prec["post_id"],
                        "group_id": g["id"],
                        "group_name": g["name"],
                        "created_at": datetime.fromtimestamp(prec["creation_time"], tz=timezone.utc).isoformat(timespec="seconds"),
                        "creation_time": prec["creation_time"],
                        "author": prec.get("author"),
                        "text": prec.get("text", ""),
                        "url": prec.get("url") or f"https://www.facebook.com/groups/{g['id']}/posts/{prec['post_id']}/",
                        "scraped_at": scraped_at,
                    }
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            total += len(kept)
            print(f"[done] {g['name']}: 抓到 {len(posts)} 篇,寫入月內貼文 {len(kept)} 篇")
            page.wait_for_timeout(3000)
        ctx.close()

    print(f"[done] 共寫入 {total} 筆快照 → {out_path}")


if __name__ == "__main__":
    main()
