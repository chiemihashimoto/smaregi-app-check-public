"""スマレジ・アプリマーケットの最新リスト更新と新着チェック（GitHub Actions から毎週月曜に実行）

これまで手元の Python で作っていた smaregi_apps.xlsx（「有効なリンク」「全件チェック結果」）を、
同じ形式のまま自動で更新し、前回との差分（新着・掲載終了）を出します。

調べ方（2つを組み合わせます）
  1. アプリ一覧を新しい順に全ページ読み、掲載中のアプリIDを集める
     → ID の番号がどこに飛んでいても拾えます
  2. https://apps.smaregi.jp/apps/<番号> に1件ずつアクセスして、開けるか（200）を確かめる
     - 毎週（weekly）: ① 前回の一覧に出た最後の番号から前後 300件
                       ② 前回有効だったURLが、引き続き有効か
                       （今回の一覧に出た番号も確認します）
     - 全件（full）  : 1 から「最大ID＋500」まで（手動で実行したときだけ）

公開リポジトリで動かすため、次のように分けています
  ・Excel はパスワード付き（リポジトリのシークレット EXCEL_PASSWORD）。社外の人は開けません
  ・公開ページ・Issue・ログには、アプリ一覧に掲載されているアプリだけを名前付きで出します
    （URLは開けるが一覧に出ていないアプリは、件数だけ）

出力
  data/smaregi_apps.xlsx  … 有効なリンク／全件チェック結果／更新履歴（パスワード付き）
  data/new-apps.json      … ツールの「新着アプリ」タブ用
  data/market-ids.json    … 最後に確認した一覧のID
  new-apps-issue.md       … 新着があったときの通知本文
"""

from __future__ import annotations

import concurrent.futures as cf
import datetime as dt
import html
import json
import os
import re
import sys
import time
import io
import unicodedata
import urllib.error
import urllib.request
from pathlib import Path

import msoffcrypto
from msoffcrypto.format.ooxml import OOXMLFile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

BASE = os.environ.get("MARKET_BASE", "https://apps.smaregi.jp")
MODE = os.environ.get("SCAN_MODE", "weekly")      # weekly / full
RANGE = int(os.environ.get("SCAN_RANGE", "300"))  # 前回の一覧の最後の番号から前後何件見るか（weekly）
FULL_EXTRA = 500                                  # full のときに最大IDの先を何件見るか
WORKERS = 4                                       # 同時アクセス数（相手サイトへの負荷を抑える）
WAIT = float(os.environ.get("SCAN_WAIT", "0.3"))                                       # 1アクセスごとの待ち時間（秒）
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/124.0.0.0 Safari/537.36")
LAST_ERRORS: dict[str, int] = {}   # 通信エラーの内訳（失敗したときの原因調査用）
TITLE_SUFFIX = re.compile(r"\s*[｜|]\s*スマレジ・アプリマーケット\s*$")

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "smaregi_apps.xlsx"
KNOWN = ROOT / "data" / "known-apps.json"
NEWF = ROOT / "data" / "new-apps.json"
SNAP = ROOT / "data" / "market-ids.json"
ISSUE = ROOT / "new-apps-issue.md"
PASSWORD = os.environ.get("EXCEL_PASSWORD", "")

JST = dt.timezone(dt.timedelta(hours=9))
NOW = dt.datetime.now(JST)
HEAD = ["番号", "URL", "ステータス", "ページタイトル（アプリ名）", "備考"]
CATS = ["CRM・顧客管理", "Eコマース", "分析", "店内モバイルオーダー", "デリバリー・テイクアウト", "予約",
        "発注・仕入", "受注管理", "販売・売上", "決済・レジ", "勤怠・給与", "セキュリティ", "会員", "業務効率化"]


# ---------- 通信 ----------
def fetch(url: str, tries: int = 3) -> tuple[int, str]:
    """(ステータス, HTML) を返す。通信エラーは -1。"""
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "ja,en;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=30) as r:
                body = r.read().decode("utf-8", "replace")
                time.sleep(WAIT)
                return r.status, body
        except urllib.error.HTTPError as e:
            LAST_ERRORS[f"HTTP {e.code}"] = LAST_ERRORS.get(f"HTTP {e.code}", 0) + 1
            time.sleep(WAIT)
            if e.code in (404, 410):
                return e.code, ""
            if i == tries - 1:
                return e.code, ""
        except Exception as e:
            key = f"{type(e).__name__}: {str(e)[:80]}"
            LAST_ERRORS[key] = LAST_ERRORS.get(key, 0) + 1
            if i == tries - 1:
                return -1, ""
        time.sleep(3)
    return -1, ""


def to_text(s: str) -> str:
    s = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", s)
    s = re.sub(r"(?i)<br\s*/?>|</(p|div|li|h\d|a|span|dd|dt)>", "\n", s)
    s = html.unescape(re.sub(r"<[^>]+>", " ", s))
    s = re.sub(r"[ \t\u3000]+", " ", s)
    return re.sub(r"\s*\n\s*", "\n", s).strip()


def norm(s) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s or ""))).lower()


def meta(page: str, prop: str) -> str:
    m = (re.search(rf'<meta[^>]+(?:property|name)=["\']{prop}["\'][^>]+content=["\']([^"\']*)["\']', page, re.I)
         or re.search(rf'<meta[^>]+content=["\']([^"\']*)["\'][^>]+(?:property|name)=["\']{prop}["\']', page, re.I))
    return html.unescape(m.group(1)).strip() if m else ""


def page_title(page: str) -> str:
    m = re.search(r"(?is)<title>(.*?)</title>", page)
    return html.unescape(m.group(1)).strip() if m else ""


# ---------- 1. 一覧から掲載中のIDを集める ----------
CARD_RE = re.compile(r'href=["\'](?:https?://apps\.smaregi\.jp)?/apps/(\d{1,6})(?=["\'?#/])')


def cards_from(page: str) -> dict[str, str]:
    hits = [(m.group(1), m.start()) for m in CARD_RE.finditer(page)]
    out: dict[str, str] = {}
    for i, (aid, at) in enumerate(hits):
        end = hits[i + 1][1] if i + 1 < len(hits) else min(len(page), at + 6000)
        out[aid] = out.get(aid, "") + page[at:end]
    return out


def parse_card(chunk: str) -> dict:
    lines = [l for l in to_text("<a " + chunk).split("\n")
             if l and not re.search(r"の紹介画像$|^(おすすめ|NEW|\d+)$", l)]
    info = {"cat": "", "priceText": "", "trial": False, "tag": "", "min": 0}
    cat_line = next((l for l in lines if any(l.startswith(c) for c in CATS)), "")
    cats, rest = [], cat_line
    while rest:
        c = next((x for x in sorted(CATS, key=len, reverse=True) if rest.startswith(x)), None)
        if not c:
            break
        cats.append(c)
        rest = rest[len(c):]
    info["cat"] = "、".join(cats)
    pl = next((l for l in lines if re.match(r"^(無料|[\d,]+円)", l)), "")
    if pl:
        info["priceText"] = pl
        info["trial"] = bool(re.search(r"(無料体験|お試し)あり", pl))
        m = re.match(r"^([\d,]+)円", pl)
        info["min"] = int(m.group(1).replace(",", "")) if m else 0
        tail = lines[lines.index(pl) + 1:]
        tag = next((l for l in tail if len(l) >= 8 and not any(l.startswith(c) for c in CATS)), "")
        info["tag"] = tag[:120]
    return info


def crawl_listing() -> tuple[dict[str, str], int | None, int]:
    seen: dict[str, str] = {}
    total, pages = None, 0
    patterns = [
        "/apps?q=&price_division=all&page={p}&length=50&order_key=id&order=desc",
        "/apps?q=&price_division=all&page={p}&length=10&order_key=id&order=desc",
        "/apps?page={p}",
    ]
    for pat in patterns:
        for p in range(1, 41):
            st, page = fetch(BASE + pat.format(p=p))
            pages += 1
            if st != 200:
                break
            if total is None:
                m = re.search(r"(?:アプリ一覧|検索結果)\s*(\d{2,4})\s*アプリ", to_text(page).replace("\n", " "))
                total = int(m.group(1)) if m else None
            added = 0
            for aid, chunk in cards_from(page).items():
                if aid not in seen:
                    seen[aid] = chunk
                    added += 1
            if not added or (total and len(seen) >= total):
                break
        if total and len(seen) >= total:
            break
    st, top = fetch(BASE + "/")   # トップページの「新着アプリ」欄
    if st == 200:
        for aid, chunk in cards_from(top).items():
            seen.setdefault(aid, chunk)
    return seen, total, pages


# ---------- 2. 番号ごとに開けるか確かめる ----------
def probe(aid: int) -> dict:
    url = f"{BASE}/apps/{aid}"
    st, page = fetch(url)
    title = TITLE_SUFFIX.sub("", page_title(page)) if st == 200 else ""
    info = {"id": str(aid), "url": f"https://apps.smaregi.jp/apps/{aid}", "status": st, "title": title}
    if st == 200:
        info["desc"] = (meta(page, "og:description") or meta(page, "description"))[:300]
        m = re.search(r"(?:開発会社|提供会社|開発元|提供元|会社名)\s*[:：]?\s*\n?\s*([^\n]{2,40})", to_text(page))
        info["dev"] = m.group(1).strip() if m else ""
    return info


# ---------- 前回の結果 ----------
def load_prev() -> tuple[dict[str, dict], list[list]]:
    prev: dict[str, dict] = {}
    history: list[list] = []
    if not XLSX.exists():
        return prev, history
    raw = io.BytesIO(XLSX.read_bytes())
    office = msoffcrypto.OfficeFile(raw)
    if office.is_encrypted():
        office.load_key(password=PASSWORD)
        plain = io.BytesIO()
        office.decrypt(plain)
        plain.seek(0)
        raw = plain
    else:
        raw.seek(0)
    wb = load_workbook(raw, read_only=True)
    if "全件チェック結果" in wb.sheetnames:
        for r in list(wb["全件チェック結果"].iter_rows(values_only=True))[1:]:
            if r and r[0] is not None:
                prev[str(r[0])] = {"id": str(r[0]), "url": r[1], "status": int(r[2]) if str(r[2]).lstrip("-").isdigit() else r[2],
                                   "title": TITLE_SUFFIX.sub("", str(r[3] or "")), "note": r[4] or ""}
    if "更新履歴" in wb.sheetnames:
        history = [list(r) for r in list(wb["更新履歴"].iter_rows(values_only=True))[1:] if r and r[0]]
    return prev, history


def save_xlsx(rows: dict[str, dict], history: list[list]) -> None:
    wb = Workbook()
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="DDEBF7")

    def sheet(ws, data, widths, head):
        ws.append(head)
        for c in ws[1]:
            c.font, c.fill = bold, head_fill
        for r in data:
            ws.append(r)
        for col, w in zip("ABCDEFGH", widths):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    def row(r):
        t = f"{r['title']}｜スマレジ・アプリマーケット" if r.get("title") else None
        return [int(r["id"]), r["url"], r["status"], t, r.get("note") or None]

    ordered = [rows[k] for k in sorted(rows, key=int)]
    valid = [row(r) for r in ordered if r["status"] == 200]
    sheet(wb.active, valid, [8, 38, 11, 60, 45], HEAD)
    wb.active.title = "有効なリンク"
    sheet(wb.create_sheet("全件チェック結果"), [row(r) for r in ordered], [8, 38, 11, 60, 45], HEAD)
    # 見つからなくなったアプリ（公開ページには載せず、Excel にだけ残す）
    gone_rows = []
    for r in ordered:
        m = re.match(r"見つからなくなった (\d{4}-\d{2}-\d{2})（前回: (.*)）$", str(r.get("note") or ""))
        if m and r["status"] != 200:
            gone_rows.append([int(r["id"]), r["url"], m.group(2), m.group(1)])
    gone_rows.sort(key=lambda x: x[3], reverse=True)
    sheet(wb.create_sheet("見つからなくなったアプリ"), gone_rows, [8, 38, 60, 16],
          ["番号", "URL", "アプリ名（前回）", "見つからなくなった日"])
    sheet(wb.create_sheet("更新履歴"), history, [20, 10, 10, 10, 60, 60],
          ["実行日時", "方式", "有効件数", "一覧掲載件数", "新着", "見つからなくなったアプリ"])
    XLSX.parent.mkdir(parents=True, exist_ok=True)
    plain, enc = io.BytesIO(), io.BytesIO()
    wb.save(plain)
    plain.seek(0)
    OOXMLFile(plain).encrypt(PASSWORD, enc)   # パスワード付きで保存（公開リポジトリでも中身は見えない）
    XLSX.write_bytes(enc.getvalue())


def read_json(p: Path, fallback):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return fallback


# ---------- 本体 ----------
def main() -> int:
    if len(PASSWORD) < 12:
        print("❌ Excel のパスワードがありません（または12文字未満です）。リポジトリの Settings → Secrets and variables → Actions で "
              "EXCEL_PASSWORD を設定してください。", file=sys.stderr)
        return 1
    try:
        prev, history = load_prev()
    except Exception:
        print("❌ data/smaregi_apps.xlsx を開けませんでした。EXCEL_PASSWORD が前回と同じか確認してください。", file=sys.stderr)
        return 1
    prev_listed_set = {str(x) for x in read_json(SNAP, {}).get("ids", [])}
    known = {str(a["id"]) for a in read_json(KNOWN, {"apps": []})["apps"]}
    prev_valid = {k for k, v in prev.items() if v["status"] == 200}

    prev_listed = [int(x) for x in read_json(SNAP, {}).get("ids", [])]

    # 接続テスト（つながらないときは、何千件も試さずにすぐ止める）
    st, page = fetch(BASE + "/apps")
    print(f"接続テスト: {BASE}/apps → {st}（{len(page)}文字）", flush=True)
    if st != 200:
        print(f"❌ アプリマーケットに接続できませんでした（応答: {st}）。内訳: {LAST_ERRORS or 'なし'}", file=sys.stderr)
        if page:
            print("応答の先頭: " + re.sub(r"\s+", " ", page[:300]), file=sys.stderr)
        return 1

    listed, total, pages = crawl_listing()
    listed_ids = {int(x) for x in listed}
    known_max = max([int(x) for x in prev_valid] + list(listed_ids) + [int(x) for x in known] + [0])

    global MODE
    if not prev:
        MODE = "full"   # 初回は比較元がないので、1番から全件を確認する
    if MODE == "full":
        targets = set(range(1, known_max + FULL_EXTRA + 1))
    else:
        # ① 前回の一覧に出た最後の番号（初回は前回有効だった最後の番号）から前後 RANGE 件
        anchor = max(prev_listed or [int(x) for x in prev_valid] or [known_max])
        targets = set(range(max(1, anchor - RANGE), anchor + RANGE + 1))
        # ② 前回有効だったURLが、引き続き有効か
        targets |= {int(x) for x in prev_valid}
        # 今回の一覧に出た番号（番号が大きく飛んだアプリの取りこぼし防止）
        targets |= listed_ids
        print(f"基準の番号: {anchor}（{max(1, anchor - RANGE)}〜{anchor + RANGE} を確認）", flush=True)

    print(f"一覧: {len(listed)}件（掲載 {total}件 / {pages}ページ） 確認する番号: {len(targets)}件（{MODE}）", flush=True)
    results: dict[str, dict] = {}
    with cf.ThreadPoolExecutor(WORKERS) as ex:
        for info in ex.map(probe, sorted(targets)):
            results[info["id"]] = info

    if not any(r["status"] == 200 for r in results.values()):
        statuses: dict[str, int] = {}
        for r in results.values():
            statuses[str(r["status"])] = statuses.get(str(r["status"]), 0) + 1
        print(f"❌ 有効なアプリが1件もありませんでした。応答の内訳: {statuses}　通信エラーの内訳: {LAST_ERRORS or 'なし'}",
              file=sys.stderr)
        return 1

    # 通信エラーの番号は、前回の結果を残す
    merged = dict(prev)
    for k, r in results.items():
        if r["status"] == -1 and k in prev:
            merged[k] = dict(prev[k], note="今回は取得エラー（前回の結果）")
        else:
            old = (prev.get(k) or {}).get("note", "")
            keep = re.findall(r"(?:新着|見つからなくなった) \d{4}-\d{2}-\d{2}[^／]*", old)
            merged[k] = dict(r, keep=keep)

    now_valid = {k for k, v in merged.items() if v["status"] == 200}
    if prev:
        new_ids = sorted(now_valid - prev_valid, key=int)
    else:
        # 初回（比較元の Excel がない）: 一覧に出ていて、社内ツールに未登録のアプリを新着とする
        new_ids = sorted((now_valid & set(listed)) - known, key=int)
    gone_ids = sorted(prev_valid - now_valid, key=int)

    for k in now_valid:
        notes = []
        if listed and k not in listed:
            notes.append("一覧に掲載なし（URLは有効）")
        if k in new_ids:
            notes.append(f"新着 {NOW:%Y-%m-%d}")
        else:
            notes += [x for x in merged[k].get("keep", []) if x.startswith("新着")]
        if merged[k].get("note", "").startswith("今回は取得エラー"):
            notes.insert(0, merged[k]["note"])
        merged[k]["note"] = "／".join(notes)
    for k, v in merged.items():
        if k in gone_ids:
            v["note"] = f"見つからなくなった {NOW:%Y-%m-%d}（前回: {prev[k]['title']}）"
        elif v["status"] != 200 and v.get("keep"):
            v["note"] = "／".join(x for x in v["keep"] if x.startswith("見つからなくなった"))

    history.append([f"{NOW:%Y-%m-%d %H:%M}", MODE, len(now_valid), total,
                    "、".join(merged[k]["title"] for k in new_ids) or "—",
                    "、".join(prev[k]["title"] for k in gone_ids) or "—"])
    save_xlsx(merged, history)

    # ツール用の新着データ
    store = read_json(NEWF, {"apps": []})
    store.setdefault("apps", [])
    have = {str(a["id"]) for a in store["apps"]}
    # 公開してよいのは、アプリ一覧に掲載されている（いた）アプリだけ
    pub_new = [k for k in new_ids if k in listed]
    pub_gone = [k for k in gone_ids if k in prev_listed_set or k in known]
    found = []
    for k in pub_new:
        if k in have or k in known:
            continue
        r, card = merged[k], parse_card(listed.get(k, ""))
        item = {"id": k, "name": r["title"] or f"ID {k}", "dev": r.get("dev", ""), "cat": card["cat"],
                "priceText": card["priceText"], "min": card["min"], "max": card["min"], "trial": card["trial"],
                "tag": card["tag"], "desc": r.get("desc", ""), "url": r["url"], "foundAt": NOW.isoformat(),
                "listed": k in listed}
        store["apps"].append(item)
        found.append(item)
    # 見つからなくなったアプリは、公開ページにはアプリ名と日付だけを出す（番号・URLは Excel にだけ残す）
    # 一覧に掲載されていなかったアプリは、名前も出さない
    store.setdefault("gone", [])
    for k in pub_gone:
        store["gone"].append({"name": prev[k]["title"], "goneAt": NOW.isoformat()})
    # 新着として公開したあとで見つからなくなったアプリは、公開ページから外す
    store["apps"] = [a for a in store["apps"] if str(a["id"]) not in set(gone_ids)]
    store.update(checkedAt=NOW.isoformat(), marketTotal=total, seenCount=len(listed),
                 lastRun={"at": NOW.isoformat(), "mode": MODE, "newIds": pub_new})
    NEWF.write_text(json.dumps(store, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    SNAP.write_text(json.dumps({"checkedAt": NOW.isoformat(), "total": total, "ids": sorted(listed, key=int)},
                               ensure_ascii=False, indent=1) + "\n", encoding="utf-8")

    # まとめ（Issue・ログ・Actions の画面は公開されるので、一覧に掲載されているアプリだけ名前を出す）
    unlisted = sorted((k for k in now_valid if listed and k not in listed), key=int)
    priv_new, priv_gone = len(new_ids) - len(pub_new), len(gone_ids) - len(pub_gone)
    md = [f"## アプリマーケット新着チェック（{NOW:%Y-%m-%d %H:%M}・{'全件' if MODE == 'full' else '毎週'}）", "",
          f"- 一覧の掲載: {total if total else '不明'}件（確認できた {len(listed)}件）",
          f"- 新着: {len(pub_new)}件　見つからなくなった: {len(gone_ids)}件（名前は Excel の「見つからなくなったアプリ」シート）", ""]
    if pub_new:
        md += ["### 新着", "| アプリ | カテゴリ | 料金 |", "|---|---|---|"]
        for k in pub_new:
            card = parse_card(listed.get(k, ""))
            md.append(f"| [{merged[k]['title']}]({merged[k]['url']}) | {card['cat'] or '—'} | {card['priceText'] or '—'} |")
        md.append("")
    if total and len(listed) < total:
        md.append(f"⚠ 一覧は掲載{total}件のうち{len(listed)}件しか確認できませんでした。ページ送りの形式が変わった可能性があります。")
    md += ["", f"一覧に出ていないアプリ（URLは有効）: {len(unlisted)}件"
           + (f"（今回増えた {priv_new}件）" if priv_new else "")
           + "。名前は Excel（パスワード付き）で確認してください。"]
    text = "\n".join(md) + "\n"
    print(text)
    ISSUE.write_text(text, encoding="utf-8")
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        with open(os.environ["GITHUB_STEP_SUMMARY"], "a", encoding="utf-8") as f:
            f.write(text)
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a", encoding="utf-8") as f:
            f.write(f"new_count={len(pub_new)}\ngone_count={len(gone_ids)}\nother_count={priv_new}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
