#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zeirishikensaku_scraper_template.py

■ 概要
- 税理士検索サイト（TOP: https://www.zeirishikensaku.jp/NzSearchContentPerson ）を
  県で検索し、一覧を抽出してExcel出力するテンプレート。
- デバッグ用に、最初の検索結果ページのHTML保存・件数ログ出力に対応。

■ 必要パッケージ
pip install requests beautifulsoup4 lxml pandas openpyxl
"""

import time
import re
import argparse
import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.zeirishikensaku.jp/NzSearchContentPerson"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
}

def fetch_page(session: requests.Session, params: dict, page: int) -> str:
    """
    都道府県などの検索パラメータとページ番号を指定してHTMLを取得。
    ※ 実サイトのパラメータ仕様に合わせて調整が必要です。
    """
    q = params.copy()
    q["page"] = page  # ←実サイトのページパラメータ名に合わせて調整
    r = session.get(BASE_URL, params=q, headers=HEADERS, timeout=20)
    r.raise_for_status()
    return r.text

def normalize_era(text: str) -> str:
    """
    平成・令和のみ抽出して '平成xx年yy月／令和aa年bb月' 形式で返す
    """
    eras = []
    for m in re.finditer(r"(平成|令和)\s*\d+年(?:\d+月)?", text):
        eras.append(m.group(0).replace(" ", ""))
    return "／".join(eras)

def extract_email(text: str) -> str:
    m = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    return m.group(0) if m else ""

def fetch_email_from_detail(session: requests.Session, url: str) -> str:
    """
    詳細ページや事務所サイトからメールアドレスを抽出（見つからなければ空文字）
    """
    if not url:
        return ""
    try:
        r = session.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
    except Exception:
        return ""
    soup = BeautifulSoup(r.text, "lxml")

    # mailto 優先
    a_mail = soup.select_one("a[href^='mailto:']")
    if a_mail:
        mail = a_mail.get("href", "").replace("mailto:", "").strip()
        if mail:
            return mail

    # テキスト全体から拾う（簡易）
    text = soup.get_text(" ", strip=True)
    email = extract_email(text)
    return email if email else ""

def parse_list(html: str) -> list[dict]:
    """
    検索結果一覧から各事務所の基本情報を抽出。
    ※ 実サイトに合わせて CSS セレクタを調整してください。
    """
    soup = BeautifulSoup(html, "lxml")
    rows = []

    # ▼ サンプルの候補セレクタ。実サイトに合わせて変更必須。
    cards = soup.select(".resultItem, .search-result-item, .listItem")
    for c in cards:
        office = c.select_one(".officeName, .name, h3")
        rep    = c.select_one(".rep, .representative, .owner")
        tel    = c.select_one(".tel, .phone")
        addr   = c.select_one(".addr, .address")
        reg    = c.select_one(".registered, .register, .reg")

        office_name = office.get_text(strip=True) if office else ""
        rep_name    = rep.get_text(strip=True) if rep else ""
        tel_text    = tel.get_text(strip=True) if tel else ""
        addr_text   = addr.get_text(strip=True) if addr else ""
        reg_text    = reg.get_text(strip=True) if reg else ""
        reg_text    = normalize_era(reg_text)

        # 詳細ページURL（あれば）
        detail_link = c.select_one("a[href]")
        detail_url = None
        if detail_link and detail_link.has_attr("href"):
            href = detail_link["href"]
            detail_url = href if href.startswith("http") else requests.compat.urljoin(BASE_URL, href)

        rows.append({
            "県": "",
            "事務所名": office_name,
            "代表者名": rep_name,
            "電話番号": tel_text,
            "メールアドレス": "",
            "住所": addr_text,
            "登録年日（平成/令和）": reg_text,
            "detail_url": detail_url
        })

    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pref", required=True, help="都道府県名（例：静岡）")
    ap.add_argument("--out", required=True,  help="出力Excelファイル名（例：静岡_税理士リスト.xlsx）")
    ap.add_argument("--delay", type=float, default=1.0, help="アクセス間隔（秒）")
    # ▼ デバッグ用オプション（今回追加）
    ap.add_argument("--debug", action="store_true", help="最初のページHTMLを保存して件数を表示")
    args = ap.parse_args()

    session = requests.Session()

    # ▼ 実サイトの検索パラメータに合わせて調整してください
    params = {
        "pref": args.pref,  # 例：'静岡'
        # 必要に応じて hidden パラメータ等を追加
    }

    all_rows = []
    page = 1

    while True:
        html = fetch_page(session, params, page=page)

        # ▼ デバッグ: 最初のページHTMLを保存
        if args.debug and page == 1:
            with open("debug_first_page.html", "w", encoding="utf-8") as f:
                f.write(html)

        rows = parse_list(html)

        # ▼ デバッグ: パース件数表示
        if args.debug:
            print(f"[DEBUG] page={page}, parsed_rows={len(rows)}")

        if not rows:
            break

        for r in rows:
            email = fetch_email_from_detail(session, r.get("detail_url"))
            r["メールアドレス"] = email if email else "記載なし"
            r["県"] = args.pref

        all_rows.extend(rows)
        page += 1
        time.sleep(args.delay)

    if not all_rows:
        print("検索結果が取得できませんでした。セレクタ/パラメータを調整してください。")
        return

    df = pd.DataFrame(all_rows)

    # 重複除去（同一事務所名＋電話番号で）
    df["key"] = df["事務所名"].fillna("") + "|" + df["電話番号"].fillna("")
    df = df.drop_duplicates(subset=["key"], keep="first").drop(columns=["key"])

    # メール有無で分割
    df_mail   = df[df["メールアドレス"] != "記載なし"].copy()
    df_nomail = df[df["メールアドレス"] == "記載なし"].copy()

    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        df_nomail.drop(columns=["detail_url"], errors="ignore").to_excel(
            w, index=False, sheet_name=f"{args.pref}_全件_メールなしのみ"
        )
        df_mail.drop(columns=["detail_url"], errors="ignore").to_excel(
            w, index=False, sheet_name=f"{args.pref}_メールあり"
        )

    print(f"Done. -> {args.out}")

if __name__ == "__main__":
    main()
