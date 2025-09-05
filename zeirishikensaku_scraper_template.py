#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
zeirishikensaku_scraper_template.py

■ 概要
- 税理士検索サイトで都道府県を指定して一覧を取得し、
  メールあり/なしで分けてExcelに保存するテンプレートです。

■ 必要な環境
- Python 3.9+
- pip install requests beautifulsoup4 lxml pandas openpyxl

■ 使い方（GitHub Actionsやローカルで実行）
python zeirishikensaku_scraper_template.py --pref 静岡 --out 静岡_税理士リスト.xlsx --delay 1.2

■ 注意
- 対象サイトの利用規約・robots.txtを必ず守ってください。
- 実際のHTML構造に合わせてCSSセレクタを調整する必要があります。
"""

import time
import re
import argparse
import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.zeirishikensaku.jp/NzSearchContentPerson"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0.0.0 Safari/537.36"
}

def fetch_page(session: requests.Session, params: dict, page: int) -> str:
    """検索ページのHTMLを取得"""
    q = params.copy()
    q["page"] = page
    resp = session.get(BASE_URL, params=q, headers=HEADERS, timeout=20)
    resp.raise_for_status()
    return resp.text

def parse_list(html: str) -> list[dict]:
    """検索結果一覧から情報を抽出"""
    soup = BeautifulSoup(html, "lxml")
    rows = []
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

        reg_text = normalize_era(reg_text)

        detail_link = c.select_one("a[href]")
        detail_url = None
        if detail_link and detail_link.has_attr("href"):
            href = detail_link["href"]
            if href.startswith("http"):
                detail_url = href
            else:
                detail_url = requests.compat.urljoin(BASE_URL, href)

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

def normalize_era(text: str) -> str:
    """平成・令和のみを抽出"""
    eras = []
    for m in re.finditer(r"(平成|令和)\s*\d+年(?:\d+月)?", text):
        eras.append(m.group(0).replace(" ", ""))
    return "／".join(eras)

def fetch_email_from_detail(session: requests.Session, url: str) -> str:
    """詳細ページからメールを探す"""
    if not url:
        return ""
    try:
        r = session.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
    except Exception:
        return ""
    soup = BeautifulSoup(r.text, "lxml")

    a_mail = soup.select_one("a[href^='mailto:']")
    if a_mail:
        mail = a_mail.get("href", "").replace("mailto:", "").strip()
        if mail:
            return mail

    text = soup.get_text(" ", strip=True)
    email = extract_email(text)
    return email if email else ""

def extract_email(text: str) -> str:
    m = re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    return m.group(0) if m else ""

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pref", required=True, help="都道府県名（例：静岡）")
    ap.add_argument("--out", required=True, help="出力Excelファイル名")
    ap.add_argument("--delay", type=float, default=1.0, help="アクセス間隔（秒）")
    args = ap.parse_args()

    session = requests.Session()
    params = {"pref": args.pref}

    all_rows = []
    page = 1

    while True:
        html = fetch_page(session, params, page=page)
        rows = parse_list(html)
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
        print("検索結果が取得できませんでした。")
        return

    df = pd.DataFrame(all_rows)

    df_mail = df[df["メールアドレス"] != "記載なし"].copy()
    df_mail = df_mail.drop_duplicates(subset=["事務所名"], keep="first")
    df_nomail = df[df["メールアドレス"] == "記載なし"].copy()

    with pd.ExcelWriter(args.out, engine="openpyxl") as writer:
        df_nomail.drop(columns=["detail_url"], errors="ignore").to_excel(
            writer, index=False, sheet_name=f"{args.pref}_全件_メールなしのみ"
        )
        df_mail.drop(columns=["detail_url"], errors="ignore").to_excel(
            writer, index=False, sheet_name=f"{args.pref}_メールあり"
        )

    print(f"Done. -> {args.out}")

if __name__ == "__main__":
    main()
