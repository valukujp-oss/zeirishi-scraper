#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ブラウザを自動で動かして検索サイトを開く
# 結果をHTMLとスクリーンショットで保存する

from playwright.sync_api import sync_playwright

TOP = "https://www.zeirishikensaku.jp/NzSearchContentPerson"

def run(pref="静岡"):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(TOP, wait_until="networkidle")
        
        # ページのHTMLを保存
        html = page.content()
        with open("playwright_first_page.html", "w", encoding="utf-8") as f:
            f.write(html)
        
        # スクリーンショットも保存
        page.screenshot(path="playwright_first_page.png", full_page=True)

        browser.close()

if __name__ == "__main__":
    run("静岡")
