"""
Power Apps モデル駆動型アプリの画面を自動キャプチャするスクリプト

使い方:
  1. 初回: python scripts/capture_screenshots.py --login
     → ブラウザが開くので手動でサインイン → 認証状態が保存される
  2. 2回目以降: python scripts/capture_screenshots.py
     → 保存済み認証で自動キャプチャ

出力先: docs/images/
"""
import argparse
import os
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

# 設定
from dotenv import load_dotenv
load_dotenv()

DATAVERSE_URL = os.environ.get("DATAVERSE_URL", "").rstrip("/")
APP_MODULE_ID = os.environ.get("APP_MODULE_ID", "")
APP_URL = f"{DATAVERSE_URL}/main.aspx?appid={APP_MODULE_ID}"

PROJECT_ROOT = Path(__file__).resolve().parent.parent
AUTH_STATE = PROJECT_ROOT / "playwright" / ".auth" / "state.json"
OUTPUT_DIR = PROJECT_ROOT / "docs" / "images"


def ensure_dirs():
    AUTH_STATE.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def login_and_save_state():
    """ブラウザを開いて手動サインイン → 認証状態を保存"""
    print("=== 手動サインイン ===")
    print("ブラウザが開きます。Power Apps にサインインしてください。")
    print("サインイン完了後、アプリが表示されたらブラウザを閉じてください。")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Power Apps にアクセス（ログインページにリダイレクトされる）
        page.goto(APP_URL)

        # ユーザーが手動でサインインするのを待つ
        print("\nサインインを待機中...")
        print("アプリが完全に読み込まれたら Enter を押してください。")
        input(">>> Enter を押す: ")

        # 認証状態を保存
        context.storage_state(path=str(AUTH_STATE))
        print(f"\n✅ 認証状態を保存: {AUTH_STATE}")

        browser.close()


def capture_screenshots():
    """保存済み認証で各画面のスクリーンショットを撮影"""
    if not AUTH_STATE.exists():
        print("❌ 認証状態が見つかりません。先に --login で認証してください。")
        print("  python scripts/capture_screenshots.py --login")
        sys.exit(1)

    print(f"=== スクリーンショット撮影 ===")
    print(f"  アプリ URL: {APP_URL}")
    print(f"  出力先: {OUTPUT_DIR}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # headed で撮影（UI が正しくレンダリングされる）
        context = browser.new_context(
            storage_state=str(AUTH_STATE),
            viewport={"width": 1920, "height": 1080},
        )
        page = context.new_page()

        def wait_and_screenshot(name, wait_sec=3):
            """指定秒待ってからスクリーンショット"""
            time.sleep(wait_sec)
            path = OUTPUT_DIR / f"{name}.png"
            page.screenshot(path=str(path), full_page=False)
            print(f"  ✅ {name}.png")

        # 1. アプリのメイン画面（チケット一覧）
        print("\n--- チケット一覧 ---")
        page.goto(APP_URL, wait_until="domcontentloaded", timeout=60000)
        time.sleep(10)  # Power Apps SPA の読み込み待ち
        wait_and_screenshot("01_ticket_list", 3)

        # 2. ナビゲーション（左サイドバー）
        print("\n--- ナビゲーション ---")
        wait_and_screenshot("02_navigation", 1)

        # 3. チャート表示
        print("\n--- チャート ---")
        try:
            # グラフアイコンをクリック
            chart_btn = page.locator('[aria-label*="グラフ"], [aria-label*="chart"], [title*="グラフ"]').first
            if chart_btn.is_visible():
                chart_btn.click()
                wait_and_screenshot("03_chart", 3)
            else:
                print("  ⚠️ チャートボタンが見つかりません")
                wait_and_screenshot("03_chart", 1)
        except Exception as e:
            print(f"  ⚠️ チャート: {e}")
            wait_and_screenshot("03_chart", 1)

        # 4. チケット新規作成画面
        print("\n--- チケット新規作成 ---")
        try:
            new_btn = page.locator('[aria-label*="新規"], [aria-label*="New"]').first
            if new_btn.is_visible():
                new_btn.click()
                time.sleep(5)
                wait_and_screenshot("04_ticket_new", 3)
            else:
                print("  ⚠️ 新規ボタンが見つかりません")
        except Exception as e:
            print(f"  ⚠️ 新規作成: {e}")

        # 5. チケット詳細（既存レコード）
        print("\n--- チケット詳細 ---")
        page.goto(APP_URL, wait_until="domcontentloaded", timeout=60000)
        time.sleep(8)
        try:
            # 最初の行をクリック
            first_row = page.locator('[data-lp-id*="MscrmControls.Grid"] div[role="row"]').nth(1)
            if first_row.is_visible():
                first_row.click()
                time.sleep(5)
                wait_and_screenshot("05_ticket_detail_general", 3)

                # 対応タブ
                try:
                    response_tab = page.locator('li[role="tab"]:has-text("対応"), li[role="tab"]:has-text("Response")')
                    if response_tab.first.is_visible():
                        response_tab.first.click()
                        wait_and_screenshot("06_ticket_detail_response", 2)
                except:
                    print("  ⚠️ 対応タブが見つかりません")
            else:
                print("  ⚠️ チケット行が見つかりません")
        except Exception as e:
            print(f"  ⚠️ チケット詳細: {e}")

        # 6. ナレッジ記事一覧
        print("\n--- ナレッジ記事 ---")
        page.goto(APP_URL, wait_until="domcontentloaded", timeout=60000)
        time.sleep(8)
        try:
            kb_link = page.locator('a:has-text("ナレッジ記事"), span:has-text("ナレッジ記事")').first
            if kb_link.is_visible():
                kb_link.click()
                time.sleep(5)
                wait_and_screenshot("07_knowledge_list", 3)
        except Exception as e:
            print(f"  ⚠️ ナレッジ: {e}")

        # 7. カテゴリ一覧
        print("\n--- カテゴリ ---")
        try:
            cat_link = page.locator('a:has-text("カテゴリ"), span:has-text("カテゴリ")').first
            if cat_link.is_visible():
                cat_link.click()
                time.sleep(5)
                wait_and_screenshot("08_category_list", 3)
        except Exception as e:
            print(f"  ⚠️ カテゴリ: {e}")

        browser.close()

    print(f"\n=== 撮影完了 ===")
    print(f"  出力先: {OUTPUT_DIR}")
    for f in sorted(OUTPUT_DIR.glob("*.png")):
        print(f"  {f.name}")


def main():
    parser = argparse.ArgumentParser(description="Power Apps スクリーンショット自動撮影")
    parser.add_argument("--login", action="store_true", help="手動サインインして認証状態を保存")
    args = parser.parse_args()

    ensure_dirs()

    if args.login:
        login_and_save_state()
    else:
        capture_screenshots()


if __name__ == "__main__":
    main()
