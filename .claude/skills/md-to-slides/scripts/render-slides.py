#!/usr/bin/env python3
"""PlaywrightでスライドHTMLをPNG画像に変換する"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

def render(html_dir: str, output_dir: str) -> None:
    html_dir = Path(html_dir)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)  # 出力フォルダがなければ作成

    # output/images/slides/ 以下の slide-*.html を名前順に取得
    targets = sorted(html_dir.glob("slide-*.html"))
    print(f"変換対象: {len(targets)} ファイル")

    with sync_playwright() as p:
        browser = p.chromium.launch()  # ヘッドレスブラウザを起動（画面は表示されない）
        for html_path in targets:
            # スライドサイズ（960×540px）に合わせたブラウザウィンドウを作成
            page = browser.new_page(viewport={"width": 960, "height": 540})
            page.set_content(html_path.read_text(encoding="utf-8"))  # HTMLを読み込む
            page.wait_for_timeout(200)  # フォントやレイアウトの描画が完了するまで少し待つ
            out = output_dir / f"{html_path.stem}.png"  # 例: slide-01.html → slide-01.png
            page.screenshot(path=str(out))  # スクリーンショットを撮影してPNGとして保存
            page.close()
            print(f"  OK {html_path.name} -> {out.name}")
        browser.close()

    print("完了")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("使用方法: python render-slides.py <HTMLフォルダ> <出力フォルダ>")
        sys.exit(1)
    render(sys.argv[1], sys.argv[2])
