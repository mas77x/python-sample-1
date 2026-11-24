"""
Pythonサンプル1本目: 商品一覧をスクレイピングしてExcelにまとめる
題材: books.toscrape.com (スクレイピング練習用に作られた公開サイト)
  - 書籍のタイトル / 価格(£) / 評価(星の数) / 在庫を3ページ分集める
  - Excel (books.xlsx) に出力する
使い方:
  pip install requests beautifulsoup4 openpyxl
  python scrape_books_to_excel.py
注意:
  - 練習用サイト以外をスクレイピングする時は、必ず利用規約と
    robots.txt を確認し、間隔をあけて(1秒以上)優しく取得すること
  - このスクリプトは requests のタイムアウトと例外処理つき
"""

import re
import time
from datetime import date

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE_URL = "https://books.toscrape.com/catalogue/page-{}.html"
HEADERS = {"User-Agent": "Mozilla/5.0 (study-project; contact: portfolio-sample)"}
PAGES = 3  # 1ページ20件 x 3ページ = 60件


def scrape_page(page: int) -> list[dict]:
    """1ページ分の書籍一覧を dict のリストで返す。失敗時は空リスト。"""
    try:
        response = requests.get(BASE_URL.format(page), headers=HEADERS, timeout=10)
        response.raise_for_status()
        response.encoding = "utf-8"  # £ が Â£ と化けるのを防ぐ
    except requests.RequestException as exc:
        print(f"ページ {page} の取得に失敗: {exc}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    books = []
    for article in soup.select("article.product_pod"):
        title = article.h3.a["title"]
        price_text = article.select_one("p.price_color").text.strip()  # 例: £51.77
        price = float(re.search(r"[\d.]+", price_text).group())  # 数字だけ抜く
        rating_word = article.select_one("p.star-rating")["class"][1]  # One..Five
        rating = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}[rating_word]
        in_stock = "In stock" in article.select_one("p.instock").text
        books.append(
            {"title": title, "price_gbp": price, "rating": rating, "in_stock": in_stock}
        )
    return books


def main() -> None:
    all_books: list[dict] = []
    for page in range(1, PAGES + 1):
        books = scrape_page(page)
        print(f"ページ {page}: {len(books)}件取得")
        all_books.extend(books)
        time.sleep(1)  # サーバーに負荷をかけない

    print(f"合計 {len(all_books)}件")

    wb = Workbook()
    ws = wb.active
    ws.title = "books"
    headers = ["タイトル", "価格(£)", "評価(星)", "在庫あり"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for book in all_books:
        ws.append([book["title"], book["price_gbp"], book["rating"], book["in_stock"]])
    for col, width in zip("ABCD", [60, 12, 10, 10]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(all_books) + 1}"

    filename = f"books_{date.today().isoformat()}.xlsx"
    wb.save(filename)
    print(f"保存しました: {filename}")

    # 簡単な集計も表示する (ポートフォリオで見せられるポイント)
    prices = [b["price_gbp"] for b in all_books]
    print(f"平均価格: £{sum(prices) / len(prices):.2f}")
    print(f"最高価格: £{max(prices):.2f} / 最低価格: £{min(prices):.2f}")


if __name__ == "__main__":
    main()
