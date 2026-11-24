python-sample-1
書籍一覧サイト（練習用）から、タイトル・価格・評価・在庫情報をスクレイピングしてexcelファイルに出力するpythonスクリプトです

入力
pip install requests beautifulsoup4 openpyxl
python scrape_books_to_excel.py

出力
・books_YYYY-MM-DD.xlsxが生成されます（60件分）
・列：タイトル/価格(￡)/評価(星の数)/在庫あり
・見出しの太字・フィルター・ウィンドウ枠固定付き

使っている技術
・requests/BeautifulSoupによるスクレイピング（取得間隔1秒、タイムアウトと例外処理あり）
・openpyxlによるExcel生成
