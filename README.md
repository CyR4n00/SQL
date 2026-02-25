# 🗄 DB マネージャー

Excel / CSV ファイルを SQLite データベースにまとめて管理し、
SQL で自由に加工して再度 CSV / Excel に書き出せるツールです。

---

## 🚀 起動方法

### GUI アプリ（画面で操作）
```
起動.bat をダブルクリック
```

### コマンドライン（ターミナルで操作）
```powershell
python cli.py --help
```

---

## 📥 インポート

### GUI
1. 左メニューの「📥 CSV インポート」または「📥 Excel インポート」をクリック
2. ファイルを選択 → 自動でテーブルが作成される

### CLI
```powershell
# CSV をインポート（テーブル名はファイル名から自動生成）
python cli.py import 売上データ.csv

# テーブル名を指定
python cli.py import 売上データ.csv --table sales

# Excel（全シートをインポート）
python cli.py import データ.xlsx

# 特定シートのみ
python cli.py import データ.xlsx --sheet Sheet1

# 既存テーブルに追記
python cli.py import 追加データ.csv --table sales --append
```

---

## ⚡ SQL 実行

### GUI
「⚡ SQL」タブにSQL文を入力して Ctrl+Enter

### CLI
```powershell
# SELECT
python cli.py sql "SELECT * FROM sales LIMIT 10"

# 集計
python cli.py sql "SELECT 商品名, SUM(売上) as 合計 FROM sales GROUP BY 商品名"

# テーブル結合
python cli.py sql "SELECT a.*, b.地域 FROM sales a JOIN regions b ON a.店舗ID = b.店舗ID"

# 対話シェル（複数のSQL を連続実行）
python cli.py shell
```

---

## 📤 エクスポート

### GUI
- SQL タブで実行後 → 「💾 CSV」または「💾 Excel」ボタン
- または「📤 エクスポート」タブで SQL を指定して出力

### CLI
```powershell
# CSV に出力
python cli.py export "SELECT * FROM sales" 出力.csv

# Excel に出力
python cli.py export "SELECT 商品名, SUM(売上) FROM sales GROUP BY 商品名" 集計.xlsx
```

---

## 📋 テーブル管理

```powershell
# 一覧表示
python cli.py tables

# テーブルを削除
python cli.py drop テーブル名

# インポート履歴
python cli.py log
```

---

## 📁 ファイル構成

```
db_manager/
├── 起動.bat          ← GUI 起動（ダブルクリック）
├── app.py            # GUI アプリ
├── cli.py            # コマンドライン
├── db_engine.py      # DB エンジン（コア）
├── requirements.txt  # 必要パッケージ
├── data/
│   └── database.db   # SQLite DB（自動生成）
└── README.md
```

---

## 動作環境

| 項目 | 要件 |
|------|------|
| OS | Windows 10/11 |
| Python | 3.8 以上 |
| パッケージ | openpyxl（起動.bat で自動インストール）|
