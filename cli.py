"""
cli.py - コマンドライン インターフェース
ターミナルから DB 操作・SQL 実行・インポート/エクスポートができます

使い方:
    python cli.py tables                          # テーブル一覧
    python cli.py import sales.csv                # CSV インポート
    python cli.py import book.xlsx                # Excel インポート
    python cli.py sql "SELECT * FROM sales"       # SQL 実行
    python cli.py export "SELECT * FROM sales" out.csv   # CSV 出力
    python cli.py shell                           # 対話 SQL シェル
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse
from pathlib import Path
from db_engine import DBEngine


# ── カラー出力 ────────────────────────────────────────────────
class C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    GREEN  = "\033[92m"
    CYAN   = "\033[96m"
    YELLOW = "\033[93m"
    RED    = "\033[91m"
    GRAY   = "\033[90m"
    WHITE  = "\033[97m"

def ok(msg):  print(f"{C.GREEN}✓ {msg}{C.RESET}")
def err(msg): print(f"{C.RED}✗ {msg}{C.RESET}")
def info(msg):print(f"{C.CYAN}  {msg}{C.RESET}")
def head(msg):print(f"\n{C.BOLD}{C.WHITE}{msg}{C.RESET}")


# ── テーブル表示 ──────────────────────────────────────────────
def print_table(columns: list, rows: list, max_rows: int = 50):
    if not columns:
        info("（結果なし）")
        return

    # 列幅計算
    widths = [len(str(c)) for c in columns]
    for row in rows[:max_rows]:
        for i, v in enumerate(row):
            widths[i] = max(widths[i], len(str(v) if v is not None else "NULL"))
    widths = [min(w, 30) for w in widths]

    sep  = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
    def fmt_row(vals):
        cells = []
        for v, w in zip(vals, widths):
            s = str(v) if v is not None else "NULL"
            s = s[:w]
            cells.append(f" {s:<{w}} ")
        return "|" + "|".join(cells) + "|"

    print(C.GRAY + sep + C.RESET)
    print(C.BOLD + fmt_row(columns) + C.RESET)
    print(C.GRAY + sep + C.RESET)
    for row in rows[:max_rows]:
        print(fmt_row(row))
    print(C.GRAY + sep + C.RESET)

    if len(rows) > max_rows:
        print(f"{C.YELLOW}  ... 他 {len(rows) - max_rows} 行（--limit で変更可）{C.RESET}")
    print(f"{C.GRAY}  {len(rows)} 行{C.RESET}\n")


# ── コマンド実装 ──────────────────────────────────────────────
def cmd_tables(engine: DBEngine, args):
    head("📋 テーブル一覧")
    tables = engine.list_tables()
    if not tables:
        info("テーブルがありません。まず import でデータを取り込んでください。")
        return
    for t in tables:
        count = engine.row_count(t)
        cols  = engine.table_info(t)
        col_names = ", ".join(c["name"] for c in cols)
        print(f"  {C.CYAN}{t:30s}{C.RESET} {C.GRAY}{count:>7,} 行  [{col_names}]{C.RESET}")
    print()


def cmd_import(engine: DBEngine, args):
    path = Path(args.file)
    ext  = path.suffix.lower()
    tbl  = args.table
    mode = "append" if args.append else "replace"

    head(f"📥 インポート: {path.name}")

    if ext == ".csv":
        result = engine.import_csv(str(path), table_name=tbl, if_exists=mode)
        ok(f"テーブル「{result['table']}」に {result['rows']:,} 行 インポートしました")
        info(f"カラム: {', '.join(result['columns'])}")

    elif ext in (".xlsx", ".xls"):
        results = engine.import_excel(str(path), sheet_name=args.sheet,
                                      table_name=tbl, if_exists=mode)
        for r in results:
            ok(f"テーブル「{r['table']}」(シート:{r['sheet']}) に {r['rows']:,} 行")
    else:
        err(f"未対応の形式: {ext}（.csv / .xlsx のみ）")


def cmd_sql(engine: DBEngine, args):
    sql = args.query
    head(f"⚡ SQL 実行")
    print(f"  {C.GRAY}{sql}{C.RESET}\n")
    result = engine.execute_sql(sql)
    if result["type"] == "select":
        print_table(result["columns"], result["rows"],
                    max_rows=getattr(args, "limit", 50))
    else:
        ok(f"{result['affected']} 行が更新されました")


def cmd_export(engine: DBEngine, args):
    head(f"📤 エクスポート")
    out  = Path(args.output)
    ext  = out.suffix.lower()

    if ext == ".csv":
        result = engine.export_csv(args.query, str(out))
        ok(f"{result['rows']:,} 行 → {result['path']}")
    elif ext in (".xlsx", ".xls"):
        result = engine.export_excel({"Sheet1": args.query}, str(out))
        ok(f"{result['total_rows']:,} 行 → {result['path']}")
    else:
        err("出力ファイルの拡張子は .csv または .xlsx にしてください")


def cmd_drop(engine: DBEngine, args):
    engine.drop_table(args.table)
    ok(f"テーブル「{args.table}」を削除しました")


def cmd_log(engine: DBEngine, args):
    head("📜 インポート履歴")
    logs = engine.import_log()
    if not logs:
        info("履歴なし")
        return
    for l in logs:
        print(f"  {C.GRAY}#{l['id']:03d}{C.RESET}  "
              f"{C.CYAN}{l['table_name']:25s}{C.RESET}  "
              f"{l['rows']:>7,} 行  "
              f"{C.GRAY}{l['source']}{C.RESET}  "
              f"{l['imported_at'][:16]}")
    print()


# ── 対話シェル ────────────────────────────────────────────────
def cmd_shell(engine: DBEngine, args):
    print(f"\n{C.BOLD}{C.CYAN}🗄  DB Shell  (exit / quit で終了){C.RESET}")
    print(f"{C.GRAY}  DB: {engine.db_path}{C.RESET}\n")
    print("  コマンド例:")
    print(f"  {C.YELLOW}.tables{C.RESET}           テーブル一覧")
    print(f"  {C.YELLOW}.schema <table>{C.RESET}   テーブル定義")
    print(f"  {C.YELLOW}.export <file>{C.RESET}    直前の SELECT 結果を CSV/Excel に保存")
    print(f"  {C.YELLOW}SQL文;{C.RESET}            SQL を実行\n")

    last_result = None

    while True:
        try:
            line = input(f"{C.BOLD}sql>{C.RESET} ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n")
            break

        if not line:
            continue
        if line.lower() in ("exit", "quit", ".exit", ".quit"):
            break

        # ドットコマンド
        if line.startswith(".tables"):
            cmd_tables(engine, None)
            continue

        if line.startswith(".schema"):
            parts = line.split()
            tbl = parts[1] if len(parts) > 1 else None
            if tbl:
                cols = engine.table_info(tbl)
                for c in cols:
                    print(f"  {C.CYAN}{c['name']:20s}{C.RESET} {c['type']}")
            continue

        if line.startswith(".export"):
            parts = line.split(maxsplit=1)
            if len(parts) < 2 or last_result is None:
                err(".export <ファイル名>  ※直前に SELECT を実行してください")
                continue
            out_path = parts[1].strip()
            ext = Path(out_path).suffix.lower()
            try:
                if ext == ".csv":
                    # last_result を直接書き出す
                    import csv as _csv
                    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
                        w = _csv.writer(f)
                        w.writerow(last_result["columns"])
                        w.writerows(last_result["rows"])
                    ok(f"{last_result['count']:,} 行 → {out_path}")
                elif ext in (".xlsx", ".xls"):
                    # ダミーSQLではなく直接データを書く
                    try:
                        import openpyxl
                        from openpyxl.styles import Font, PatternFill
                        wb = openpyxl.Workbook()
                        ws = wb.active
                        ws.append(last_result["columns"])
                        for row in last_result["rows"]:
                            ws.append(row)
                        wb.save(out_path)
                        ok(f"{last_result['count']:,} 行 → {out_path}")
                    except ImportError:
                        err("openpyxl が必要です: pip install openpyxl")
                else:
                    err(".csv または .xlsx を指定してください")
            except Exception as e:
                err(str(e))
            continue

        # SQL 実行（複数行対応：; が来るまで続ける）
        sql = line
        while not sql.rstrip().endswith(";") and not sql.upper().lstrip().startswith("SELECT") \
              and not any(sql.upper().lstrip().startswith(k)
                          for k in ("INSERT","UPDATE","DELETE","CREATE","DROP","ALTER","WITH")):
            try:
                cont = input("   ... ").strip()
                sql += " " + cont
            except (EOFError, KeyboardInterrupt):
                break

        try:
            result = engine.execute_sql(sql)
            if result["type"] == "select":
                print_table(result["columns"], result["rows"])
                last_result = result
            else:
                ok(f"{result['affected']} 行が更新されました")
        except RuntimeError as e:
            err(str(e))


# ── メイン ────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        prog="python cli.py",
        description="📊 スプレッドシート DB マネージャー")
    parser.add_argument("--db", default=None, help="DBファイルパス（省略時はデフォルト）")
    sub = parser.add_subparsers(dest="cmd")

    # tables
    sub.add_parser("tables", help="テーブル一覧を表示")

    # import
    p_imp = sub.add_parser("import", help="CSV / Excel をインポート")
    p_imp.add_argument("file",  help="インポートするファイルパス")
    p_imp.add_argument("--table", default=None, help="テーブル名（省略時はファイル名）")
    p_imp.add_argument("--sheet", default=None, help="Excelシート名（省略時は全シート）")
    p_imp.add_argument("--append", action="store_true", help="既存テーブルに追記")

    # sql
    p_sql = sub.add_parser("sql", help="SQL を実行")
    p_sql.add_argument("query", help="SQL文")
    p_sql.add_argument("--limit", type=int, default=50, help="表示行数上限")

    # export
    p_exp = sub.add_parser("export", help="SQL結果を CSV / Excel に出力")
    p_exp.add_argument("query",  help="SELECT文")
    p_exp.add_argument("output", help="出力ファイルパス（.csv or .xlsx）")

    # drop
    p_drop = sub.add_parser("drop", help="テーブルを削除")
    p_drop.add_argument("table", help="テーブル名")

    # log
    sub.add_parser("log", help="インポート履歴を表示")

    # shell
    sub.add_parser("shell", help="対話 SQL シェルを起動")

    args = parser.parse_args()

    from db_engine import DB_PATH
    db_path = Path(args.db) if args.db else DB_PATH
    engine = DBEngine(db_path)

    try:
        cmds = {
            "tables": cmd_tables,
            "import": cmd_import,
            "sql":    cmd_sql,
            "export": cmd_export,
            "drop":   cmd_drop,
            "log":    cmd_log,
            "shell":  cmd_shell,
        }
        if args.cmd in cmds:
            cmds[args.cmd](engine, args)
        else:
            parser.print_help()
    except Exception as e:
        err(str(e))
    finally:
        engine.close()


if __name__ == "__main__":
    main()
