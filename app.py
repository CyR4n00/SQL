"""
app.py - GUI データベースマネージャー
tkinter ベースのデスクトップアプリ
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
from db_engine import DBEngine, DB_PATH

# ── カラーパレット ────────────────────────────────────────────
C = {
    "bg":       "#0f1117",
    "surface":  "#1a1d27",
    "surface2": "#22263a",
    "border":   "#2e3250",
    "accent":   "#5b8dee",
    "green":    "#4ade80",
    "red":      "#f87171",
    "gold":     "#fbbf24",
    "text":     "#e2e8f0",
    "muted":    "#64748b",
    "header_bg":"#1e3a5f",
}

FONT_MAIN  = ("Meiryo UI", 10)
FONT_BOLD  = ("Meiryo UI", 10, "bold")
FONT_MONO  = ("Courier New", 10)
FONT_SMALL = ("Meiryo UI", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🗄 DB マネージャー")
        self.geometry("1300x820")
        self.minsize(1000, 600)
        self.configure(bg=C["bg"])

        self.engine = DBEngine()
        self._build()
        self._refresh_tables()

    # ── UI 構築 ───────────────────────────────────────────────
    def _build(self):
        # ── ヘッダー
        hdr = tk.Frame(self, bg=C["surface"], pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🗄 DB マネージャー",
                 font=("Meiryo UI", 14, "bold"),
                 bg=C["surface"], fg=C["accent"]).pack(side="left", padx=20)
        tk.Label(hdr, text=str(self.engine.db_path),
                 font=FONT_SMALL, bg=C["surface"], fg=C["muted"]).pack(side="left")

        # DB 切り替えボタン
        tk.Button(hdr, text="📂 DB を開く",
                  command=self._open_db,
                  bg=C["surface2"], fg=C["text"],
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=10, pady=4).pack(side="right", padx=16)

        # ── メイン（左: テーブルパネル、右: 作業エリア）
        paned = tk.PanedWindow(self, orient="horizontal",
                               bg=C["border"], sashwidth=4)
        paned.pack(fill="both", expand=True, padx=0, pady=0)

        # 左パネル
        left = tk.Frame(paned, bg=C["surface"], width=220)
        paned.add(left, minsize=180)
        self._build_left(left)

        # 右パネル（タブ）
        right = tk.Frame(paned, bg=C["bg"])
        paned.add(right, minsize=600)
        self._build_right(right)

        # ── ステータスバー
        self.status_var = tk.StringVar(value="準備完了")
        tk.Label(self, textvariable=self.status_var,
                 bg=C["surface2"], fg=C["muted"],
                 font=("Courier New", 9), anchor="w", padx=12
                 ).pack(fill="x", side="bottom")

    def _build_left(self, parent):
        """左：テーブルツリー & インポートボタン"""
        # タイトル
        tk.Label(parent, text="TABLES", font=("Courier New", 9, "bold"),
                 bg=C["surface"], fg=C["muted"],
                 pady=8, padx=12, anchor="w").pack(fill="x")

        # テーブルリスト
        frame = tk.Frame(parent, bg=C["surface"])
        frame.pack(fill="both", expand=True, padx=8)

        sb = tk.Scrollbar(frame, orient="vertical")
        self.table_list = tk.Listbox(frame,
                                     bg=C["surface2"], fg=C["text"],
                                     font=FONT_MAIN,
                                     selectbackground=C["accent"],
                                     selectforeground="#fff",
                                     relief="flat", bd=0,
                                     yscrollcommand=sb.set)
        sb.config(command=self.table_list.yview)
        sb.pack(side="right", fill="y")
        self.table_list.pack(fill="both", expand=True)
        self.table_list.bind("<<ListboxSelect>>", self._on_table_select)
        self.table_list.bind("<Double-Button-1>", self._preview_table)

        # ボタン群
        btn_frame = tk.Frame(parent, bg=C["surface"])
        btn_frame.pack(fill="x", padx=8, pady=8)

        def btn(text, cmd, fg=C["text"]):
            b = tk.Button(btn_frame, text=text, command=cmd,
                          bg=C["surface2"], fg=fg,
                          font=FONT_SMALL, relief="flat", cursor="hand2",
                          padx=8, pady=5, anchor="w")
            b.pack(fill="x", pady=2)
            return b

        btn("📥 CSV インポート",           self._import_csv,       C["green"])
        btn("📥 Excel インポート",        self._import_excel,     C["green"])
        btn("📊 Google スプレッドシート", self._open_gsheets_tab, C["gold"])
        btn("🗑 テーブル削除",            self._drop_table,       C["red"])
        btn("🔄 更新",                    self._refresh_tables)

    def _build_right(self, parent):
        """右：タブ（SQL / プレビュー / エクスポート / ログ）"""
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook",       background=C["bg"],     borderwidth=0)
        style.configure("TNotebook.Tab",   background=C["surface2"], foreground=C["muted"],
                         padding=[14, 6], font=FONT_SMALL)
        style.map("TNotebook.Tab",
                  background=[("selected", C["surface"])],
                  foreground=[("selected", C["text"])])

        # タブ1: SQL エディタ
        self._build_tab_sql(nb)
        # タブ2: テーブルプレビュー
        self._build_tab_preview(nb)
        # タブ3: エクスポート
        self._build_tab_export(nb)
        # タブ4: インポートログ
        self._build_tab_log(nb)
        # タブ5: Google スプレッドシート
        self._build_tab_gsheets(nb)

        self.notebook = nb

    # ── タブ: SQL エディタ ────────────────────────────────────
    def _build_tab_sql(self, nb):
        tab = tk.Frame(nb, bg=C["bg"])
        nb.add(tab, text="⚡ SQL")

        # ツールバー
        bar = tk.Frame(tab, bg=C["surface2"], pady=6, padx=10)
        bar.pack(fill="x")
        tk.Button(bar, text="▶ 実行  (Ctrl+Enter)",
                  command=self._run_sql,
                  bg=C["accent"], fg="#fff",
                  font=FONT_BOLD, relief="flat", cursor="hand2",
                  padx=16, pady=5).pack(side="left")
        tk.Button(bar, text="🗑 クリア",
                  command=lambda: self.sql_editor.delete("1.0", "end"),
                  bg=C["surface"], fg=C["muted"],
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=10, pady=5).pack(side="left", padx=8)

        # クイッククエリ
        quick_frame = tk.Frame(bar, bg=C["surface2"])
        quick_frame.pack(side="right")
        tk.Label(quick_frame, text="クイック:", bg=C["surface2"],
                 fg=C["muted"], font=FONT_SMALL).pack(side="left")
        for label, sql in [
            ("全件",    "SELECT * FROM {table} LIMIT 100"),
            ("件数",    "SELECT COUNT(*) FROM {table}"),
            ("カラム", "PRAGMA table_info('{table}')"),
        ]:
            tk.Button(quick_frame, text=label,
                      command=lambda s=sql: self._quick_query(s),
                      bg=C["surface"], fg=C["muted"],
                      font=FONT_SMALL, relief="flat", cursor="hand2",
                      padx=8, pady=4).pack(side="left", padx=2)

        # エディタ
        editor_frame = tk.Frame(tab, bg=C["bg"])
        editor_frame.pack(fill="x", padx=10, pady=(8, 0))
        self.sql_editor = tk.Text(editor_frame, height=7,
                                  bg=C["surface"], fg=C["text"],
                                  insertbackground=C["text"],
                                  font=FONT_MONO, relief="flat",
                                  padx=12, pady=8,
                                  wrap="none")
        self.sql_editor.pack(fill="x")
        self.sql_editor.bind("<Control-Return>", lambda e: self._run_sql())
        self.sql_editor.insert("1.0",
            "-- SQL を入力して Ctrl+Enter または「実行」を押してください\n"
            "-- 例: SELECT * FROM テーブル名 LIMIT 100\n")

        # 結果エリア
        self.result_label = tk.Label(tab, text="",
                                     bg=C["bg"], fg=C["muted"],
                                     font=FONT_SMALL, anchor="w", padx=12)
        self.result_label.pack(fill="x", pady=(6, 2))

        result_frame = tk.Frame(tab, bg=C["bg"])
        result_frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self.result_tree, _ = self._make_treeview(result_frame)

        # エクスポートボタン（結果直接）
        exp_bar = tk.Frame(tab, bg=C["surface2"], pady=5, padx=10)
        exp_bar.pack(fill="x")
        tk.Label(exp_bar, text="結果を保存:", bg=C["surface2"],
                 fg=C["muted"], font=FONT_SMALL).pack(side="left")
        tk.Button(exp_bar, text="💾 CSV",
                  command=lambda: self._export_result("csv"),
                  bg=C["surface"], fg=C["green"],
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=10, pady=4).pack(side="left", padx=4)
        tk.Button(exp_bar, text="💾 Excel",
                  command=lambda: self._export_result("xlsx"),
                  bg=C["surface"], fg=C["green"],
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=10, pady=4).pack(side="left", padx=4)

        self._last_result = None

    # ── タブ: プレビュー ──────────────────────────────────────
    def _build_tab_preview(self, nb):
        tab = tk.Frame(nb, bg=C["bg"])
        nb.add(tab, text="👁 プレビュー")

        bar = tk.Frame(tab, bg=C["surface2"], pady=6, padx=10)
        bar.pack(fill="x")
        tk.Label(bar, text="テーブル:", bg=C["surface2"],
                 fg=C["muted"], font=FONT_SMALL).pack(side="left")
        self.preview_table_var = tk.StringVar()
        self.preview_combo = ttk.Combobox(bar,
                                          textvariable=self.preview_table_var,
                                          font=FONT_SMALL, width=25, state="readonly")
        self.preview_combo.pack(side="left", padx=6)
        tk.Button(bar, text="表示",
                  command=self._preview_table,
                  bg=C["accent"], fg="#fff",
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=12, pady=4).pack(side="left")
        self.preview_count = tk.Label(bar, text="",
                                      bg=C["surface2"], fg=C["muted"],
                                      font=FONT_SMALL)
        self.preview_count.pack(side="right", padx=10)

        preview_frame = tk.Frame(tab, bg=C["bg"])
        preview_frame.pack(fill="both", expand=True, padx=10, pady=8)
        self.preview_tree, _ = self._make_treeview(preview_frame)

    # ── タブ: エクスポート ────────────────────────────────────
    def _build_tab_export(self, nb):
        tab = tk.Frame(nb, bg=C["bg"])
        nb.add(tab, text="📤 エクスポート")

        frm = tk.Frame(tab, bg=C["bg"])
        frm.pack(fill="both", expand=True, padx=24, pady=20)

        def label(text):
            tk.Label(frm, text=text, bg=C["bg"], fg=C["muted"],
                     font=FONT_SMALL, anchor="w").pack(fill="x", pady=(12, 2))

        # SQL 入力
        label("① エクスポートする SQL（SELECT文）")
        self.export_sql = tk.Text(frm, height=5,
                                  bg=C["surface"], fg=C["text"],
                                  insertbackground=C["text"],
                                  font=FONT_MONO, relief="flat",
                                  padx=10, pady=8)
        self.export_sql.pack(fill="x")
        self.export_sql.insert("1.0", "SELECT * FROM テーブル名")

        # 出力先
        label("② 出力先ファイルを指定")
        path_row = tk.Frame(frm, bg=C["bg"])
        path_row.pack(fill="x")
        self.export_path_var = tk.StringVar(value="output.csv")
        tk.Entry(path_row, textvariable=self.export_path_var,
                 bg=C["surface"], fg=C["text"],
                 font=FONT_MONO, relief="flat",
                 insertbackground=C["text"]).pack(side="left", fill="x", expand=True)
        tk.Button(path_row, text="参照",
                  command=self._browse_export,
                  bg=C["surface2"], fg=C["text"],
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=10, pady=5).pack(side="left", padx=6)

        # 実行
        tk.Button(frm, text="💾 エクスポート実行",
                  command=self._do_export,
                  bg=C["green"], fg="#000",
                  font=FONT_BOLD, relief="flat", cursor="hand2",
                  padx=20, pady=10).pack(anchor="w", pady=16)

        self.export_result_label = tk.Label(frm, text="",
                                            bg=C["bg"], fg=C["green"],
                                            font=FONT_MAIN, anchor="w")
        self.export_result_label.pack(fill="x")

    # ── タブ: インポートログ ──────────────────────────────────
    def _build_tab_log(self, nb):
        tab = tk.Frame(nb, bg=C["bg"])
        nb.add(tab, text="📜 ログ")

        bar = tk.Frame(tab, bg=C["surface2"], pady=6, padx=10)
        bar.pack(fill="x")
        tk.Button(bar, text="🔄 更新",
                  command=self._refresh_log,
                  bg=C["surface"], fg=C["text"],
                  font=FONT_SMALL, relief="flat", cursor="hand2",
                  padx=10, pady=4).pack(side="left")

        log_frame = tk.Frame(tab, bg=C["bg"])
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)

        cols = ("id", "table_name", "rows", "source", "imported_at")
        self.log_tree, _ = self._make_treeview(log_frame, columns=cols,
            headings=["#", "テーブル名", "行数", "ソース", "インポート日時"],
            widths=[40, 180, 80, 320, 150])
        self._refresh_log()

    # ── ツリービュー ヘルパー ─────────────────────────────────
    def _make_treeview(self, parent, columns=None, headings=None, widths=None):
        style = ttk.Style()
        style.configure("Custom.Treeview",
                        background=C["surface"],
                        foreground=C["text"],
                        fieldbackground=C["surface"],
                        rowheight=24,
                        font=FONT_SMALL)
        style.configure("Custom.Treeview.Heading",
                        background=C["header_bg"],
                        foreground="#fff",
                        font=FONT_BOLD, relief="flat")
        style.map("Custom.Treeview",
                  background=[("selected", C["accent"])],
                  foreground=[("selected", "#fff")])

        frame = tk.Frame(parent, bg=C["bg"])
        frame.pack(fill="both", expand=True)

        if columns is None:
            columns = []

        tree = ttk.Treeview(frame, columns=columns or [],
                             show="headings",
                             style="Custom.Treeview")

        vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree.grid(row=0, column=0, sticky="nsew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        if headings:
            for col, hd, w in zip(columns, headings, widths or [100]*len(columns)):
                tree.heading(col, text=hd)
                tree.column(col, width=w, minwidth=50)

        return tree, frame

    def _populate_tree(self, tree, columns: list, rows: list):
        """ツリービューにデータをセット"""
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=max(100, len(col)*10), minwidth=60)
        tree.delete(*tree.get_children())
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            tree.insert("", "end", values=row, tags=(tag,))
        tree.tag_configure("even", background=C["surface"])
        tree.tag_configure("odd",  background=C["surface2"])

    # ── テーブル操作 ──────────────────────────────────────────
    def _refresh_tables(self):
        self.table_list.delete(0, "end")
        tables = self.engine.list_tables()
        for t in tables:
            count = self.engine.row_count(t)
            self.table_list.insert("end", f"  {t}  ({count:,}行)")
        # コンボも更新
        if hasattr(self, "preview_combo"):
            self.preview_combo["values"] = tables

    def _on_table_select(self, event=None):
        sel = self.table_list.curselection()
        if not sel:
            return
        raw = self.table_list.get(sel[0]).strip()
        table = raw.split("(")[0].strip()
        if hasattr(self, "preview_table_var"):
            self.preview_table_var.set(table)

    def _preview_table(self, event=None):
        table = self.preview_table_var.get() if hasattr(self, "preview_table_var") else None
        if not table:
            sel = self.table_list.curselection()
            if not sel:
                return
            raw = self.table_list.get(sel[0]).strip()
            table = raw.split("(")[0].strip()

        try:
            result = self.engine.execute_sql(f'SELECT * FROM "{table}" LIMIT 500')
            self._populate_tree(self.preview_tree,
                                result["columns"], result["rows"])
            self.preview_count.config(
                text=f"{self.engine.row_count(table):,} 行（最大500件表示）")
            self.notebook.select(1)  # プレビュータブへ
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def _drop_table(self):
        sel = self.table_list.curselection()
        if not sel:
            messagebox.showinfo("選択", "削除するテーブルを選択してください")
            return
        raw = self.table_list.get(sel[0]).strip()
        table = raw.split("(")[0].strip()
        if messagebox.askyesno("確認", f"テーブル「{table}」を削除しますか？"):
            self.engine.drop_table(table)
            self._refresh_tables()
            self.status_var.set(f"テーブル「{table}」を削除しました")

    # ── インポート ────────────────────────────────────────────
    def _import_csv(self):
        path = filedialog.askopenfilename(
            title="CSV ファイルを選択",
            filetypes=[("CSV", "*.csv"), ("すべて", "*.*")])
        if not path:
            return
        self._run_in_thread(
            lambda: self.engine.import_csv(path),
            on_done=lambda r: self._on_import_done(
                f"✓ 「{r['table']}」に {r['rows']:,} 行 インポート完了"))

    def _import_excel(self):
        path = filedialog.askopenfilename(
            title="Excel ファイルを選択",
            filetypes=[("Excel", "*.xlsx *.xls"), ("すべて", "*.*")])
        if not path:
            return
        self._run_in_thread(
            lambda: self.engine.import_excel(path),
            on_done=lambda r: self._on_import_done(
                f"✓ {len(r)} シートをインポート完了"))

    def _on_import_done(self, msg):
        self.status_var.set(msg)
        self._refresh_tables()
        self._refresh_log()
        messagebox.showinfo("インポート完了", msg)

    # ── SQL 実行 ──────────────────────────────────────────────
    def _run_sql(self):
        sql = self.sql_editor.get("1.0", "end").strip()
        sql = "\n".join(l for l in sql.splitlines()
                        if not l.strip().startswith("--"))
        if not sql:
            return
        try:
            result = self.engine.execute_sql(sql)
            if result["type"] == "select":
                self._populate_tree(self.result_tree,
                                    result["columns"], result["rows"])
                self.result_label.config(
                    text=f"  結果: {result['count']:,} 行",
                    fg=C["green"])
                self._last_result = result
            else:
                self.result_label.config(
                    text=f"  ✓ {result['affected']} 行が更新されました",
                    fg=C["green"])
                self._refresh_tables()
            self.status_var.set("SQL 実行完了")
        except RuntimeError as e:
            self.result_label.config(text=f"  ✗ {e}", fg=C["red"])
            self.status_var.set(f"SQL エラー: {e}")

    def _quick_query(self, template: str):
        table = self.preview_table_var.get() or ""
        if not table and self.table_list.curselection():
            raw = self.table_list.get(self.table_list.curselection()[0]).strip()
            table = raw.split("(")[0].strip()
        sql = template.format(table=table or "テーブル名")
        self.sql_editor.delete("1.0", "end")
        self.sql_editor.insert("1.0", sql)
        self.notebook.select(0)

    # ── エクスポート ──────────────────────────────────────────
    def _export_result(self, fmt: str):
        if not self._last_result:
            messagebox.showinfo("エラー", "先に SQL を実行してください")
            return
        ext = ".csv" if fmt == "csv" else ".xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("すべて", "*.*")])
        if not path:
            return
        try:
            if fmt == "csv":
                import csv as _csv
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = _csv.writer(f)
                    w.writerow(self._last_result["columns"])
                    w.writerows(self._last_result["rows"])
            else:
                import openpyxl
                from openpyxl.styles import Font, PatternFill
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(self._last_result["columns"])
                for row in self._last_result["rows"]:
                    ws.append(row)
                wb.save(path)
            messagebox.showinfo("完了", f"保存しました:\n{path}")
            self.status_var.set(f"エクスポート完了: {path}")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def _browse_export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
        if path:
            self.export_path_var.set(path)

    def _do_export(self):
        sql  = self.export_sql.get("1.0", "end").strip()
        path = self.export_path_var.get().strip()
        if not sql or not path:
            messagebox.showinfo("入力エラー", "SQL と出力先を入力してください")
            return
        try:
            ext = Path(path).suffix.lower()
            if ext == ".csv":
                result = self.engine.export_csv(sql, path)
                msg = f"✓ {result['rows']:,} 行 → {result['path']}"
            else:
                result = self.engine.export_excel({"データ": sql}, path)
                msg = f"✓ {result['total_rows']:,} 行 → {result['path']}"
            self.export_result_label.config(text=msg, fg=C["green"])
            self.status_var.set(msg)
        except Exception as e:
            self.export_result_label.config(text=f"✗ {e}", fg=C["red"])

    # ── ログ ──────────────────────────────────────────────────
    def _refresh_log(self):
        if not hasattr(self, "log_tree"):
            return
        logs = self.engine.import_log()
        self.log_tree.delete(*self.log_tree.get_children())
        for l in logs:
            self.log_tree.insert("", "end", values=(
                l["id"], l["table_name"], f"{l['rows']:,}",
                l["source"], l["imported_at"][:16]))

    # ── DB 切り替え ───────────────────────────────────────────
    def _open_db(self):
        path = filedialog.askopenfilename(
            title="DB ファイルを選択",
            filetypes=[("SQLite", "*.db *.sqlite *.sqlite3"), ("すべて", "*.*")])
        if path:
            self.engine.close()
            self.engine = DBEngine(Path(path))
            self.title(f"🗄 DB マネージャー — {path}")
            self._refresh_tables()
            self._refresh_log()
            self.status_var.set(f"DB を切り替えました: {path}")

    # ── スレッド実行 ──────────────────────────────────────────
    def _run_in_thread(self, fn, on_done=None):
        self.status_var.set("処理中...")

        def worker():
            try:
                result = fn()
                if on_done:
                    self.after(0, on_done, result)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("エラー", str(e)))
                self.after(0, lambda: self.status_var.set(f"エラー: {e}"))

        threading.Thread(target=worker, daemon=True).start()


    # ── Google Sheets タブ ────────────────────────────────────
    def _open_gsheets_tab(self):
        self.notebook.select(4)

    def _build_tab_gsheets(self, nb):
        tab = tk.Frame(nb, bg=C["bg"])
        nb.add(tab, text="📊 Google Sheets")
        self._gs_client = None

        # 認証セクション
        auth_box = tk.LabelFrame(tab, text=" 🔑 認証 ",
                                 bg=C["bg"], fg=C["gold"], font=FONT_BOLD,
                                 bd=1, relief="groove", labelanchor="nw")
        auth_box.pack(fill="x", padx=16, pady=(14, 6))
        inner = tk.Frame(auth_box, bg=C["bg"])
        inner.pack(fill="x", padx=12, pady=10)

        self._gs_auth_var = tk.StringVar(value="service")
        radio_frame = tk.Frame(inner, bg=C["bg"])
        radio_frame.pack(anchor="w")
        for val, label in [("service", "サービスアカウント（JSON キーファイル）"),
                            ("oauth",  "OAuth（Googleアカウントでブラウザログイン）")]:
            tk.Radiobutton(radio_frame, text=label, variable=self._gs_auth_var,
                           value=val, bg=C["bg"], fg=C["text"],
                           selectcolor=C["surface2"], activebackground=C["bg"],
                           font=FONT_SMALL).pack(side="left", padx=(0, 24))

        key_row = tk.Frame(inner, bg=C["bg"])
        key_row.pack(fill="x", pady=(8, 0))
        tk.Label(key_row, text="認証ファイル:", bg=C["bg"], fg=C["muted"],
                 font=FONT_SMALL).pack(side="left")
        self._gs_key_var = tk.StringVar(value="credentials/service_account.json")
        tk.Entry(key_row, textvariable=self._gs_key_var,
                 bg=C["surface"], fg=C["text"], font=FONT_MONO, relief="flat",
                 width=44, insertbackground=C["text"]).pack(side="left", padx=6)
        tk.Button(key_row, text="参照", command=self._gs_browse_key,
                  bg=C["surface2"], fg=C["text"], font=FONT_SMALL, relief="flat",
                  cursor="hand2", padx=8, pady=4).pack(side="left")

        auth_btn_row = tk.Frame(inner, bg=C["bg"])
        auth_btn_row.pack(anchor="w", pady=(10, 0))
        tk.Button(auth_btn_row, text="🔗 認証する", command=self._gs_authenticate,
                  bg=C["accent"], fg="#fff", font=FONT_BOLD, relief="flat",
                  cursor="hand2", padx=16, pady=6).pack(side="left")
        self._gs_auth_status = tk.Label(auth_btn_row, text="  未認証",
                                        bg=C["bg"], fg=C["muted"], font=FONT_SMALL)
        self._gs_auth_status.pack(side="left", padx=10)

        # スプレッドシート選択
        ss_box = tk.LabelFrame(tab, text=" 📄 スプレッドシート選択 ",
                               bg=C["bg"], fg=C["text"], font=FONT_BOLD,
                               bd=1, relief="groove", labelanchor="nw")
        ss_box.pack(fill="x", padx=16, pady=6)
        ss_inner = tk.Frame(ss_box, bg=C["bg"])
        ss_inner.pack(fill="x", padx=12, pady=10)

        url_row = tk.Frame(ss_inner, bg=C["bg"])
        url_row.pack(fill="x")
        tk.Label(url_row, text="URL または ID:", bg=C["bg"], fg=C["muted"],
                 font=FONT_SMALL).pack(side="left")
        self._gs_url_var = tk.StringVar()
        tk.Entry(url_row, textvariable=self._gs_url_var,
                 bg=C["surface"], fg=C["text"], font=FONT_MONO, relief="flat",
                 width=52, insertbackground=C["text"]).pack(side="left", padx=6)
        tk.Button(url_row, text="シート一覧取得", command=self._gs_load_sheets,
                  bg=C["surface2"], fg=C["text"], font=FONT_SMALL, relief="flat",
                  cursor="hand2", padx=10, pady=4).pack(side="left")

        sheet_row = tk.Frame(ss_inner, bg=C["bg"])
        sheet_row.pack(fill="x", pady=(8, 0))
        tk.Label(sheet_row, text="シート名:", bg=C["bg"], fg=C["muted"],
                 font=FONT_SMALL).pack(side="left")
        self._gs_sheet_var = tk.StringVar()
        self._gs_sheet_combo = ttk.Combobox(sheet_row, textvariable=self._gs_sheet_var,
                                             font=FONT_SMALL, width=28, state="readonly")
        self._gs_sheet_combo.pack(side="left", padx=6)
        tk.Label(sheet_row, text="テーブル名（DB）:", bg=C["bg"], fg=C["muted"],
                 font=FONT_SMALL).pack(side="left", padx=(20, 0))
        self._gs_table_var = tk.StringVar()
        tk.Entry(sheet_row, textvariable=self._gs_table_var,
                 bg=C["surface"], fg=C["text"], font=FONT_MONO, relief="flat",
                 width=20, insertbackground=C["text"]).pack(side="left", padx=6)

        imp_row = tk.Frame(ss_inner, bg=C["bg"])
        imp_row.pack(anchor="w", pady=(10, 0))
        tk.Button(imp_row, text="📥 DB に取り込む", command=self._gs_import,
                  bg=C["green"], fg="#000", font=FONT_BOLD, relief="flat",
                  cursor="hand2", padx=16, pady=7).pack(side="left")
        self._gs_import_status = tk.Label(imp_row, text="", bg=C["bg"],
                                          fg=C["green"], font=FONT_SMALL)
        self._gs_import_status.pack(side="left", padx=10)

        # プレビュー
        prev_box = tk.LabelFrame(tab, text=" 👁 プレビュー（先頭100行） ",
                                 bg=C["bg"], fg=C["text"], font=FONT_BOLD,
                                 bd=1, relief="groove", labelanchor="nw")
        prev_box.pack(fill="both", expand=True, padx=16, pady=(6, 14))
        prev_inner = tk.Frame(prev_box, bg=C["bg"])
        prev_inner.pack(fill="both", expand=True, padx=8, pady=8)
        self._gs_preview_tree, _ = self._make_treeview(prev_inner)

    def _gs_browse_key(self):
        path = filedialog.askopenfilename(
            title="認証ファイルを選択",
            filetypes=[("JSON", "*.json"), ("すべて", "*.*")])
        if path:
            self._gs_key_var.set(path)

    def _gs_authenticate(self):
        from google_sheets import GoogleSheetsClient, check_packages, install_hint
        if not check_packages():
            messagebox.showerror("パッケージ不足", install_hint())
            return
        def do_auth():
            if self._gs_auth_var.get() == "service":
                return GoogleSheetsClient.from_service_account(self._gs_key_var.get() or None)
            else:
                return GoogleSheetsClient.from_oauth(self._gs_key_var.get() or None)
        def on_done(client):
            self._gs_client = client
            self._gs_auth_status.config(text="  ✓ 認証成功", fg=C["green"])
            self.status_var.set("Google Sheets 認証完了")
        self._gs_auth_status.config(text="  認証中...", fg=C["gold"])
        self._run_in_thread(do_auth, on_done=on_done)

    def _gs_load_sheets(self):
        if not self._gs_client:
            messagebox.showinfo("認証", "先に「認証する」を押してください"); return
        url = self._gs_url_var.get().strip()
        if not url:
            messagebox.showinfo("入力", "URL または ID を入力してください"); return
        from google_sheets import GoogleSheetsClient
        def do_load():
            sid = GoogleSheetsClient.extract_id(url)
            return self._gs_client.list_sheets(sid)
        def on_done(sheets):
            names = [s["title"] for s in sheets]
            self._gs_sheet_combo["values"] = names
            if names:
                self._gs_sheet_var.set(names[0])
                import re
                self._gs_table_var.set(re.sub(r"[^\w]", "_", names[0]).lower())
            self.status_var.set(f"{len(names)} シート取得")
        self._run_in_thread(do_load, on_done=on_done)

    def _gs_import(self):
        if not self._gs_client:
            messagebox.showinfo("認証", "先に「認証する」を押してください"); return
        url        = self._gs_url_var.get().strip()
        sheet_name = self._gs_sheet_var.get().strip()
        table_name = self._gs_table_var.get().strip()
        if not url:
            messagebox.showinfo("入力", "URL または ID を入力してください"); return
        from google_sheets import GoogleSheetsClient
        import re
        def do_import():
            sid  = GoogleSheetsClient.extract_id(url)
            data = self._gs_client.get_sheet_data(sid, sheet_name or None)
            if not data["columns"]:
                raise ValueError("データが空です")
            tbl = re.sub(r"[^\w]", "_", table_name or sheet_name or "gsheet").lower()
            self.engine.conn.execute(f'DROP TABLE IF EXISTS "{tbl}"')
            cols_def = ", ".join(f'"{h}" TEXT' for h in data["columns"])
            self.engine.conn.execute(f'CREATE TABLE "{tbl}" ({cols_def})')
            ph       = ", ".join("?" * len(data["columns"]))
            cols_str = ", ".join(f'"{h}"' for h in data["columns"])
            self.engine.conn.executemany(
                f'INSERT INTO "{tbl}" ({cols_str}) VALUES ({ph})', data["rows"])
            from datetime import datetime
            self.engine.conn.execute(
                "INSERT INTO _import_log (table_name,source,rows,imported_at) VALUES(?,?,?,?)",
                (tbl, f"GoogleSheets:{url}[{sheet_name}]",
                 len(data["rows"]), datetime.now().isoformat()))
            self.engine.conn.commit()
            return {"table": tbl, "rows": len(data["rows"]),
                    "columns": data["columns"], "preview": data}
        def on_done(r):
            msg = f"✓ 「{r['table']}」に {r['rows']:,} 行 取り込み完了"
            self._gs_import_status.config(text=msg)
            self.status_var.set(msg)
            self._refresh_tables()
            self._refresh_log()
            self._populate_tree(self._gs_preview_tree,
                                r["preview"]["columns"], r["preview"]["rows"][:100])
        self._gs_import_status.config(text="取り込み中...", fg=C["gold"])
        self._run_in_thread(do_import, on_done=on_done)

    def on_close(self):
        self.engine.close()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()
