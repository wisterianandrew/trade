import tkinter as tk
from tkinter import ttk, messagebox
from collections import defaultdict
import sub1

_GET_DATE_HINT = "20日以内購入なら入力"


class TradingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("半自動売却サポートツール")
        self.data = {}
        self.sheets = None

        self.create_widgets()

    def create_widgets(self):
        frame_top = tk.Frame(self.root)
        frame_top.pack(pady=5)

        tk.Button(frame_top, text="判定実行", command=self.run_analysis).pack(side=tk.LEFT, padx=5)

        self.is_holiday = False
        self._holiday_btn = tk.Button(
            frame_top, text="祝日 OFF", width=8,
            bg="#dddddd", fg="#666666",
            command=self._toggle_holiday,
        )
        self._holiday_btn.pack(side=tk.LEFT, padx=5)

        self.status_var = tk.StringVar()
        tk.Label(self.root, textvariable=self.status_var, fg="#666666", font=("", 9),
                 anchor=tk.W).pack(fill=tk.X, padx=8)

        columns = (
            "適用*", "銘柄コード", "口座区分", "銘柄名", "取得日*",
            "現在値", "最新の終値", "最新の5日sma", "直近高値", "参考 %",
            "高値用 %*", "売却提案*", "売却理由", "売却数量*"
        )

        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        col_widths = {
            "適用*": 40, "銘柄コード": 70, "口座区分": 70, "銘柄名": 120, "取得日*": 130,
            "現在値": 70, "最新の終値": 80, "最新の5日sma": 80, "直近高値": 90, "参考 %":80,
            "高値用 %*": 80, "売却提案*": 100, "売却理由": 70, "売却数量*": 90,
        }
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 80), anchor=tk.CENTER)

        self.tree.tag_configure("header", background="#4a7ebf", foreground="white", font=("", 9, "bold"))

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double_click)

        frame_bottom = tk.Frame(self.root)
        frame_bottom.pack(pady=5)

        tk.Button(frame_bottom, text="注文実行", command=self.execute_orders).pack()

    # -------------------------
    # ヘルパー
    # -------------------------
    def _toggle_holiday(self):
        self.is_holiday = not self.is_holiday
        if self.is_holiday:
            self._holiday_btn.config(text="祝日 ON", bg="#e67e22", fg="white")
        else:
            self._holiday_btn.config(text="祝日 OFF", bg="#dddddd", fg="#666666")

    def _set_status(self, msg: str):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _progress_cb(self, done: int, total: int, name: str):
        self._set_status(f"データ取得中... {name}  ({done}/{total})")

    @staticmethod
    def _idx(val, i):
        try:
            return val[i]
        except (TypeError, IndexError):
            return "-"

    @staticmethod
    def _fmt_date(get_date, stock_type):
        if get_date == 0  and stock_type == "spot":
            return f"｢{_GET_DATE_HINT}｣"
        return f"｢{get_date}｣"

    @staticmethod
    def _find_stock(row_data):
        # 現物：code+account_classで個別のStockを特定。信用：code単位のcredit_stocks_code（集約）を返す
        if row_data["stock_type"] == "spot":
            for s in sub1.spot_stocks.get(row_data["code"], []):
                if s.account_class == row_data["account_class"]:
                    return s
            return None
        return sub1.credit_stocks_code.get(row_data["code"])

    @staticmethod
    def _propagate_to_credit_lots(code, **fields):
        # enabled・number_of_rule3はルール計算・注文実行が個別ロット（credit_stocks）を見るため合わせて反映
        for s in sub1.credit_stocks.get(code, []):
            for key, value in fields.items():
                setattr(s, key, value)

    @staticmethod
    def _market_data_looks_connected():
        # 直近高値が全銘柄0、または全銘柄同じ値のときはRSS取得に失敗している可能性が高い
        values = [s.recent_high for stocks in sub1.spot_stocks.values() for s in stocks]
        values += [s.recent_high for s in sub1.credit_stocks_code.values()]
        if not values:
            return True
        if all(v == 0 for v in values):
            return False
        if len(values) >= 2 and len(set(values)) == 1:
            return False
        return True

    # -------------------------
    # データ読み込み・表示
    # -------------------------
    def load_data(self):
        self.data = {}
        self.tree.delete(*self.tree.get_children())

        self.tree.insert("", tk.END,
            values=("", "【現物】", "", "", "", "", "", "", "", "", "", "", "", ""),
            tags=("header",))

        for stocks in sub1.spot_stocks.values():
            for stock in stocks:
                iid = f"spot_{stock.code}_{stock.account_class}"
                self.data[iid] = {
                    "on": stock.enabled,
                    "code": stock.code,
                    "account_class": stock.account_class,
                    "name": stock.name,
                    "quantity": stock.quantity,
                    "cell_quantity": stock.cell_quantity,
                    "order_suggest": stock.order_suggest,
                    "price": stock.price,
                    "recent_high": stock.recent_high,
                    "stock_type": "spot",
                    "reason": stock.sell_reason,
                    "reference_percent": stock.reference_percent,
                    "number_of_rule3": stock.number_of_rule3,
                    "get_date": stock.get_date,
                }
                self.tree.insert("", tk.END, iid=iid, values=(
                    "✓" if stock.enabled else "□",
                    stock.code,
                    stock.account_class,
                    stock.name,
                    self._fmt_date(stock.get_date, "spot"),   # #5 編集可
                    stock.price,
                    self._idx(stock.endprice, 0),
                    self._idx(stock.sma, 0),
                    stock.recent_high,
                    f"{stock.reference_percent:.2f}",
                    f"｢{stock.number_of_rule3}｣",              # #11 編集可
                    f"｢{stock.order_suggest}｣",                # #12 編集可
                    stock.sell_reason,
                    f"｢{stock.cell_quantity}｣",                # #14 編集可
                ))

        self.tree.insert("", tk.END,
            values=("", "【信用】", "", "", "", "", "", "", "", "", "", "", "", ""),
            tags=("header",))

        for stock in sub1.credit_stocks_code.values():
            iid = f"credit_{stock.code}"
            self.data[iid] = {
                "on": stock.enabled,
                "code": stock.code,
                "account_class": "",
                "name": stock.name,
                "quantity": stock.quantity,
                "cell_quantity": stock.cell_quantity,
                "order_suggest": stock.order_suggest,
                "price": stock.price,
                "recent_high": stock.recent_high,
                "stock_type": "credit",
                "reason": stock.sell_reason,
                "reference_percent": stock.reference_percent,
                "number_of_rule3": stock.number_of_rule3,
                "get_date": stock.get_date,
            }
            self.tree.insert("", tk.END, iid=iid, values=(
                "✓" if stock.enabled else "□",
                stock.code,
                "",
                stock.name,
                self._fmt_date(stock.get_date, "credit"),  # #5 編集可
                stock.price,
                self._idx(stock.endprice, 0),
                self._idx(stock.sma, 0),
                stock.recent_high,
                f"{stock.reference_percent:.2f}",
                f"｢{stock.number_of_rule3}｣",              # #11 編集可
                f"｢{stock.order_suggest}｣",                # #12 編集可
                stock.sell_reason,
                f"｢{stock.cell_quantity}｣",                # #14 編集可
            ))

    def run_analysis(self):
        self._set_status("Excelを開いています...")
        sheets = sub1.open_xlsx()
        self.sheets = sheets
        self._set_status("現物データ取得中...")
        sub1.get_my_spot_stock(sheets["spot"])
        # get_infoの直近高値計算で現物のget_dateを使うため、get_infoより前に現物設定を復元する
        sub1.load_spot_settings()
        self._set_status("信用データ取得中...")
        sub1.get_my_credit_stock(sheets["credit"])
        sub1.get_info(sheets["rss_chart"], sheets["rss_sma"], self._progress_cb)
        sub1.build_credit_stocks_code()
        # credit_stocks_codeの作成後でないとcredit側の設定を反映できないため、get_infoの後で読み込む
        sub1.load_credit_settings()
        self._set_status("ルール計算中...")
        sub1.calc_sell_price_basedon_rules(self.is_holiday)
        sub1.save_settings()
        self._set_status("")
        if not self._market_data_looks_connected():
            messagebox.showwarning("警告", "エクセルとマーケットスピードの接続を確認してください")
        self.load_data()

    # -------------------------
    # ON/OFFクリック切替
    # -------------------------
    def on_click(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)

        if not item:
            return

        if col == "#1":
            row_data = self.data.get(item)
            if row_data is None:
                return

            new_value = not row_data["on"]
            if new_value and row_data["cell_quantity"] == 0:
                messagebox.showwarning("警告", f"【{row_data['name']}】\n保有数量が100株未満のため有効化できません")
                return
            row_data["on"] = new_value

            stock = self._find_stock(row_data)
            if stock is None:
                return
            stock.enabled = new_value
            if row_data["stock_type"] == "credit":
                self._propagate_to_credit_lots(row_data["code"], enabled=new_value)
            sub1.save_settings()

            symbol = "✓" if new_value else "□"
            values = list(self.tree.item(item, "values"))
            values[0] = symbol
            self.tree.item(item, values=values)

    # -------------------------
    # 編集可能セルのダブルクリック
    # #5(取得日)  #11(高値ルール%)  #12(売却提案)  #14(売却数量)
    # 表示値は [value] 形式
    # -------------------------
    def on_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        item = sel[0]
        col = self.tree.identify_column(event.x)

        if col not in ("#5", "#11", "#12", "#14"):
            return

        row_data = self.data.get(item)
        if row_data is None:
            return

        col_index = int(col.replace("#", "")) - 1
        raw = self.tree.item(item, "values")[col_index].strip("｢｣")
        if raw == _GET_DATE_HINT:
            raw = ""

        # 売却提案が「成行」のときは変更不可
        if col == "#12" and raw == "成行":
            messagebox.showinfo("編集不可", "三日ルール・二日ルール適用時の成行注文は変更できません")
            return

        edit_entry = tk.Entry(self.root)
        edit_entry.insert(0, raw)
        edit_entry.focus()

        def save_edit(_):
            new_value = edit_entry.get().strip()
            if not new_value:
                edit_entry.destroy()
                return
            stock = self._find_stock(row_data)
            if stock is None:
                edit_entry.destroy()
                return

            if col == "#5":  # 取得日 (YYYYMMDD)
                try:
                    val = int(new_value)
                except ValueError:
                    messagebox.showerror("入力エラー", "YYYYMMDD形式の整数を入力してください\n例：20260616")
                    return
                if len(str(val)) != 8:
                    messagebox.showerror("入力エラー", "YYYYMMDD形式（8桁）で入力してください\n例：20260616")
                    return
                row_data["get_date"] = val
                stock.get_date = val
                sub1.save_settings()
                display = f"｢{val}｣"

            elif col == "#11":  # 高値ルール%
                try:
                    val = float(new_value)
                except ValueError:
                    messagebox.showerror("入力エラー", "数値（%）を入力してください\n例：5 または 2.5")
                    return
                if val <= 0 or val >= 100:
                    messagebox.showerror("入力エラー", "0より大きく100未満の値を入力してください")
                    return
                row_data["number_of_rule3"] = val
                stock.number_of_rule3 = val
                if row_data["stock_type"] == "credit":
                    self._propagate_to_credit_lots(row_data["code"], number_of_rule3=val)
                sub1.save_settings()
                display = f"｢{val}｣"

            elif col == "#12":  # 売却提案（intのみ。成行はここに到達しない）
                try:
                    price = int(new_value)
                except ValueError:
                    messagebox.showerror("入力エラー", "呼値に対応した整数を入力してください")
                    return
                if not sub1.is_valid_tick_price(price):
                    tick = sub1.get_tick_size(price)
                    messagebox.showerror("入力エラー", f"呼値エラー：{price}円は無効です。\nこの価格帯の呼値は{tick}円単位です。")
                    return
                row_data["order_suggest"] = price
                stock.order_suggest = price
                # 信用は_find_stockが集約(summary)を返すため、実発注が見る個別ロットにも伝播させる
                if row_data["stock_type"] == "credit":
                    self._propagate_to_credit_lots(row_data["code"], order_suggest=price)
                display = f"｢{price}｣"

            else:  # #14 売却数量
                try:
                    val = int(new_value)
                except ValueError:
                    messagebox.showerror("入力エラー", "整数を入力してください")
                    return
                if val <= 0 or val % 100 != 0:
                    messagebox.showerror("入力エラー", "100の倍数を入力してください")
                    return
                if val > row_data["quantity"]:
                    messagebox.showerror("入力エラー", f"保有数量（{row_data['quantity']}）を超えられません")
                    return
                row_data["cell_quantity"] = val
                stock.cell_quantity = val
                display = f"｢{val}｣"

            values = list(self.tree.item(item, "values"))
            values[col_index] = display
            self.tree.item(item, values=values)
            edit_entry.destroy()

        def cancel_edit(_):
            edit_entry.destroy()

        edit_entry.bind("<Return>", save_edit)
        edit_entry.bind("<Escape>", cancel_edit)

        edit_entry.place(
            x=event.x_root - self.root.winfo_rootx(),
            y=event.y_root - self.root.winfo_rooty()
        )

    # -------------------------
    # 注文確認ダイアログ構築
    # -------------------------
    @staticmethod
    def _insert_confirm_text(txt, selected):
        txt.tag_configure("section",   font=("", 15, "bold"), spacing1=6, spacing3=2)
        txt.tag_configure("divider",   font=("", 10), foreground="#888888", spacing3=8)
        txt.tag_configure("rule",      font=("", 13, "bold"), spacing1=10, spacing3=4)
        txt.tag_configure("stockname", font=("", 13, "bold"), spacing1=2)
        txt.tag_configure("normal",    font=("", 12), spacing1=2)
        txt.tag_configure("emphasis",  font=("", 13, "bold"), foreground="#2471a3",
                          background="#d6eaf8")

        for section, label in [("spot", "【現物】"), ("credit", "【信用】")]:
            rows = [r for r in selected if r["stock_type"] == section]
            if not rows:
                continue

            txt.insert(tk.END, f"{label}\n", "section")
            txt.insert(tk.END, "─" * 44 + "\n", "divider")

            by_reason = defaultdict(list)
            for r in rows:
                by_reason[r["reason"]].append(r)

            for reason in ["三日ルール", "二日ルール", "高値ルール"]:
                if reason not in by_reason:
                    continue
                txt.insert(tk.END, f"▼  {reason}\n", "rule")
                for r in by_reason[reason]:
                    suggest = f"{r['order_suggest']}円" if r["order_suggest"] != "成行" else "成行"
                    txt.insert(tk.END, f"    {r['code']}  {r['name']}\n", "stockname")
                    if reason == "三日ルール":
                        txt.insert(tk.END, "    終値が3日間SMAを下回ったため\n", "normal")
                    elif reason == "二日ルール":
                        txt.insert(tk.END, "    高値が2日間SMAを下回ったため\n", "normal")
                    else:
                        pct = 100 - r["number_of_rule3"]
                        txt.insert(tk.END, f"    直近高値 {r['recent_high']}円 の {pct}%\n", "normal")
                    txt.insert(tk.END, "    売却提案：", "normal")
                    txt.insert(tk.END, f"  {suggest}  ", "emphasis")
                    txt.insert(tk.END, "    売却数量：", "normal")
                    txt.insert(tk.END, f"  {r['cell_quantity']}株  ", "emphasis")
                    txt.insert(tk.END, "\n\n", "normal")

            txt.insert(tk.END, "\n", "normal")

    # -------------------------
    # 注文（仮）
    # -------------------------
    def execute_orders(self):
        zone = sub1._get_time_zone(sub1.datetime.now(), self.is_holiday)
        """
        if zone == "day_off":
            messagebox.showwarning("休場日", "本日は土日・祝日のため注文できません")
            return
        """
        if zone is None:
            messagebox.showwarning("システム対応時間外", "注文可能時間は以下の通りです\n・14:00〜15:30\n・17:00〜翌08:59")
            return

        selected = [row for row in self.data.values() if row["on"]]

        if not selected:
            messagebox.showwarning("警告", "注文対象がありません")
            return

        # 未対応の口座区分・建市場は発注前に検出して一括中止する（部分発注を防ぐ）
        problems = sub1.validate_enabled_orders()
        if problems:
            messagebox.showerror(
                "注文前チェック",
                "以下の銘柄は口座区分／建市場が未対応のため注文できません。\n"
                "該当銘柄をOFFにするか、データを確認してください。\n\n"
                + "\n".join(problems)
            )
            return

        # 発注数量が0でなく100の倍数か最終チェック（不正値があれば一括中止）
        qty_problems = sub1.validate_order_quantities()
        if qty_problems:
            messagebox.showerror(
                "注文前チェック",
                "以下の銘柄は発注数量が不正です（0でなく100の倍数が必要）。\n"
                "該当銘柄の売却数量を修正してください。\n\n"
                + "\n".join(qty_problems)
            )
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("注文確認")
        dlg.geometry("780x560")
        dlg.grab_set()
        dlg.resizable(True, True)

        txt_frame = tk.Frame(dlg)
        txt_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        vsb = ttk.Scrollbar(txt_frame, orient=tk.VERTICAL)
        txt = tk.Text(txt_frame, wrap=tk.WORD, font=("", 12), yscrollcommand=vsb.set)
        vsb.config(command=txt.yview)
        self._insert_confirm_text(txt, selected)
        txt.config(state=tk.DISABLED)

        txt.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        txt_frame.rowconfigure(0, weight=1)
        txt_frame.columnconfigure(0, weight=1)

        confirmed = tk.BooleanVar(value=False)

        btn_frame = tk.Frame(dlg)
        btn_frame.pack(pady=(0, 12))

        def on_ok():
            confirmed.set(True)
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        tk.Button(btn_frame, text="実行", width=10, command=on_ok).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="キャンセル", width=10, command=on_cancel).pack(side=tk.LEFT, padx=10)

        dlg.wait_window()

        if confirmed.get():
            if self.sheets is None:
                messagebox.showwarning("警告", "「判定実行」が行われていないため、注文を実行できません")
                return

            for row in selected:
                print(f"注文: {row['code']} {row['cell_quantity']}株 {row['order_suggest']} 理由:{row['reason']}")

            success = sub1.order(self.sheets["id"], self.sheets["check_order"], self.sheets["control"])
            if success:
                messagebox.showinfo("完了", "注文処理が完了しました")
            else:
                messagebox.showerror("エラー", "注文処理でエラーが発生しました。RssOrderListシートをご確認ください")


if __name__ == "__main__":
    root = tk.Tk()
    app = TradingApp(root)
    #起動時のウィンドウサイズ
    root.geometry("1300x650")
    root.mainloop()
