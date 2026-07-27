from dataclasses import dataclass, field
import win32com.client
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import time
import json
import os

test = True
order_flag = False

@dataclass
class Stock:
    code: str #銘柄コード
    name: str #銘柄名
    quantity: int #保有数量
    price: float #現在値
    cell_quantity : int #売却数量
    sma: list[float] = field(default_factory=list) #移動平均
    endprice: list[float] = field(default_factory=list) #終値
    high: list[float] = field(default_factory=list) #高値
    recent_high: float = 0 #直近高値
    account_class: str = "" #口座区分（現物）
    get_date : int = 0 #取得日（現物）・建日（信用）
    build_price : float = 0 #建単価（信用）
    build_market : str = "" #建市場（信用）
    order_suggest: int | str = "" #売却提案
    sell_reason: str = "" #売却理由
    rule1: bool = False #序列1
    rule2: bool = False #序列2
    rule3: bool = False #序列3
    number_of_rule3 : float = 5 #固有のpercent　
    reference_percent : float = 5 #後で使う
    enabled: bool = False #ルール適用する銘柄かどうか　

#credit_stocksの同一code内のStockをまとめた要約用（描画・発注数量指定にのみ使用、個別ロットの詳細は持たない）
@dataclass
class Sum_stock(Stock):
    pass

spot_stocks : dict[str, list[Stock]] = {}
credit_stocks : dict[str, list[Stock]] = {}
credit_stocks_code : dict[str, Sum_stock] = {}

_SAVE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "save.json")

#現物はcode+account_classの二重キー、信用はcredit_stocks_code（code単位、get_dateは管理しない）で保存
def save_settings():
    data = {"spot": {}, "credit": {}}
    for code, stocks in spot_stocks.items():
        for stock in stocks:
            data["spot"].setdefault(code, {})[stock.account_class] = {
                "enabled": stock.enabled,
                "get_date": stock.get_date,
                "number_of_rule3": stock.number_of_rule3,
            }
    for code, stock in credit_stocks_code.items():
        data["credit"][code] = {
            "enabled": stock.enabled,
            "number_of_rule3": stock.number_of_rule3,
        }
    # 書き込み中の異常終了でファイルが空になるのを防ぐため、一時ファイル経由で置き換える
    tmp_path = _SAVE_PATH + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    os.replace(tmp_path, _SAVE_PATH)

# save.jsonを読み込む（存在しない・空・壊れている場合はNoneを返す）
def _read_settings():
    if not os.path.exists(_SAVE_PATH):
        return None
    with open(_SAVE_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    if not content.strip():
        return None
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        print(f"{_SAVE_PATH} が壊れているため設定の復元をスキップしました")
        return None

# 現物の設定を復元。get_dateはget_infoの直近高値計算で使うため、get_infoより前に呼ぶこと
def load_spot_settings():
    data = _read_settings()
    if data is None:
        return
    for code, accounts in data.get("spot", {}).items():
        for stock in spot_stocks.get(code, []):
            saved = accounts.get(stock.account_class)
            if saved is None:
                continue
            stock.enabled = saved.get("enabled", False)
            stock.get_date = saved.get("get_date", 0)
            stock.number_of_rule3 = saved.get("number_of_rule3", 5)

# 信用の設定を復元。credit_stocks_codeを参照するため、build_credit_stocks_codeの後に呼ぶこと
def load_credit_settings():
    data = _read_settings()
    if data is None:
        return
    for code, saved in data.get("credit", {}).items():
        enabled = saved.get("enabled", False)
        number_of_rule3 = saved.get("number_of_rule3", 5)
        if code in credit_stocks_code:
            credit_stocks_code[code].enabled = enabled
            credit_stocks_code[code].number_of_rule3 = number_of_rule3
        # ルール適用は個別ロット（credit_stocks）に対して行うので合わせて反映
        for stock in credit_stocks.get(code, []):
            stock.enabled = enabled
            stock.number_of_rule3 = number_of_rule3

# ファイルを開く
def open_xlsx():
    _PATH = r"C:\Users\wiste\Documents\program\trading\rakuten.xlsx"
    excel = win32com.client.Dispatch("Excel.Application")
    wb = None
    for w in excel.Workbooks:
        if w.FullName.lower() == _PATH.lower():
            wb = w
            break
    if wb is None:
        wb = excel.Workbooks.Open(_PATH)
    return {
        "excel": excel,
        "wb": wb,
        "spot": wb.Worksheets("MarketSheet1"),
        "credit": wb.Worksheets("MarketSheet2"),
        "control": wb.Worksheets("ControlSheet"),
        "id":wb.Worksheets("RssOrderIDList"),
        "check_order":wb.Worksheets("RssOrderList"),
        "rss_market": wb.Worksheets("RssMarket"),
        "rss_chart": wb.Worksheets("RssChartPast"),
        "rss_sma": wb.Worksheets("RssTrendSMA"),
    }

#現物保有の銘柄取得
def get_my_spot_stock(ws):
    spot_stocks.clear()
    # 3行目から読む（1,2行目はヘッダー）
    row = 3
    while True:
        code = ws.Range(f"A{row}").Value
        # 空セルまたは区切り行で終了
        if code is None or (isinstance(code,str) and "-" in code):
            break
        # 銘柄名にfloat以外を含むならスキップ(Aとかには対応できない)
        if not isinstance(code,float):
            print(f"非対応銘柄: {code}")

            row += 1
            continue
        
        code = int(code)
        name = ws.Range(f"B{row}").Value
        quantity = ws.Range(f"D{row}").Value
        price = ws.Range(f"G{row}").Value

        account_class = ws.Range(f"C{row}").Value

        # Stock型のstockを作る
        stock = Stock(
            code=str(code),
            name=name,
            account_class=account_class,
            quantity=int(quantity),
            price=float(price),
            cell_quantity=int(quantity) // 100 * 100,
        )
        # dictに追加(同じcodeでも口座区分違いなどで複数持てる)（特定とNISAを想定）
        spot_stocks.setdefault(stock.code, []).append(stock)

        row += 1
    # -------------------------
    # デバッグ表示
    # -------------------------
    if test:
        print("===== 現物保有銘柄 =====")
        for stocks in spot_stocks.values():
            for stock in stocks:
                print(
                    f"code={stock.code}, "
                    f"name={stock.name}, "
                    f"quantity={stock.quantity}, "
                    f"price={stock.price}"
                )
        print("=======================")

#信用建玉の銘柄取得（買建しか想定してないけど）（rssが分かれてるから素直に分ける）
def get_my_credit_stock(ws):
    credit_stocks.clear()
    # 3行目から読む
    row = 3
    while True:
        code = ws.Range(f"A{row}").Value
        # 空セルまたは区切り行で終了
        if code is None or (isinstance(code,str) and "-" in code):
            break
        # 銘柄名にfloat以外を含むならスキップ
        if not isinstance(code,float):
            print(f"非対応銘柄: {code}")

            row += 1
            continue

        code = int(code)
        name = ws.Range(f"B{row}").Value
        quantity = ws.Range(f"H{row}").Value
        price = ws.Range(f"M{row}").Value

        build_market = ws.Range(f"D{row}").Value
        build_price = ws.Range(f"J{row}").Value
        get_date = ws.Range(f"K{row}").Value

        # credit_stocksの中に同じcodeに対応するstockが既にあるかも（建日違いなど）→listで保持
        stock = Stock(
            code=str(code),
            name=name,
            quantity=int(quantity),
            price=float(price),
            cell_quantity=int(quantity) // 100 * 100,
            build_market=str(build_market),
            build_price=float(build_price),
            get_date = int(get_date)
        )
        # dictに追加(同じcodeでも建日・建市場・建値違いなどで複数持てる)(code、建数、建日、建市場、建値が一緒の時どうなるか検証)
        credit_stocks.setdefault(stock.code, []).append(stock)

        row += 1

    # -------------------------
    # デバッグ表示
    # -------------------------
    if test:
        print("===== 信用建玉 =====")
        for stocks in credit_stocks.values():
            for stock in stocks:
                print(
                    f"code={stock.code}, "
                    f"name={stock.name}, "
                    f"quantity={stock.quantity}, "
                    f"price={stock.price}"
                )
        print("=======================")

#必要になった関数１（一つの銘柄の直近高値,3日分の高値、終値取得）
def get_high_end_recent(ws, stock, my_date):
    # 20日前のdate（YYYYMMDD）
    date = (datetime.today()
        - timedelta(days=20)
    ).strftime("%Y%m%d")
    target_cell = "A1"
    
    # RSS関数を書き込み
    ws.Range(target_cell).Value = (
        f'=RssChartPast($A$2:$J$2,'
        f'{stock.code},"D",'
        f'{date},100)'
    )
    #エクセルに再計算をお願いする
    ws.Application.Calculate()
    # Excel/RSS更新待ち
    time.sleep(0.5)

    # 直近高値探索(最近の方から)
    # 実行時が購入日の時、直近高値は現在値（購入日=当日は遡る過去データが無いため）
    # my_dateはint(YYYYMMDD)なので、todayも同じ形式に揃えて比較する
    today_int = int(datetime.today().strftime("%Y%m%d"))
    if my_date == today_int:
        stock.recent_high = stock.price
    else:
        row = 30 
        max_high = 0
        while (row >= 3):
            value1 = str(ws.Range(f"F{row}").Value)
            value2 = str(ws.Range(f"I{row}").Value)
            # 区切り行までスキップ（一見問題ありな実装だがrssの仕様上OK！）
            if value1 == "None" or "-" in value1:
                row -= 1
                continue
            
            # 購入日以前にはもどらないようにする
            if(my_date > int((ws.Range(f"D{row}").Value).replace("/", ""))):
                break

            # strなのでcastしてあげよう
            high = float(max(float(value1),float(value2)))

            # 最大値更新
            if high >= max_high: 
                max_high = high
            else:
                break

            row -= 1
        # Stockへ保存
        stock.recent_high = max_high
    
    # 高値、終値探索、提案パーセント計算
    endprice_all = []
    high_all = []
    small_all = []
    start_all = []
    percent_all = []
    row = 30
    # 一旦全部取得して
    while (row>=3):
        value1 = str(ws.Range(f"G{row}").Value)
        value2 = str(ws.Range(f"I{row}").Value)
        value3 = str(ws.Range(f"F{row}").Value)
        value4 = str(ws.Range(f"H{row}").Value)
        # 区切り行までスキップ（一見問題ありな実装だがrssの仕様上OK！）
        if value1 == "None" or "-" in value1:
            row -= 1
            continue

        endprice_all.append(float(value2))        
        high_all.append(float(value1))
        small_all.append(float(value4))
        start_all.append(float(value3))
        percent_all.append((float(value1) - float(value4)) / float(value3))

        row -= 1
    # 最新3件を逆順で取得、推奨のパーセントを計算
    stock.endprice = endprice_all[:3]
    stock.high = high_all[:3]
    stock.reference_percent = (sum(percent_all) / len(percent_all)) * 100 

    # -------------------------
    # デバッグ表示
    # -------------------------
    if test:
        print(
            f"{stock.code} "
            f"{stock.name} "
            f"recent_high={stock.recent_high} "
            f"endprice={stock.endprice} "
            f"high={stock.high}"
        )

# 必要になった関数2(sma取得)
def get_sma(ws, stock):
    target_cell = "A1"
    ws.Range(target_cell).Value = (
        f'=RssTrendSMA($A$2:$L$2,'
        f'{stock.code}'
        f',"D",4,5)'
    )
    #エクセルに再計算をお願いする
    ws.Application.Calculate()
    # Excel/RSS更新待ち
    time.sleep(0.5)

    # sma取得
    row = 10
    sma_all = []
    while (row >= 3):
        value = str(ws.Range(f"F{row}").Value)
        # 区切り行までスキップ（一見問題ありな実装だがrssの仕様上OK！）
        if value == "None" or "-" in value:
            row -= 1
            continue
        sma_all.append(float(value))
        
        row -= 1
    
    stock.sma = sma_all[:4]

    # -------------------------
    # デバッグ表示
    # -------------------------
    if test:
        if len(stock.sma) < 4:
            print(f"{stock.code}: SMAデータ不足 ({len(stock.sma)}件)")
        else:
            print(f"{stock.code} sma:{stock.sma[0]},{stock.sma[1]},{stock.sma[2]},{stock.sma[3]}")

#stockの中身を埋める（銘柄codeごとに1回だけ取得し、同じcodeの残りのStockにはコピーする）
def get_info(ws1, ws2, progress_cb=None):
    code_groups = [stocks for stocks in spot_stocks.values() if stocks] + \
                  [stocks for stocks in credit_stocks.values() if stocks]
    total = len(code_groups)
    for i, stocks in enumerate(code_groups):
        min_date = min(s.get_date for s in stocks)
        first = stocks[0]
        get_high_end_recent(ws1, first, min_date)
        get_sma(ws2, first)
        for other in stocks[1:]:
            other.sma = list(first.sma)
            other.recent_high = first.recent_high
            other.endprice = list(first.endprice)
            other.high = list(first.high)
            other.reference_percent = first.reference_percent
        if progress_cb:
            progress_cb(i + 1, total, first.name)

#credit_stocksをcodeごとにまとめてcredit_stocks_codeを作る（get_info実行後に呼ぶ）
def build_credit_stocks_code():
    credit_stocks_code.clear()
    for code, stocks in credit_stocks.items():
        if not stocks:
            continue
        first = stocks[0]
        total_quantity = sum(s.quantity for s in stocks)
        credit_stocks_code[code] = Sum_stock(
            code=first.code,
            name=first.name,
            quantity=total_quantity,
            price=first.price,
            cell_quantity=total_quantity // 100 * 100,
            sma=list(first.sma),
            endprice=list(first.endprice),
            high=list(first.high),
            recent_high=first.recent_high,
            get_date=min(s.get_date for s in stocks),
            build_price=sum(s.build_price * s.quantity for s in stocks) / total_quantity,
            reference_percent=first.reference_percent,
        )

# 呼値の算出
def get_tick_size(price: float) -> int:
    if price <= 3_000:
        return 1
    elif price <= 5_000:
        return 5
    elif price <= 30_000:
        return 10
    elif price <= 50_000:
        return 50
    elif price <= 300_000:
        return 100
    elif price <= 3_000_000:
        return 1_000
    elif price <= 5_000_000:
        return 5_000
    else:
        return 100_000

# rule3で使用、呼値に対応したorder_suggestの算出
def floor_to_tick(price: float) -> int:
    tick = get_tick_size(price)
    return int(price // tick) * tick

# mainでの呼値チェック用関数
def is_valid_tick_price(price: float) -> bool:
    tick = get_tick_size(price)
    return price % tick == 0

# 実行時の時間帯を把握
def _get_time_zone(now: datetime, is_holiday: bool = False) -> str | None:
    if now.weekday() >= 5 or is_holiday:
        return "day_off"       # 土日または祝日
    if now.hour >= 17:
        return "after_close"   # 17:00-23:59
    if now.hour <= 8:
        return "before_open"   # 00:00-08:59
    if now.hour == 14 or (now.hour == 15 and now.minute <= 30):
        return "closing"       # 14:00-15:30
    return None


# rule1, order_suggestとsell_reasonを埋める
def rule1(stock):
    if len(stock.endprice) < 3 or len(stock.sma) < 3:
        if test: print(f"{stock.code}: データ不足でrule1スキップ")
        return
    zone = _get_time_zone(datetime.now())
    if zone is None:
        if test: print("ルール１の適用時間外")
        return
    if zone in ("after_close", "before_open", "day_off"):
        cond = (stock.endprice[0] < stock.sma[0] and
                stock.endprice[1] < stock.sma[1] and
                stock.endprice[2] < stock.sma[2])
    else:  # closing
        cond = (stock.price < stock.sma[0] and
                stock.endprice[0] < stock.sma[1] and
                stock.endprice[1] < stock.sma[2])
    if cond:
        stock.order_suggest = "成行"
        stock.sell_reason = "三日ルール"


# rule2
def rule2(stock):
    if len(stock.high) < 2 or len(stock.sma) < 2:
        if test: print(f"{stock.code}: データ不足でrule2スキップ")
        return
    zone = _get_time_zone(datetime.now())
    if zone is None:
        if test: print("ルール２の適用時間外")
        return
    if zone in ("after_close", "before_open", "day_off"):
        cond = (stock.high[0] < stock.sma[0] and
                stock.high[1] < stock.sma[1])
    else:  # closing
        cond = (stock.price < stock.sma[0] and
                stock.high[0] < stock.sma[1])
    if cond:
        stock.order_suggest = "成行"
        stock.sell_reason = "二日ルール"


# rule3
def rule3(stock):
    zone = _get_time_zone(datetime.now())
    if zone is None:
        if test: print("ルール３の適用時間外")
        return
    else: 
        raw = stock.recent_high * (100 - stock.number_of_rule3) / 100
        stock.order_suggest = floor_to_tick(raw)
        stock.sell_reason = "高値ルール"


#1銘柄にrule1 > rule2 > rule3の優先順でルールを適用
def _apply_rules(stock):
    rule1(stock)
    if not stock.sell_reason == "三日ルール":
        rule2(stock)
        if not stock.sell_reason == "二日ルール":
            rule3(stock)

#売却値段の算出、rule1 > rule2 > rule3にしてほしいみたい
def calc_sell_price_basedon_rules():
    # 現物・信用の個別ロット（発注はこの個別ロットのsell_reason/order_suggestを使う）
    for stocks in list(spot_stocks.values()) + list(credit_stocks.values()):
        for stock in stocks:
            _apply_rules(stock)

            # -------------------------
            # デバッグ表示
            # -------------------------
            if test:
                print(
                    f"{stock.code} "
                    f"{stock.order_suggest},{stock.sell_reason}"
                )

    # 信用の集約（GUI表示・注文確認ダイアログはこの集約側を参照するため同じルールを適用）
    # 同一codeのロットは入力値が同一なので、個別ロットと同じ結果になる
    for stock in credit_stocks_code.values():
        _apply_rules(stock)
        if test:
            print(
                f"[集約] {stock.code} "
                f"{stock.order_suggest},{stock.sell_reason}"
            )

# 注文処理系
def get_proper_id(ws) -> int:
    id = 1
    # 3行目から読む
    row = 3
    while True:
        number = ws.Range(f"A{row}").Value
        # 空セルまたは区切り行で終了
        if number is None or (isinstance(number,str) and "-" in number):
            break
        id += 1
        row += 1
    return id

# 銘柄コードを辞書キー（str(int)）に正規化する。
# Excelはコードをfloat(7203.0)で返すことがあり、そのままではspot_stocks等のstrキーと一致しないため。
def _normalize_code(value) -> str:
    try:
        return str(int(float(value)))
    except (TypeError, ValueError):
        return str(value).strip()

# 執行中銘柄達の取得
def get_executing_stocks(ws) -> list[int]:
    list = []
    # 3行目から読む
    row = 3
    while True:
        state1 = ws.Range(f"C{row}").Value
        # 空セルまたは区切り行で終了
        if state1 is None or (isinstance(state1,str) and "-" in state1):
            break
        code = _normalize_code(ws.Range(f"F{row}").Value)
        enabled = (any(s.enabled for s in spot_stocks.get(code, [])) or
                   any(s.enabled for s in credit_stocks.get(code, [])))
        if state1 == "執行中" and enabled:
            list.append(int(ws.Range(f"A{row}").Value))
        row += 1
    return list

#キャンセル注文（次のidを返す）
def cancel(ws,id,list):
    for order_id in list:
        time.sleep(0.5)
        ws.Range(f"A1").Value = (
            f'=RssCancelOrder('
            f'{id},{order_flag},{order_id})'
        )
        id += 1

    #　注文関数はエクセルに残らないようにする（次回起動時のバグのもと）
    ws.Range(f"A1").Value = None
    return id

def cancel_order(ws_id, ws_check, ws_order):
    # 取り消し注文入力に使うidを取得
    id = get_proper_id(ws_id)
    # 執行中　かつ　取引に出そうとしている銘柄 の注文番号を取得
    executing_list = get_executing_stocks(ws_check)
    # id executing_listから取り消し注文（消費後の次のidを返す）
    return cancel(ws_order,id,executing_list)

# 口座区分・建市場のコード対応表（発注APIの引数値）。未対応値は-1でAPIが弾く。
_ACCOUNT_CLASS_MAP = {"特定": 0, "一般": 1, "NISA": 2, "旧NISA": 3}
_BUILD_MARKET_MAP  = {"東証": 1, "JNX": 2, "JAX": 5, "Chi-X": 6}

# 発注対象(enabled)のうち、口座区分・建市場が未対応の銘柄の説明リストを返す（空なら問題なし）
def validate_enabled_orders() -> list[str]:
    problems = []
    for stocks in spot_stocks.values():
        for s in stocks:
            if s.enabled and s.account_class not in _ACCOUNT_CLASS_MAP:
                problems.append(f"現物 {s.code} {s.name}：口座区分「{s.account_class}」が未対応")
    for code, stocks in credit_stocks.items():
        summary = credit_stocks_code.get(code)
        if summary is None or not summary.enabled:
            continue
        for s in stocks:
            if s.build_market not in _BUILD_MARKET_MAP:
                problems.append(f"信用 {s.code} {s.name}：建市場「{s.build_market}」が未対応")
    return problems

# 発注直前の最終チェック：発注対象の売却数量が0でなく100の倍数か（GUI編集をすり抜けた不正値の保険）
def validate_order_quantities() -> list[str]:
    problems = []
    # 現物：enabledロットのcell_quantity
    for stocks in spot_stocks.values():
        for s in stocks:
            if not s.enabled:
                continue
            if s.cell_quantity <= 0 or s.cell_quantity % 100 != 0:
                problems.append(f"現物 {s.code} {s.name}：売却数量 {s.cell_quantity} が不正（0でなく100の倍数が必要）")
    # 信用：発注数量指定に使う集約(summary)のcell_quantity
    for summary in credit_stocks_code.values():
        if not summary.enabled:
            continue
        if summary.cell_quantity <= 0 or summary.cell_quantity % 100 != 0:
            problems.append(f"信用 {summary.code} {summary.name}：売却数量 {summary.cell_quantity} が不正（0でなく100の倍数が必要）")
    return problems

#現物、ruleごとに売り（発注に使ったidを消費して次のidを返す）
def spot_cell(ws, id, stock):
    time.sleep(0.5)
    cla = _ACCOUNT_CLASS_MAP.get(stock.account_class, -1)

    if stock.sell_reason == "三日ルール" or stock.sell_reason == "二日ルール":
        #zoneで変わるかと思ったが、期間指定のところを今週中にしたので、同じ注文でよくなった。
        ws.Range(f"A1").Value = (
            f'=RssStockOrder('
            f'{id},{order_flag},{stock.code},1,0,1,{stock.cell_quantity}'
            f',0,,2,,{cla})'
        )
    elif stock.sell_reason == "高値ルール":
        date = (datetime.today()
            + timedelta(days=20)
            ).strftime("%Y%m%d")
        ws.Range(f"A1").Value = (
            f'=RssStockOrder('
            f'{id},{order_flag},{stock.code},1,2,1,{stock.cell_quantity}'
            f',,,5,{date},{cla},{stock.order_suggest},2,0)'
        )
    id += 1
    return id

#信用、ruleごとに売り（qtyは呼び出し側でcredit_stocks_codeの残数量に合わせてクランプ済み）
def credit_cell(ws, id, stock, qty):
    time.sleep(0.5)
    mar = _BUILD_MARKET_MAP.get(stock.build_market, -1)
    pri = stock.build_price
    get = stock.get_date
    if stock.sell_reason == "三日ルール" or stock.sell_reason == "二日ルール":
        ws.Range(f"A1").Value = (
            f'=RssMarginCloseOrder('
            f'{id},{order_flag},{stock.code},1,0,1,2,{qty}'
            f',0,,2,,0,{get},{pri},{mar})'
        )
    elif stock.sell_reason == "高値ルール":
        date = (datetime.today()
            + timedelta(days=20)
            ).strftime("%Y%m%d")
        ws.Range(f"A1").Value = (
            f'=RssMarginCloseOrder('
            f'{id},{order_flag},{stock.code},1,2,1,2,{qty}'
            f',,,5,{date},0,{get},{pri},{mar},{stock.order_suggest},2,0)'
        )

    id += 1
    return id

def cell_order(ws_id, ws_order, id=None):
    # 新規注文入力に使うidを取得（取消注文から連番を引き継いだ場合はそれを使う）
    if id is None:
        id = get_proper_id(ws_id)
    # 現物の注文
    for stocks in spot_stocks.values():
        for stock in stocks:
            if stock.enabled == True:
                id = spot_cell(ws_order,id,stock)
    # 信用の注文：credit_stocks_code[code].cell_quantityに達するまで、建日の古いロットから順に売る
    for code, stocks in credit_stocks.items():
        summary = credit_stocks_code.get(code)
        if summary is None or not summary.enabled:
            continue
        remaining = summary.cell_quantity
        for stock in sorted(stocks, key=lambda s: s.get_date):
            if remaining <= 0:
                break
            qty = min(stock.cell_quantity, remaining)
            id = credit_cell(ws_order, id, stock, qty)
            remaining -= qty
    # 注文関数はエクセルに残さない
    ws_order.Range(f"A1").Value = None
    return id
    

def check_order(ws) -> bool:
    row = 3
    while(True):
        # 空セルまたは区切り行で終了
        v = ws.Range(f"F{row}").value
        if v is None or (isinstance(v,str) and "-" in v):
            break
        if "エラー" in v:
            return False
        row += 1
    return True



def order(ws_id, ws_check, ws_order) -> bool:
    #執行中の注文を全取り消し（消費後の次のidを受け取る）
    next_id = cancel_order(ws_id, ws_check, ws_order)
    #売り注文を出す（取消で消費したidの続きから採番し、重複を防ぐ）
    cell_order(ws_id, ws_order, next_id)
    #注文後の確認(完了画面用)
    return check_order(ws_check)



if __name__ == "__main__":
    print("semi_auto trading system started!")
    # 想定している流れ
    if test:
        sheet=open_xlsx()
        get_my_spot_stock(sheet["spot"])
        # get_infoの直近高値計算で現物のget_dateを使うため、get_infoより前に現物設定を復元する
        load_spot_settings()
        get_my_credit_stock(sheet["credit"])
        get_info(sheet["rss_chart"],sheet["rss_sma"])
        build_credit_stocks_code()
        # credit_stocks_codeの作成後でないとcredit側の設定を反映できないため、get_infoの後で読み込む
        load_credit_settings()
        calc_sell_price_basedon_rules()
        order(sheet["id"], sheet["check_order"], sheet["control"])
