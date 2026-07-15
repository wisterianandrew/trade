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
    sma: list[float] = field(default_factory=list)
    endprice: list[float] = field(default_factory=list)
    high: list[float] = field(default_factory=list)
    recent_high: float = 0 #直近高値

    account_class: str = "" #口座区分（現物用）

    get_date : int = 0 #取得日（現物用）・建日（信用用）
    build_price : float #建単価（信用用）
    build_market : str #建市場（信用用）

    order_suggest: int | str = ""
    
    sell_reason: str = "" #確認画面に出す用
    rule1: bool = False #序列1
    rule2: bool = False #序列2
    rule3: bool = False #序列3

    number_of_rule3 : float = 5 #percent　
    reference_percent : int = 0 #やれっていわれたやつ

    enabled: bool = False #ルール適用する銘柄かどうか　

spot_stocks : dict[str, Stock] = {}
credit_stocks : dict[str, Stock] = {}

_SAVE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "save.json")

def save_settings():
    data = {"spot": {}, "credit": {}}
    for code, stock in spot_stocks.items():
        data["spot"][code] = {
            "enabled": stock.enabled,
            "get_date": stock.get_date,
            "number_of_rule3": stock.number_of_rule3,
        }
    for code, stock in credit_stocks.items():
        data["credit"][code] = {
            "enabled": stock.enabled,
            "get_date": stock.get_date,
            "number_of_rule3": stock.number_of_rule3,
        }
    with open(_SAVE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def load_settings():
    if not os.path.exists(_SAVE_PATH):
        return
    with open(_SAVE_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    for code, saved in data.get("spot", {}).items():
        if code in spot_stocks:
            spot_stocks[code].enabled = saved.get("enabled", False)
            spot_stocks[code].get_date = saved.get("get_date", 0)
            spot_stocks[code].number_of_rule3 = saved.get("number_of_rule3", 5)
    for code, saved in data.get("credit", {}).items():
        if code in credit_stocks:
            credit_stocks[code].enabled = saved.get("enabled", False)
            credit_stocks[code].get_date = saved.get("get_date", 0)
            credit_stocks[code].number_of_rule3 = saved.get("number_of_rule3", 5)

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
        cell_quantity = ws.Range(f"D{row}").Value

        account_class = ws.Range(f"C{row}").Value


        # Stock型のstockを作る
        stock = Stock(
            code=str(code),
            name=name,
            account_class=account_class,
            quantity=int(quantity),
            price=float(price),
            cell_quantity=int(cell_quantity) // 100 * 100,
        )
        # dictに追加(dictionary?)
        spot_stocks[stock.code] = stock

        row += 1
    # -------------------------
    # デバッグ表示
    # -------------------------
    if test:
        print("===== 現物保有銘柄 =====")
        for stock in spot_stocks.values():
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
        order_quantity = ws.Range(f"H{row}").Value

        build_market = ws.Range(f"D{row}").Value
        build_price = ws.Range(f"J{row}").Value
        get_date = ws.Range(f"K{row}").Value

        #credit_stocksの中に同じcodeに対応するstockが既にあるかも（後で確認）
        stock = Stock(
            code=str(code),
            name=name,
            quantity=int(quantity),
            price=float(price),
            cell_quantity=int(order_quantity) // 100 * 100,
        )

        credit_stocks[stock.code] = stock

        row += 1

    # -------------------------
    # デバッグ表示
    # -------------------------
    if test:
        print("===== 信用建玉 =====")
        for stock in credit_stocks.values():
            print(
                f"code={stock.code}, "
                f"name={stock.name}, "
                f"quantity={stock.quantity}, "
                f"price={stock.price}"
            )
        print("=======================")

#必要になった関数１（一つの銘柄の直近高値,3日分の高値、終値取得）
def get_high_end_recent(ws, stock):
    # 20日前のdate（YYYYMMDD）
    date = (datetime.today()
        - timedelta(days=20)
    ).strftime("%Y%m%d")
    target_cell = "A1"
    
    never_change_number = 100 # never_change_number > row
    # RSS関数を書き込み
    ws.Range(target_cell).Value = (
        f'=RssChartPast($A$2:$J$2,'
        f'{stock.code},"D",'
        f'{date},{never_change_number})'
    )
    #エクセルに再計算をお願いする
    ws.Application.Calculate()
    # Excel/RSS更新待ち（0.5s）
    time.sleep(0.5)

    # 直近高値探索(最近の方から)
    # 実行時が購入日の時、直近高値は現在値
    if(stock.get_date == datetime.today()):
        stock.recent_high = stock.price
    else:
        row = 30 # never_change_number > row
        max_high = 0
        while (row >= 3):
            value1 = str(ws.Range(f"F{row}").Value)
            value2 = str(ws.Range(f"I{row}").Value)
            # 区切り行までスキップ（一見問題ありな実装だがrssの仕様上OK！）
            if value1 == "None" or "-" in value1:
                row -= 1
                continue
            
            # 購入日以前にはもどらないようにする
            if(stock.get_date > int((ws.Range(f"D{row}").Value).replace("/", ""))):
                if test: print("symbol")
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
    
    # 高値、終値探索
    endprice_all = []
    high_all = []
    row = 30
    # 一旦全部取得して
    while (row>=3):
        value1 = str(ws.Range(f"G{row}").Value)
        value2 = str(ws.Range(f"I{row}").Value)
        # 区切り行までスキップ（一見問題ありな実装だがrssの仕様上OK！）
        if value1 == "None" or "-" in value1:
            row -= 1
            continue

        endprice_all.append(float(value2))        
        high_all.append(float(value1))

        row -= 1
    # 最新3件を逆順で取得
    stock.endprice = endprice_all[:3]
    stock.high = high_all[:3]

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
    # Excel/RSS更新待ち（0.5s）
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

#stockの中身を埋める
def get_info(ws1, ws2, progress_cb=None):
    all_stocks = list(spot_stocks.values()) + list(credit_stocks.values())
    total = len(all_stocks)
    for i, stock in enumerate(all_stocks):
        get_high_end_recent(ws1, stock)
        get_sma(ws2, stock)
        if progress_cb:
            progress_cb(i + 1, total, stock.name)

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


#売却値段の算出、rule1 > rule2 > rule3にしてほしいみたい
def calc_sell_price_basedon_rules():
    # 現物
    for stock in spot_stocks.values():
        zone = _get_time_zone(datetime.now())
        rule1(stock)
        if not stock.sell_reason == "三日ルール":
            rule2(stock)
            if not stock.sell_reason == "二日ルール":
                rule3(stock)
        
        # -------------------------
        # デバッグ表示
        # -------------------------
        if test:
            print(
                f"{stock.code} "
                f"{stock.order_suggest},{stock.sell_reason}"
            )
    # 信用
    for stock in credit_stocks.values():
        rule1(stock)
        if not stock.sell_reason == "三日ルール":
            rule2(stock)
            if not stock.sell_reason == "二日ルール":
                rule3(stock)

        # -------------------------
        # デバッグ表示
        # -------------------------
        if test:
            print(
                f"{stock.code} "
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

# 執行中銘柄達の取得
def get_executing_stocks(ws) -> list[int]:
    list = []
    # 3行目から読む
    row = 3
    while True:
        state = ws.Range(f"C{row}").Value
        # 空セルまたは区切り行で終了
        if state is None or (isinstance(state,str) and "-" in state):
            break
        if state == "執行中" and (spot_stocks[ws.Range(f"F{row}").Value].enabled == True or credit_stocks[ws.Range(f"F{row}").Value].enabled == True):
            list.append(int(ws.Range(f"A{row}").Value))
        row += 1
    return list

#キャンセル注文
def cancel(ws,id,list):
    row=1 #「下に続けて入力する形式」にするかも
    for order_id in list:
        time.sleep(0.5)
        ws.Range(f"A{row}").Value = (
            f'=RssCancelOrder('
            f'{id},{order_flag},{order_id})'
        )
        id += 1
        
    #　注文関数はエクセルに残らないようにする（次回起動時のバグのもと）
    ws.Range(f"A{row}").Value = None #「下に続けて入力する形式」にするなら、sheet全体を初期化

def cancel_order(ws_id, ws_check, ws_order):
    # 取り消し注文入力に使うidを取得
    id = get_proper_id(ws_id)
    # 執行中　かつ　取引に出そうとしている銘柄 の注文番号を取得
    executing_list = get_executing_stocks(ws_check)
    # id executing_listから取り消し注文
    cancel(ws_order,id,executing_list)

#ruleごと、zoneごと？に売り
def spot_cell(ws, id, stock):
    time.sleep(0.5)
    cla=-1
    if stock.account_class == "特定":
        cla = 0
    elif stock.account_class == "一般":
        cla = 1
    elif stock.account_class == "NISA":
        cla = 2
    elif stock.account_class == "旧NISA":
        cla = 3
    
    if stock.cell_reason == "三日ルール" or stock.cell_reason == "二日ルール":
        #zoneで変わるかと思ったが、期間指定のところを今週中にしたので、同じ注文でよくなった。
        ws.Range(f"A1").Value = (
            f'=RssStockOrder('
            f'{id},{order_flag},{stock.code},1,0,1,{stock.cell_quantity}'
            f',0,,2,,{cla})'
        )
    elif stock.cell_reason == "高値ルール":
        #修正
        date = (datetime.today()
            - timedelta(days=20)
            ).strftime("%Y%m%d")
        ws.Range(f"A1").Value = (
            f'=RssStockOrder('
            f'{id},{order_flag},{stock.code},1,0,1,{stock.cell_quantity}'
            f',0,,2,,{cla})'
        )
    id += 1

#ruleごと、zoneごと？に売り
def credit_cell(ws, id, stock):
    time.sleep(0.5)
    #stock.cell_quantityはそのまま使えない、建の数超えるかも（stockのつくりが悪い）
    # -> get_dateが古いものから順に、合計cell_quantityになるまで売り注文（厄介だね。再帰？）
    qty = stock.cell_quantity #データ構造とともに後で直す
    if stock.build_maket == "東証":
        mar = 1
    elif stock.build_maket == "JNX":
        mar = 2
    if stock.build_maket == "JAX":
        mar = 5
    if stock.build_maket == "Chi-X":
        mar = 6
    pri = stock.build_price
    get = stock.get_date
    if stock.cell_reason == "三日ルール" or stock.cell_reason == "二日ルール":
        ws.Range(f"A1").Value = (
            f'=RssMarginCloseOrder('
            f'{id},{order_flag},{stock.code},1,0,1,2,{qty}'
            f',0,,2,,0,{get},{pri},{mar})'
        )
    elif stock.cell_reason == "高値ルール":
        #修正
        ws.Range(f"A1").Value = (
            f'=RssMarginCloseOrder('
            f'{id},{order_flag},{stock.code},1,0,1,2,{qty}'
            f',0,,2,,0,{get},{pri},{mar})'
        )

    id += 1

def cell_order(ws_id, ws_order):
    # 新規注文入力に使うidを取得
    id = get_proper_id(ws_id)
    # 現物の注文
    for stock in spot_stocks.values():
        if stock.enabled == True:
            spot_cell(ws_order,id,stock)
    # 信用の注文
    for stock in credit_stocks.values():
        if stock.enabled == True:
            credit_cell(ws_order,id,stock)
    # 注文関数はエクセルに残さない
    ws_order.Range(f"A1").Value = None
    

#def check_orders():


def order(ws_id, ws_check, ws_order):
    zone = _get_time_zone(datetime.now())
    #執行中の注文を全取り消し
    cancel_order()
    #売り注文を出す
    cell_order()
    #注文後の確認(完了画面用)
    #check_orders()



if __name__ == "__main__":
    print("semi_auto trading system started!")
    sheet=open_xlsx()
    get_my_spot_stock(sheet["spot"])
    get_my_credit_stock(sheet["credit"])
    load_settings()
    get_info(sheet["rss_chart"],sheet["rss_sma"])
    calc_sell_price_basedon_rules()
    order(sheet["id"], sheet["check_order"], sheet["control"])
