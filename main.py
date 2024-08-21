import yfinance as yf
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 한글깨짐방지
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

def get_stock_data(tickers, start, end, price_type='Close'):
    adjusted_start = (pd.to_datetime(start) - pd.DateOffset(months=12)).strftime('%Y-%m-%d')
    data = {}
    for ticker in tickers:
        stock = yf.Ticker(ticker)
        df = stock.history(start=adjusted_start, end=end)
        df.index = df.index.tz_localize(None)  # tz-aware 제거
        data[ticker] = df[price_type]
    return pd.DataFrame(data)

def get_closest_date(data, ticker, date):
    date = pd.to_datetime(date).tz_localize(None)  # tz-aware 제거
    dates = data[ticker].dropna().index
    if date not in dates:
        closest_date = dates.asof(date)
        return closest_date
    return date

def calculate_momentum_scores(data, date):
    date = pd.to_datetime(date).tz_localize(None)  # tz-aware 제거
    momentum_scores = {}
    for ticker in data.columns:
        closest_date_now = get_closest_date(data, ticker, date)
        closest_date_1m = get_closest_date(data, ticker, date - pd.DateOffset(months=1))
        closest_date_3m = get_closest_date(data, ticker, date - pd.DateOffset(months=3))
        closest_date_6m = get_closest_date(data, ticker, date - pd.DateOffset(months=6))
        closest_date_12m = get_closest_date(data, ticker, date - pd.DateOffset(months=12))

        price_now = data.loc[closest_date_now, ticker]
        price_1m = data.loc[closest_date_1m, ticker]
        price_3m = data.loc[closest_date_3m, ticker]
        price_6m = data.loc[closest_date_6m, ticker]
        price_12m = data.loc[closest_date_12m, ticker]

        # NaN 값이 있는 경우 처리
        if pd.isna(price_now) or pd.isna(price_1m) or pd.isna(price_3m) or pd.isna(price_6m) or pd.isna(price_12m):
            score = -np.inf  # 데이터가 부족한 경우 매우 낮은 점수 부여
        else:
            score = ((price_now / price_1m - 1) * 12 +
                     (price_now / price_3m - 1) * 4 +
                     (price_now / price_6m - 1) * 2 +
                     (price_now / price_12m - 1) * 1)

        momentum_scores[ticker] = score
    return momentum_scores

def simulate_trading(start, end, initial_cash, buy_commission=0.001, sell_commission=0.001, price_type='Close'):
    attack_assets = ["SPY", "EFA", "EEM", "AGG", "QQQ"]
    safe_assets = ["LQD", "IEF", "SHY"]

    tickers = attack_assets + safe_assets
    data = get_stock_data(tickers, start, end, price_type)

    cash = initial_cash
    portfolio = {ticker: 0 for ticker in tickers}
    value_history = []
    trade_history = []
    start_date = pd.to_datetime(start).tz_localize(None)  # tz-aware 제거
    dates = pd.date_range(start=start_date, end=end, freq=pd.DateOffset(months=1))

    total_steps = len(dates)
    for i, date in enumerate(dates):
        momentum_scores = calculate_momentum_scores(data, date)

        best_attack_asset = max(attack_assets, key=lambda x: momentum_scores.get(x, -np.inf))
        best_safe_asset = max(safe_assets, key=lambda x: momentum_scores.get(x, -np.inf))

        if all(momentum_scores[asset] <= 0 for asset in attack_assets):
            best_asset = best_safe_asset
            strategy = '안전자산 투자'
        else:
            best_asset = best_attack_asset
            strategy = '공격형자산 투자'

        # 매도
        for ticker in portfolio:
            if portfolio[ticker] > 0:
                closest_date = get_closest_date(data, ticker, date)
                unit_price = data.loc[closest_date, ticker]
                total_price = portfolio[ticker] * unit_price * (1 - sell_commission)
                cash += total_price
                portfolio[ticker] = 0

        # 매수
        if best_asset:
            closest_date = get_closest_date(data, best_asset, date)
            unit_price = data.loc[closest_date, best_asset]
            shares_to_buy = cash // (unit_price * (1 + buy_commission))
            total_price = shares_to_buy * unit_price * (1 + buy_commission)
            portfolio[best_asset] += shares_to_buy
            cash -= total_price

        total_value = cash + sum(portfolio[ticker] * data.loc[get_closest_date(data, ticker, date), ticker] for ticker in portfolio)

        value_history.append(total_value)
        trade_history.append({
            'date': date,
            'strategy': strategy,
            'best_asset': best_asset,
            'cash': cash,
            'total_value': total_value,
            'current_return': ((total_value / initial_cash) - 1) * 100,
            'portfolio': portfolio.copy()
        })

        # 진행률 표시
        print(f"진행률: {i + 1}/{total_steps} ({(i + 1) / total_steps * 100:.2f}%)")

    final_value = value_history[-1]
    annual_return = ((final_value / initial_cash) ** (1 / (len(dates) / 12)) - 1) * 100

    plt.figure(figsize=(14, 7))
    plt.plot(dates[:len(value_history)], value_history, label='포트폴리오 가치')
    plt.xlabel('날짜')
    plt.ylabel('가치 (원)')
    plt.title('포트폴리오 가치 변동')
    plt.legend()
    plt.show()

    print(f"시작일: {start}, 종료일: {end}")

    # 매달 데이터 표 형식으로 출력 및 엑셀로 저장
    trade_df = pd.DataFrame(trade_history)
    trade_df['date'] = trade_df['date'].dt.strftime('%Y-%m-%d')
    print(trade_df)

    # 엑셀 파일로 저장
    output_filename = 'simulation_results.xlsx'
    with pd.ExcelWriter(output_filename) as writer:
        trade_df.to_excel(writer, index=False, sheet_name='Simulation Results')

    return final_value, annual_return

# 설정값 고정
start = "2023-03-01"  # 예시 시작일
end = "2024-08-07"
initial_cash = 1000000  # 초기 자산
buy_commission = 0.001  # 매수 커미션
sell_commission = 0.001  # 매도 커미션
price_type = 'Close'  # 'Close' 또는 'Open'

# 시뮬레이션 실행
final_value, annual_return = simulate_trading(start, end, initial_cash, buy_commission, sell_commission, price_type)
print(f"최종 포트폴리오 가치: {final_value:.2f}원")
print(f"연간 평균 수익률: {annual_return:.2f}%")
