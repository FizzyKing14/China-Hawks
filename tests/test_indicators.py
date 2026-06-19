import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import indicators, news, risk, signals  # noqa: E402


def test_sma_basic():
    assert indicators.sma([1, 2, 3, 4], 2) == 3.5
    assert indicators.sma([1, 2], 5) is None


def test_ema_matches_sma_seed():
    prices = [10, 11, 12, 13, 14]
    series = indicators.ema_series(prices, 3)
    assert series[1] is None
    assert series[2] == sum(prices[:3]) / 3


def test_rsi_all_gains_is_100():
    prices = list(range(1, 20))
    assert indicators.rsi(prices, 14) == 100.0


def test_rsi_in_range():
    prices = [44, 44.34, 44.09, 44.15, 43.61, 44.33, 44.83, 45.10,
              45.42, 45.84, 46.08, 45.89, 46.03, 45.61, 46.28, 46.28]
    r = indicators.rsi(prices, 14)
    assert r is not None and 0 <= r <= 100


def test_vwap():
    assert indicators.vwap([10, 20], [1, 1]) == 15.0
    assert indicators.vwap([10, 20], [0, 0]) is None


def test_crossover():
    assert indicators.crossover(9, 11, 10, 10) == "golden"
    assert indicators.crossover(11, 9, 10, 10) == "death"
    assert indicators.crossover(11, 12, 10, 10) is None


def test_signal_uptrend_buy():
    # Steadily rising prices -> uptrend, should not be a sell.
    closes = [float(i) for i in range(1, 260)]
    sig = signals.evaluate("TEST", closes)
    assert sig.action in ("buy", "hold")
    assert sig.metrics["rsi"] == 100.0


def test_risk_rejects_exclusion():
    cfg = risk.RiskConfig(exclusions=("GME",))
    plan = risk.plan_order(symbol="GME", side="buy", price=20, config=cfg,
                           portfolio_value=10000, buying_power=10000,
                           current_positions={})
    assert not plan.approved
    assert "exclusion" in plan.reason.lower()


def test_risk_caps_position_pct():
    cfg = risk.RiskConfig(max_position_pct=0.10, max_order_usd=10000, min_order_usd=10)
    plan = risk.plan_order(symbol="AAPL", side="buy", price=100, config=cfg,
                           portfolio_value=10000, buying_power=10000,
                           current_positions={})
    # 10% of 10k = $1000 -> 10 shares at $100
    assert plan.approved
    assert math.isclose(plan.notional_usd, 1000, rel_tol=1e-6)


def test_risk_sell_without_position():
    cfg = risk.RiskConfig()
    plan = risk.plan_order(symbol="AAPL", side="sell", price=100, config=cfg,
                           portfolio_value=10000, buying_power=10000,
                           current_positions={})
    assert not plan.approved


def test_news_bullish_and_bearish():
    bull = news.classify("AAPL", ["Apple beats estimates, raises guidance"])
    assert bull.tilt > 0 and not bull.block
    bear = news.classify("XYZ", ["XYZ misses, faces lawsuit and downgrade"])
    assert bear.tilt < 0


def test_news_tilt_clamped():
    headlines = ["beats upgrade buyback record revenue wins contract partnership"]
    assert news.classify("AAA", headlines).tilt == 2


def test_news_halt_blocks():
    a = news.classify("HALT", ["Trading halt issued; SEC investigation opened"])
    assert a.block is True


def test_signal_news_tilt_applied():
    closes = [100.0] * 260
    base = signals.evaluate("FLAT", closes)
    tilted = signals.evaluate("FLAT", closes, news_tilt=2)
    assert tilted.score == base.score + 2


def test_risk_pdt_guard():
    cfg = risk.RiskConfig(pdt_protection=True)
    plan = risk.plan_order(symbol="AAPL", side="buy", price=100, config=cfg,
                           portfolio_value=10000, buying_power=10000,
                           current_positions={}, day_trades_used=3,
                           account_value=5000)
    assert not plan.approved
    assert "pdt" in plan.reason.lower()
