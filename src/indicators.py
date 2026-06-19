"""Pure-Python technical indicators.

No third-party dependencies so the agent can compute signals anywhere. All
functions take a list of closing prices ordered oldest -> newest and either
return a single latest value or a list aligned to the input (with ``None`` where
there isn't enough history yet).
"""

from __future__ import annotations

from typing import Optional, Sequence


def sma(prices: Sequence[float], period: int) -> Optional[float]:
    """Simple moving average of the most recent ``period`` prices."""
    if period <= 0:
        raise ValueError("period must be positive")
    if len(prices) < period:
        return None
    window = prices[-period:]
    return sum(window) / period


def ema_series(prices: Sequence[float], period: int) -> list[Optional[float]]:
    """Exponential moving average aligned to ``prices``.

    Seeded with the SMA of the first ``period`` values, as is conventional.
    """
    if period <= 0:
        raise ValueError("period must be positive")
    out: list[Optional[float]] = [None] * len(prices)
    if len(prices) < period:
        return out
    k = 2 / (period + 1)
    prev = sum(prices[:period]) / period
    out[period - 1] = prev
    for i in range(period, len(prices)):
        prev = prices[i] * k + prev * (1 - k)
        out[i] = prev
    return out


def ema(prices: Sequence[float], period: int) -> Optional[float]:
    """Latest EMA value, or ``None`` if there isn't enough history."""
    return ema_series(prices, period)[-1] if prices else None


def rsi(prices: Sequence[float], period: int = 14) -> Optional[float]:
    """Wilder's RSI for the most recent ``period`` window.

    Returns a value in [0, 100], or ``None`` if history is insufficient.
    """
    if period <= 0:
        raise ValueError("period must be positive")
    if len(prices) < period + 1:
        return None

    gains = 0.0
    losses = 0.0
    for i in range(1, period + 1):
        change = prices[i] - prices[i - 1]
        if change >= 0:
            gains += change
        else:
            losses -= change
    avg_gain = gains / period
    avg_loss = losses / period

    # Wilder smoothing across the remaining history.
    for i in range(period + 1, len(prices)):
        change = prices[i] - prices[i - 1]
        gain = max(change, 0.0)
        loss = max(-change, 0.0)
        avg_gain = (avg_gain * (period - 1) + gain) / period
        avg_loss = (avg_loss * (period - 1) + loss) / period

    if avg_loss == 0:
        return 100.0
    rs = avg_gain / avg_loss
    return 100 - (100 / (1 + rs))


def vwap(prices: Sequence[float], volumes: Sequence[float]) -> Optional[float]:
    """Volume-weighted average price over the supplied intraday bars."""
    if len(prices) != len(volumes):
        raise ValueError("prices and volumes must be the same length")
    total_vol = sum(volumes)
    if total_vol == 0:
        return None
    return sum(p * v for p, v in zip(prices, volumes)) / total_vol


def crossover(short_prev: float, short_now: float,
              long_prev: float, long_now: float) -> Optional[str]:
    """Detect a moving-average crossover between two consecutive points.

    Returns ``"golden"`` (short crosses above long, bullish),
    ``"death"`` (short crosses below long, bearish), or ``None``.
    """
    was_below = short_prev <= long_prev
    now_above = short_now > long_now
    if was_below and now_above:
        return "golden"
    was_above = short_prev >= long_prev
    now_below = short_now < long_now
    if was_above and now_below:
        return "death"
    return None
