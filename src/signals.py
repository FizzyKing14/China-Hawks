"""Combine indicators into a single, explainable trade signal.

This is intentionally simple and transparent: each rule contributes a vote and a
human-readable reason. The agent surfaces these reasons to the user instead of
saying "the AI thinks so".
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Sequence

from . import indicators


@dataclass
class Signal:
    symbol: str
    action: str            # "buy" | "sell" | "hold"
    score: int             # net of bullish (+) and bearish (-) votes
    reasons: list[str] = field(default_factory=list)
    metrics: dict = field(default_factory=dict)


def evaluate(
    symbol: str,
    closes: Sequence[float],
    volumes: Optional[Sequence[float]] = None,
    *,
    short_period: int = 50,
    long_period: int = 200,
    rsi_period: int = 14,
    rsi_oversold: float = 30.0,
    rsi_overbought: float = 70.0,
    buy_threshold: int = 2,
    sell_threshold: int = -2,
    news_tilt: int = 0,
) -> Signal:
    """Score a symbol from its price history.

    ``closes`` must be ordered oldest -> newest. ``volumes`` (intraday, same
    length as a recent price window) enables the VWAP rule when provided.
    ``news_tilt`` (typically from ``news.classify``, clamped to [-2, 2]) folds a
    fresh-headline catalyst into the score.
    """
    reasons: list[str] = []
    score = 0
    metrics: dict = {}

    # --- Moving-average crossover -------------------------------------------
    short_series = indicators.ema_series(closes, short_period)
    long_series = indicators.ema_series(closes, long_period)
    if len(closes) >= long_period + 1 and short_series[-2] is not None and long_series[-2] is not None:
        cross = indicators.crossover(
            short_series[-2], short_series[-1], long_series[-2], long_series[-1]
        )
        metrics["ema_short"] = round(short_series[-1], 4)
        metrics["ema_long"] = round(long_series[-1], 4)
        if cross == "golden":
            score += 2
            reasons.append(f"Golden cross: EMA{short_period} crossed above EMA{long_period}")
        elif cross == "death":
            score -= 2
            reasons.append(f"Death cross: EMA{short_period} crossed below EMA{long_period}")
        elif short_series[-1] > long_series[-1]:
            score += 1
            reasons.append(f"Uptrend: EMA{short_period} above EMA{long_period}")
        else:
            score -= 1
            reasons.append(f"Downtrend: EMA{short_period} below EMA{long_period}")

    # --- RSI -----------------------------------------------------------------
    r = indicators.rsi(closes, rsi_period)
    if r is not None:
        metrics["rsi"] = round(r, 2)
        if r < rsi_oversold:
            score += 1
            reasons.append(f"RSI {r:.1f} < {rsi_oversold:g} (oversold)")
        elif r > rsi_overbought:
            score -= 1
            reasons.append(f"RSI {r:.1f} > {rsi_overbought:g} (overbought)")

    # --- VWAP ----------------------------------------------------------------
    if volumes:
        window = closes[-len(volumes):]
        vw = indicators.vwap(window, volumes)
        if vw is not None and closes:
            metrics["vwap"] = round(vw, 4)
            last = closes[-1]
            if last < vw:
                score += 1
                reasons.append(f"Price {last:.2f} below VWAP {vw:.2f} (cheap)")
            elif last > vw:
                score -= 1
                reasons.append(f"Price {last:.2f} above VWAP {vw:.2f} (rich)")

    # --- News catalyst -------------------------------------------------------
    if news_tilt:
        score += news_tilt
        metrics["news_tilt"] = news_tilt
        direction = "bullish" if news_tilt > 0 else "bearish"
        reasons.append(f"News tilt {news_tilt:+d} ({direction})")

    if not reasons:
        reasons.append("Insufficient history for a signal")

    if score >= buy_threshold:
        action = "buy"
    elif score <= sell_threshold:
        action = "sell"
    else:
        action = "hold"

    return Signal(symbol=symbol, action=action, score=score, reasons=reasons, metrics=metrics)
