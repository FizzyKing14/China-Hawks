"""Lightweight, deterministic news-catalyst classifier.

Day trading is news-driven. The agent fetches fresh headlines at runtime (via
web search / a news MCP) and passes them here to get a *structured, explainable*
tilt instead of an opaque vibe. This is keyword-based on purpose: transparent,
testable, and fast. The agent's own judgment refines it — this just floors the
analysis in something reproducible.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterable

# Keywords are lowercased substrings. Order doesn't matter; all matches count.
_BULLISH = (
    "beats", "beat estimates", "raises guidance", "raised guidance", "upgrade",
    "upgraded", "record revenue", "record profit", "buyback", "approval",
    "approved", "wins contract", "awarded", "partnership", "acquire", "acquisition",
    "outperform", "surge", "soars", "tops estimates", "strong demand", "expansion",
)
_BEARISH = (
    "misses", "missed estimates", "cuts guidance", "lowers guidance", "downgrade",
    "downgraded", "recall", "lawsuit", "fraud", "investigation", "probe", "fine",
    "layoffs", "plunge", "plummets", "warns", "warning", "weak demand", "miss",
    "delisting", "bankruptcy", "default",
)
# High-uncertainty / "do not enter" catalysts — these gate trading regardless of
# direction. Day trading into these is how small accounts get wiped.
_HALT_RISK = (
    "halted", "trading halt", "circuit breaker", "sec investigation",
    "earnings after the bell", "earnings tonight", "fda decision pending",
    "going concern", "delisted",
)


@dataclass
class NewsAssessment:
    symbol: str
    tilt: int                          # net bullish (+) / bearish (-) signal votes
    block: bool = False                # if True, do NOT trade this name now
    reasons: list[str] = field(default_factory=list)


def _matches(text: str, keywords: Iterable[str]) -> list[str]:
    t = text.lower()
    return [k for k in keywords if k in t]


def classify(symbol: str, headlines: Iterable[str]) -> NewsAssessment:
    """Score a batch of headlines for a symbol.

    ``tilt`` is clamped to [-2, 2] so news never single-handedly dominates the
    technical signal. ``block=True`` means a halt-risk catalyst was detected and
    the agent should stand aside.
    """
    tilt = 0
    block = False
    reasons: list[str] = []

    for h in headlines:
        halt = _matches(h, _HALT_RISK)
        if halt:
            block = True
            reasons.append(f"HALT-RISK: '{h.strip()}' ({', '.join(halt)})")
            continue
        bull = _matches(h, _BULLISH)
        bear = _matches(h, _BEARISH)
        if bull:
            tilt += len(bull)
            reasons.append(f"+ '{h.strip()}' ({', '.join(bull)})")
        if bear:
            tilt -= len(bear)
            reasons.append(f"- '{h.strip()}' ({', '.join(bear)})")

    tilt = max(-2, min(2, tilt))
    if not reasons:
        reasons.append("No actionable headlines")
    return NewsAssessment(symbol=symbol.upper(), tilt=tilt, block=block, reasons=reasons)
