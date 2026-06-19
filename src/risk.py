"""Risk management: turn a signal + account state + config into an order plan.

Every order the agent places must pass through ``plan_order``. If it returns a
rejection, the agent does not call ``place_equity_order``. Period.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class RiskConfig:
    max_position_pct: float = 0.10      # max fraction of portfolio per symbol
    max_positions: int = 10            # max distinct holdings
    min_order_usd: float = 10.0
    max_order_usd: float = 1000.0
    exclusions: tuple[str, ...] = ()   # symbols never to trade
    # PDT was abolished by the SEC (eff. 2026-06-04). Default off. Only set True
    # if your broker is still inside its phase-in window (up to 2027-10-20) and
    # enforces the old 3-day-trades / $25k regime.
    pdt_protection: bool = False


@dataclass
class OrderPlan:
    symbol: str
    side: str                # "buy" | "sell"
    approved: bool
    quantity: float = 0.0
    notional_usd: float = 0.0
    reason: str = ""


def plan_order(
    *,
    symbol: str,
    side: str,
    price: float,
    config: RiskConfig,
    portfolio_value: float,
    buying_power: float,
    current_positions: dict[str, float],     # symbol -> quantity held
    day_trades_used: int = 0,
    account_value: float = 0.0,
) -> OrderPlan:
    """Validate and size an order against hard risk limits.

    Returns an ``OrderPlan`` with ``approved=False`` and a ``reason`` whenever a
    limit is hit, so the rejection is always explainable to the user.
    """
    symbol = symbol.upper()

    if side not in ("buy", "sell"):
        return OrderPlan(symbol, side, False, reason=f"Unknown side '{side}'")

    if symbol in {s.upper() for s in config.exclusions}:
        return OrderPlan(symbol, side, False, reason=f"{symbol} is on the exclusion list")

    if price <= 0:
        return OrderPlan(symbol, side, False, reason="Non-positive price")

    held = current_positions.get(symbol, 0.0)

    # --- Legacy PDT guard (only if broker still in phase-in; default off) ----
    if config.pdt_protection and account_value < 25_000 and day_trades_used >= 3:
        return OrderPlan(
            symbol, side, False,
            reason=f"PDT guard: {day_trades_used} day trades used on a sub-$25k account",
        )

    # --- SELL ---------------------------------------------------------------
    if side == "sell":
        if held <= 0:
            return OrderPlan(symbol, side, False, reason=f"No position in {symbol} to sell")
        notional = held * price
        return OrderPlan(symbol, side, True, quantity=held, notional_usd=notional,
                         reason="Closing full position")

    # --- BUY ----------------------------------------------------------------
    if held == 0 and len(current_positions) >= config.max_positions:
        return OrderPlan(symbol, side, False,
                         reason=f"At max_positions ({config.max_positions})")

    # Cap by portfolio %: don't let this symbol exceed max_position_pct.
    target_cap = portfolio_value * config.max_position_pct
    current_exposure = held * price
    room_by_pct = max(0.0, target_cap - current_exposure)

    budget = min(room_by_pct, config.max_order_usd, buying_power)

    if budget < config.min_order_usd:
        return OrderPlan(
            symbol, side, False,
            reason=(f"Budget ${budget:.2f} below min ${config.min_order_usd:.2f} "
                    f"(pct cap, max order, or buying power)"),
        )

    quantity = budget / price
    return OrderPlan(symbol, side, True, quantity=round(quantity, 6),
                     notional_usd=round(quantity * price, 2),
                     reason=f"Sized to ${budget:.2f} within limits")
