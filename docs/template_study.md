# Template study — what we learned from existing bots

Notes distilled from reviewing several open-source Robinhood / stock trading
bots. The goal was to extract the *common, battle-tested* ideas and avoid each
project's individual mistakes.

## Repos reviewed

| Repo | Stars | Takeaway |
|---|---|---|
| `siropkin/robinhood-ai-trading-bot` | ~116 | LLM-in-the-loop design; demo/manual/auto modes; PDT + budget guards. Closest to our use case. |
| `2018kguo/RobinhoodBot` | ~227 | Classic SMA golden/death-cross swing strategy; clean config split. |
| `Jake0303/RobinHood-RSI-Trading-Bot` | ~220 | Minimal RSI mean-reversion loop. Good for understanding the bare signal. |
| `abghorba/Robinhood-Trading-Bot` | ~90 | Base-class + per-strategy modules (SMA, VWAP, sentiment). Good structure. |

## Common signal toolkit (the consensus)

Almost every bot converges on the same handful of indicators:

1. **Moving-average crossover (SMA/EMA)** — 50-day vs 200-day.
   - *Golden cross* (short crosses above long) → bullish.
   - *Death cross* (short crosses below long) → bearish.
2. **RSI (14)** — momentum / mean reversion.
   - `< 30` oversold (potential buy), `> 70` overbought (potential sell).
3. **VWAP** — price below VWAP = relatively cheap intraday; above = rich.
4. **Analyst ratings** (Robinhood-provided) — buy/hold/sell aggregate as a tilt.

We implement 1–3 deterministically in `src/indicators.py`. Analyst ratings are an
agent-side qualitative input (fetched via MCP `search`/quotes context).

## Risk-management patterns worth copying

- **Operating modes**: `demo` (no orders), `manual` (preview + confirm), `auto`.
- **PDT guard** *(now largely historical)*: the SEC abolished the PDT rule
  effective 2026-06-04, so the old "3 day trades / 5 days under $25k" cap no
  longer applies. We keep an optional guard for brokers still in their phase-in
  window, but the real brake is a self-imposed **daily loss limit**.
- **Hard caps**: max % of portfolio per position, max number of positions,
  min/max dollar size per order, exclusion list.
- **Idempotency / no double-fills**: check `get_equity_orders` for an existing
  open order on the same symbol before placing another.

## Anti-patterns we deliberately avoid

- **Letting the LLM compute indicators.** It hallucinates numbers. Compute in code.
- **Storing username/password.** The MCP uses OAuth; there are no credentials to store.
- **Tight polling loops that ignore market hours / PDT.** Respect both.
- **Placing orders without a dry-run.** Always `review_equity_order` first.
- **Sentiment-from-Twitter as a primary signal** (used by one repo) — noisy and
  fragile; at most a minor tilt, never a trigger.
</content>
