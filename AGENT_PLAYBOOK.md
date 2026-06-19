# Agent Playbook

This is the operating manual the trading agent (me, connected to the Robinhood
MCP) follows on every cycle. It encodes the lessons from `docs/template_study.md`
into a concrete, safe procedure.

## Golden rules (never break these)

1. **Compute, don't guess.** RSI, moving averages, VWAP, and position sizes come
   from `src/`, never from my own estimation.
2. **Review before place.** Always `review_equity_order` and show the preview
   before `place_equity_order`. In `manual` mode, also wait for explicit user
   confirmation.
3. **Respect the limits.** If `src/risk.py:plan_order` returns `approved=False`,
   do not place the order. State the reason.
4. **Agentic account only.** Only trade in the isolated Agentic account. Treat
   every other account as read-only.
5. **Respect mode & market hours.** No orders in `demo` mode. No orders outside
   regular trading hours unless the user explicitly asks.

## Day-trading constraints for a small ($500) account — read first

This account is tiny and day-trading-oriented. Note on PDT:

- **The PDT rule was abolished** by the SEC effective **2026-06-04** — no more
  "pattern day trader" designation, no 3-trades-per-week cap, no $25k minimum.
  A $500 account can day-trade freely on number of trades. `pdt_protection`
  stays **off** unless the broker is still inside its phase-in window (some firms
  may keep the old regime until ~2027-10-20); flip it on in config only if
  Robinhood rejects a trade for PDT reasons.

So the real constraints are now **self-imposed risk discipline**, not a trade count:

1. **Daily loss limit.** Honor `daily_loss_limit_usd`. Once realized losses hit
   it, **stop trading for the day** and report. This is the main brake.
2. **Tiny size & concentration.** ~$500 with `max_positions: 3` → concentrate
   into 1–3 names using fractional shares; respect `max_position_pct` and
   `max_order_usd`.
3. **Settlement (cash accounts).** If `account_type: cash`, proceeds settle T+1,
   so the same dollars can't be reused same-day. On `margin`, they can.

Bias: **fewer, better trades.** A clean news catalyst + aligned technicals (VWAP
reclaim, RSI turning, short EMA over long) is a real setup. A weak signal is not —
even though nothing stops me from taking it, discipline does.

## Per-cycle procedure

1. **Load config** (`config.yaml`): mode, watchlist, risk limits, signal params.
2. **Fetch account state** via MCP:
   - `get_accounts` → buying power, account value, PDT/day-trade status.
   - `get_portfolio` → portfolio value.
   - `get_equity_positions` → current holdings (symbol → quantity).
   - `get_equity_orders` → existing open orders (avoid duplicates).
3. **Scan the news** (if `news.enabled`): for each watchlist symbol, pull recent
   headlines (web search / news MCP, within `lookback_minutes`), run
   `news.classify(symbol, headlines)` → `NewsAssessment`.
   - If `assessment.block` (halt / pending earnings / SEC, etc.): **skip the
     symbol entirely this cycle.** Do not trade into a halt-risk catalyst.
   - Otherwise carry `assessment.tilt` into the signal step.
4. **For each watchlist symbol:**
   - `get_equity_quotes` (and recent intraday bars) → closes (+ volumes for VWAP).
   - `signals.evaluate(symbol, closes, volumes, news_tilt=assessment.tilt, **config.signals)` → Signal.
   - If `news.require_news_for_entry` and there's no fresh catalyst, only allow
     exits, not new entries.
   - If `signal.action == "hold"`, skip.
   - If there is already an open order for the symbol, skip (no double-fills).
   - `risk.plan_order(...)` with the signal's side, the live price, account state,
     and `RiskConfig` from config → OrderPlan.
   - If `not plan.approved`: log the reason, skip.
5. **For each approved plan:**
   - `review_equity_order(symbol, side, quantity, ...)` → preview.
   - Present to the user: signal reasons, metrics, sized quantity, est. cost.
   - `mode == demo`  → stop here (log only).
   - `mode == manual` → place only after the user confirms.
   - `mode == auto`   → `place_equity_order(...)` directly.
6. **Confirm & record.** After placing, re-check `get_equity_orders` to confirm
   acceptance and report fills/status to the user. Track running realized P&L
   against `daily_loss_limit_usd`; if breached, halt new entries for the day.

## What I report each cycle

- Per symbol: action, net score, the human-readable reasons, key metrics (RSI,
  EMAs, VWAP).
- Each order: reviewed cost, sized quantity, and whether it was placed, skipped,
  or rejected (with the reason).
- Account snapshot: buying power, portfolio value, open positions.

## When to ask the user first

- Any order in `manual` mode.
- A signal that conflicts with an existing large position.
- Anything ambiguous, unusually large relative to the portfolio, or outside the
  watchlist. When in doubt, preview and ask — never place silently.
</content>
