# China-Hawks 🦅 — Robinhood Agentic Trading Kit

A small, auditable toolkit for running an **AI trading agent against Robinhood's
Trading MCP** (`agent.robinhood.com/mcp/trading`).

The design philosophy, learned from studying several open-source trading bots:

> **The LLM agent decides *what* to consider and *when* to act. The deterministic
> Python core decides the *numbers*.** The agent must never invent an RSI value,
> a moving average, or a position size — it computes them with the code in `src/`
> and is bound by the limits in `config.yaml`.

## What's here

| Path | What it is |
|---|---|
| `docs/robinhood_mcp.md` | Reference for the Robinhood MCP tools + the mandatory order workflow |
| `docs/template_study.md` | Notes distilled from the open-source bots this kit learned from |
| `AGENT_PLAYBOOK.md` | The operating manual the agent (me) follows each trading cycle |
| `config.example.yaml` | Risk limits, watchlist, mode. Copy to `config.yaml` |
| `src/indicators.py` | Pure-Python SMA, EMA, RSI, VWAP, crossover detection |
| `src/signals.py` | Combines indicators (+ news tilt) into a `buy` / `sell` / `hold` signal |
| `src/news.py` | Deterministic news-catalyst classifier (bullish/bearish/halt-risk) |
| `src/risk.py` | Position sizing, exposure caps, exclusions, PDT guard |
| `src/nba_trade.py` | NBA-style trade balancer: CBA salary-match validation + counter-offer search (NBA 2K trade-finder style) |
| `tests/` | Unit tests for the deterministic core |

## How it fits together

```
Robinhood MCP                 This repo (deterministic)            Agent (LLM)
─────────────                 ─────────────────────────            ───────────
get_equity_quotes  ──prices──▶  indicators.py ──▶ signals.py ──┐
get_equity_positions ─state──▶  risk.py  (sizing + caps)  ◀────┘  decides per
get_accounts ──buying power─▶        │                            config + signal
                                     ▼
review_equity_order  ◀──── proposed order ──── (agent shows preview)
place_equity_order   ◀──── only after review + (manual) confirmation
```

## Quick start

```bash
# 1. connect the MCP (one time)
claude mcp add robinhood-trading --transport http https://agent.robinhood.com/mcp/trading

# 2. set your risk profile
cp config.example.yaml config.yaml   # then edit limits + watchlist

# 3. run the tests for the deterministic core
python -m pytest tests/ -q
```

Then tell the agent to run a cycle following `AGENT_PLAYBOOK.md`.

## ⚠️ Disclaimer

Educational software. Trading equities carries real risk of loss. Start in
`demo` mode, fund the Agentic account with only what you can afford to lose, and
keep `auto` mode off until you trust the behavior.
</content>
