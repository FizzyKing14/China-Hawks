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
| `src/nba_data.py` | Fetches real rosters + contract salaries from ESPN's public API, cached to `data/nba_rosters.json` |
| `web/app.py` | Small Flask UI over `nba_trade.py` — pick two teams, validate a trade, or auto-suggest counter offers |
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

## NBA trade balancer (side project)

A small unrelated tool lives alongside the trading kit: an NBA trade
validator/suggester with a Flask UI.

```bash
pip install -r requirements.txt
python -c "from src import nba_data; nba_data.refresh()"   # pulls live rosters from ESPN
python web/app.py                                          # http://127.0.0.1:5000
```

Pick two teams, then either check a specific trade against the simplified CBA
salary-match rule, or select a player you want and get auto-suggested
counter-offer packages from the other team's roster — the "other team offers
a fair trade" behavior from NBA 2K's trade finder / ESPN's Trade Machine.
Salary data comes from ESPN's public roster API (no scraping, no API key);
skill ratings aren't available from a free source, so suggestions rank purely
on salary fit.

## ⚠️ Disclaimer

Educational software. Trading equities carries real risk of loss. Start in
`demo` mode, fund the Agentic account with only what you can afford to lose, and
keep `auto` mode off until you trust the behavior.
</content>
