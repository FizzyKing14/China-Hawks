# Robinhood Trading MCP — Tool Reference

Endpoint: `https://agent.robinhood.com/mcp/trading` (HTTP transport, OAuth).

Add it to Claude Code:

```bash
claude mcp add robinhood-trading --transport http https://agent.robinhood.com/mcp/trading
```

## Scope & safety model

- **Equities only** during the beta (options tools are forthcoming).
- The agent can **only place trades in your isolated Robinhood Agentic account**.
  Every other account is read-only. The Agentic account starts **unfunded** — you
  move money in deliberately.
- Robinhood does **not** supervise the connected agent. All guardrails
  (position caps, exclusions, max trade size, PDT) must be enforced by *us*.
- Auth is via OAuth — the agent never sees your password.

## Tools

### Read
| Tool | Purpose |
|---|---|
| `get_accounts` | Account numbers, balances, buying power |
| `get_portfolio` | Portfolio value / equity / allocations |
| `get_equity_positions` | Current holdings, quantities, cost basis |
| `get_equity_quotes` | Live quotes for symbols |
| `get_equity_orders` | Order history / open orders |
| `search` | Look up instruments / symbols |

### Watchlists
| Tool | Purpose |
|---|---|
| `get_watchlists` | List watchlists |
| `add_to_watchlist` | Add a symbol |
| `update_watchlist` | Modify a watchlist |

### Trading
| Tool | Purpose |
|---|---|
| `review_equity_order` | **Dry-run / validate an order, return estimated cost & effects** |
| `place_equity_order` | Submit a real order |
| `cancel_equity_order` | Cancel an open order |

## The mandatory order pattern

```
1. get_equity_quotes(symbol)        -> current price
2. (compute signal locally via src/)
3. risk checks (config.yaml)        -> approved size or reject
4. review_equity_order(...)         -> show the user the preview
5. place_equity_order(...)          -> ONLY after review passes (and, in manual
                                        mode, after explicit user confirmation)
```

> Never call `place_equity_order` without a preceding `review_equity_order` in
> the same decision. This is the single most important rule.

> Note: exact parameter schemas for these tools are resolved at runtime from the
> live MCP server; treat the names above as stable and read each tool's schema
> before first use.
</content>
</invoke>
