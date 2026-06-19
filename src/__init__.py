"""Deterministic trading core for the Robinhood agentic kit.

The agent calls these functions to turn raw prices into signals and to size
positions within hard risk limits. Nothing here talks to a broker — that is the
agent's job via the Robinhood MCP tools.
"""

from . import indicators, news, risk, signals

__all__ = ["indicators", "signals", "risk", "news"]
