"""Fetch real NBA rosters + contract salaries from ESPN's public API.

ESPN's roster endpoint (undocumented but public, no key required) returns each
player's current-season salary, years remaining, and trade restriction flags
straight from their contract data — no scraping or manual entry needed. We
don't have a free source for skill ratings, so ``Player.rating`` is left at 0
for fetched data; ``nba_trade.suggest_counter_offers`` degrades gracefully to
ranking purely by salary fit when ratings are unknown.

Results are cached to a local JSON file so the web app doesn't refetch on
every request; call ``refresh()`` to pull fresh data.
"""

from __future__ import annotations

import json
import os
import urllib.request
from typing import Optional

from .nba_trade import Player

TEAMS_URL = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams?limit=40"
ROSTER_URL = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams/{team_id}/roster"

CACHE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "nba_rosters.json")


def _fetch_json(url: str, timeout: float = 15.0) -> dict:
    with urllib.request.urlopen(url, timeout=timeout) as resp:
        return json.loads(resp.read())


def _fetch_team_list() -> list[dict]:
    data = _fetch_json(TEAMS_URL)
    return [t["team"] for t in data["sports"][0]["leagues"][0]["teams"]]


def _fetch_team_roster(team_id: str) -> list[dict]:
    data = _fetch_json(ROSTER_URL.format(team_id=team_id))
    return data.get("athletes", [])


def refresh() -> dict[str, list[Player]]:
    """Pull every team's live roster + salary data from ESPN and cache it."""
    rosters: dict[str, list[Player]] = {}
    raw: dict[str, list[dict]] = {}

    for team in _fetch_team_list():
        team_name = team["displayName"]
        athletes = _fetch_team_roster(team["id"])
        players = []
        raw_players = []
        for a in athletes:
            contract = a.get("contract") or {}
            salary = contract.get("salary")
            if not salary:
                continue  # two-way / non-standard contracts without a cap number
            players.append(
                Player(
                    name=a["fullName"],
                    team=team_name,
                    salary=float(salary),
                    rating=0,
                    age=a.get("age", 0) or 0,
                    years_left=contract.get("yearsRemaining", 0) or 0,
                )
            )
            raw_players.append(
                {
                    "name": a["fullName"],
                    "salary": salary,
                    "age": a.get("age", 0) or 0,
                    "years_left": contract.get("yearsRemaining", 0) or 0,
                    "position": (a.get("position") or {}).get("abbreviation", ""),
                }
            )
        rosters[team_name] = players
        raw[team_name] = raw_players

    os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
    with open(CACHE_PATH, "w") as f:
        json.dump(raw, f, indent=2)

    return rosters


def load_cached() -> Optional[dict[str, list[Player]]]:
    """Load rosters from the local cache, or None if it hasn't been fetched yet."""
    if not os.path.exists(CACHE_PATH):
        return None
    with open(CACHE_PATH) as f:
        raw = json.load(f)
    return {
        team: [
            Player(
                name=p["name"],
                team=team,
                salary=float(p["salary"]),
                rating=0,
                age=p.get("age", 0),
                years_left=p.get("years_left", 0),
            )
            for p in players
        ]
        for team, players in raw.items()
    }
