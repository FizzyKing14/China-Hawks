"""Fetch real NBA rosters + contract salaries from ESPN's public API.

ESPN's roster endpoint (undocumented but public, no key required) returns each
player's current-season salary, years remaining, and trade restriction flags
straight from their contract data — no scraping or manual entry needed. We
don't have a free source for skill ratings, so ``Player.rating`` is left at 0
for fetched data; ``nba_trade.suggest_counter_offers`` degrades gracefully to
ranking purely by salary fit when ratings are unknown.

ESPN's salary figure is occasionally stale for a recently re-signed player
(confirmed case: CJ McCollum's 1yr/$21M deal with Atlanta showed as $30.6M).
HoopsHype's per-team salary page tends to catch these sooner, so we
cross-check it and prefer its number when the two disagree — headshot,
position, and jersey number still come from ESPN, which HoopsHype doesn't
provide cleanly.

Results are cached to a local JSON file so the web app doesn't refetch on
every request; call ``refresh()`` to pull fresh data.
"""

from __future__ import annotations

import json
import os
import re
import urllib.request
from typing import Optional

from .nba_trade import Player

TEAMS_URL = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams?limit=40"
ROSTER_URL = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams/{team_id}/roster"
HOOPSHYPE_TEAM_URL = "https://hoopshype.com/salaries/teams/x/{team_id}/"
CURRENT_SEASON = 2026

CACHE_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "nba_rosters.json")

_NAME_SUFFIXES = (" jr.", " jr", " sr.", " sr", " iii", " ii", " iv", " v")


def _normalize_name(name: str) -> str:
    """Fold suffix variations so 'Jimmy Butler III' matches 'Jimmy Butler'."""
    name = name.lower().strip()
    for suffix in _NAME_SUFFIXES:
        if name.endswith(suffix):
            return name[: -len(suffix)].strip()
    return name


def _fetch_json(url: str, timeout: float = 15.0) -> dict:
    with urllib.request.urlopen(url, timeout=timeout) as resp:
        return json.loads(resp.read())


def _fetch_team_list() -> list[dict]:
    data = _fetch_json(TEAMS_URL)
    return [t["team"] for t in data["sports"][0]["leagues"][0]["teams"]]


def _fetch_team_roster(team_id: str) -> list[dict]:
    data = _fetch_json(ROSTER_URL.format(team_id=team_id))
    return data.get("athletes", [])


def _fetch_hoopshype_salaries(team_id: str) -> dict[str, float]:
    """Best-effort fetch of a team's current-season salaries from HoopsHype.

    Returns an empty dict (never raises) if the page layout changes or the
    request fails — callers should treat this as an optional cross-check,
    not a hard dependency.
    """
    try:
        req = urllib.request.Request(
            HOOPSHYPE_TEAM_URL.format(team_id=team_id),
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", "ignore")
        match = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.S)
        if not match:
            return {}
        queries = json.loads(match.group(1))["props"]["pageProps"]["dehydratedState"]["queries"]
    except Exception:
        return {}

    for query in queries:
        data = query.get("state", {}).get("data")
        if not isinstance(data, dict):
            continue
        contracts = (data.get("contracts") or {}).get("contracts")
        if not contracts or len(contracts) < 2:
            continue
        salaries: dict[str, float] = {}
        for c in contracts:
            for season in c.get("seasons", []):
                if season.get("season") == CURRENT_SEASON and not season.get("terminated"):
                    salaries[_normalize_name(c["playerName"])] = float(season["salary"])
                    break
        if salaries:
            return salaries
    return {}


def refresh() -> dict[str, list[Player]]:
    """Pull every team's live roster + salary data from ESPN and cache it."""
    rosters: dict[str, list[Player]] = {}
    raw: dict[str, list[dict]] = {}

    for team in _fetch_team_list():
        team_name = team["displayName"]
        athletes = _fetch_team_roster(team["id"])
        hoopshype_salaries = _fetch_hoopshype_salaries(team["id"])
        players = []
        raw_players = []
        for a in athletes:
            contract = a.get("contract") or {}
            salary = contract.get("salary")
            if not salary:
                continue  # two-way / non-standard contracts without a cap number
            hh_salary = hoopshype_salaries.get(_normalize_name(a["fullName"]))
            if hh_salary is not None:
                salary = hh_salary
            headshot = (a.get("headshot") or {}).get("href", "")
            position = (a.get("position") or {}).get("abbreviation", "")
            jersey = a.get("jersey", "") or ""
            players.append(
                Player(
                    name=a["fullName"],
                    team=team_name,
                    salary=float(salary),
                    rating=0,
                    age=a.get("age", 0) or 0,
                    years_left=contract.get("yearsRemaining", 0) or 0,
                    headshot=headshot,
                    position=position,
                    jersey=jersey,
                )
            )
            raw_players.append(
                {
                    "name": a["fullName"],
                    "salary": salary,
                    "age": a.get("age", 0) or 0,
                    "years_left": contract.get("yearsRemaining", 0) or 0,
                    "position": position,
                    "headshot": headshot,
                    "jersey": jersey,
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
                headshot=p.get("headshot", ""),
                position=p.get("position", ""),
                jersey=p.get("jersey", ""),
            )
            for p in players
        ]
        for team, players in raw.items()
    }
