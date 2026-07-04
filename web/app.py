"""Minimal Flask front end for the NBA trade balancer in src/nba_trade.py.

Real rosters + salaries come from src/nba_data.py (ESPN's public roster API),
cached to data/nba_rosters.json. All trade-legality and ranking logic lives in
src/nba_trade.py — this file only wires HTTP requests to it.
"""

from __future__ import annotations

import os
import sys

from flask import Flask, redirect, render_template, request, url_for

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import nba_data  # noqa: E402
from src.nba_trade import suggest_counter_offers, validate_trade  # noqa: E402

app = Flask(__name__)


def get_rosters() -> dict:
    rosters = nba_data.load_cached()
    if rosters is None:
        rosters = nba_data.refresh()
    return rosters


def render_page(team_a: str = "", team_b: str = "", validation=None, offers=None, error=None):
    rosters = get_rosters()
    teams = sorted(rosters.keys())
    roster_a = sorted(rosters.get(team_a, []), key=lambda p: -p.salary)
    roster_b = sorted(rosters.get(team_b, []), key=lambda p: -p.salary)
    return render_template(
        "index.html",
        teams=teams,
        team_a=team_a,
        team_b=team_b,
        roster_a=roster_a,
        roster_b=roster_b,
        validation=validation,
        offers=offers,
        error=error,
    )


@app.route("/")
def index():
    return render_page(request.args.get("team_a", ""), request.args.get("team_b", ""))


@app.route("/validate", methods=["POST"])
def validate():
    rosters = get_rosters()
    team_a, team_b = request.form["team_a"], request.form["team_b"]
    by_name_a = {p.name: p for p in rosters.get(team_a, [])}
    by_name_b = {p.name: p for p in rosters.get(team_b, [])}
    give = [by_name_a[n] for n in request.form.getlist("give") if n in by_name_a]
    get = [by_name_b[n] for n in request.form.getlist("get") if n in by_name_b]
    if not give or not get:
        return render_page(team_a, team_b, error="Pick at least one player on each side.")
    return render_page(team_a, team_b, validation=validate_trade(give, get))


@app.route("/suggest", methods=["POST"])
def suggest():
    rosters = get_rosters()
    team_a, team_b = request.form["team_a"], request.form["team_b"]
    by_name_b = {p.name: p for p in rosters.get(team_b, [])}
    want = [by_name_b[n] for n in request.form.getlist("want") if n in by_name_b]
    if not want:
        return render_page(team_a, team_b, error="Pick at least one player you want from the other team.")
    offer_pool = rosters.get(team_a, [])
    offers = suggest_counter_offers(want, offer_pool, max_players_out=3, top_n=5)
    return render_page(team_a, team_b, offers=offers)


@app.route("/refresh")
def refresh():
    nba_data.refresh()
    return redirect(url_for("index", team_a=request.args.get("team_a", ""), team_b=request.args.get("team_b", "")))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
