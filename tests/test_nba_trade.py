import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.nba_trade import Player, suggest_counter_offers, validate_trade  # noqa: E402


def test_validate_trade_within_threshold():
    a = [Player("A1", "Team A", 10_000_000, rating=85)]
    b = [Player("B1", "Team B", 11_500_000, rating=84)]
    result = validate_trade(a, b)
    assert result.approved


def test_validate_trade_rejects_lopsided_salary():
    a = [Player("A1", "Team A", 2_000_000, rating=70)]
    b = [Player("B1", "Team B", 20_000_000, rating=90)]
    result = validate_trade(a, b)
    assert not result.approved
    assert "cap:" in result.reason


def test_suggest_counter_offers_finds_balanced_package():
    want = [Player("Star", "Team B", 30_000_000, rating=92)]
    offer_pool = [
        Player("Role1", "Team A", 13_000_000, rating=78),
        Player("Role2", "Team A", 12_000_000, rating=76),
        Player("Bench", "Team A", 2_000_000, rating=60),
        Player("Overpaid", "Team A", 40_000_000, rating=70),
    ]
    offers = suggest_counter_offers(want, offer_pool, max_players_out=2, top_n=3)
    assert offers
    assert all(o.score for o in offers)
    # the two mid-salary role players together should out-rank an overpriced single player
    best_names = {p.name for p in offers[0].give}
    assert best_names == {"Role1", "Role2"}


def test_suggest_counter_offers_empty_when_nothing_clears_cap():
    want = [Player("Star", "Team B", 30_000_000, rating=92)]
    offer_pool = [Player("Cheap", "Team A", 1_000_000, rating=60)]
    offers = suggest_counter_offers(want, offer_pool)
    assert offers == []
