"""NBA-style trade balancer: salary-cap validation + counter-offer search.

Same philosophy as the rest of ``src/``: this module never invents a number.
Given player data supplied by the caller (real roster/salary figures, or a
fantasy-league's scoring stats standing in for "salary"), it deterministically
checks whether a trade clears the NBA's salary-matching rule and searches the
offering side's roster for combinations that would balance a target trade —
the "other team counters with a fair offer" behavior from NBA 2K's trade
finder / ESPN's Trade Machine.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from itertools import combinations
from typing import Sequence


@dataclass(frozen=True)
class Player:
    name: str
    team: str
    salary: float
    rating: int = 75          # overall rating, 0-99
    age: int = 25
    years_left: int = 1
    headshot: str = ""        # image URL, empty when unknown
    position: str = ""
    jersey: str = ""


@dataclass
class TradeValidation:
    approved: bool
    side_a_salary: float
    side_b_salary: float
    reason: str


@dataclass
class TradeOffer:
    give: tuple[Player, ...]
    get: tuple[Player, ...]
    score: float
    reasons: list[str] = field(default_factory=list)


def validate_trade(
    side_a_sends: Sequence[Player],
    side_b_sends: Sequence[Player],
    *,
    match_pct: float = 1.25,
    apron_buffer: float = 100_000.0,
) -> TradeValidation:
    """Check the simplified NBA CBA salary-matching rule for a two-side trade.

    Each side's incoming salary must not exceed ``match_pct`` of what it sent
    out, plus ``apron_buffer``. Both directions are checked independently
    since the threshold isn't symmetric when salaries differ.
    """
    a_out = sum(p.salary for p in side_a_sends)
    b_out = sum(p.salary for p in side_b_sends)

    a_receives_too_much = b_out > a_out * match_pct + apron_buffer
    b_receives_too_much = a_out > b_out * match_pct + apron_buffer

    if a_receives_too_much:
        reason = (
            f"Side sending ${a_out:,.0f} can't receive ${b_out:,.0f} "
            f"(cap: ${a_out * match_pct + apron_buffer:,.0f})"
        )
        return TradeValidation(False, a_out, b_out, reason)
    if b_receives_too_much:
        reason = (
            f"Side sending ${b_out:,.0f} can't receive ${a_out:,.0f} "
            f"(cap: ${b_out * match_pct + apron_buffer:,.0f})"
        )
        return TradeValidation(False, a_out, b_out, reason)

    return TradeValidation(True, a_out, b_out, "Salaries match within CBA threshold")


def suggest_counter_offers(
    want: Sequence[Player],
    offer_pool: Sequence[Player],
    *,
    max_players_out: int = 3,
    top_n: int = 3,
    match_pct: float = 1.25,
    apron_buffer: float = 100_000.0,
) -> list[TradeOffer]:
    """Search ``offer_pool`` for combinations that would balance a trade for ``want``.

    Mirrors "the other team counters with a fair package": every combination
    of 1..``max_players_out`` players from ``offer_pool`` is checked against
    the CBA salary rule, then ranked by how closely it matches ``want`` on
    salary and total rating. Only trades that pass ``validate_trade`` are
    returned.
    """
    want_salary = sum(p.salary for p in want)
    want_rating = sum(p.rating for p in want)

    candidates: list[TradeOffer] = []
    for size in range(1, max_players_out + 1):
        for combo in combinations(offer_pool, size):
            validation = validate_trade(
                combo, want, match_pct=match_pct, apron_buffer=apron_buffer
            )
            if not validation.approved:
                continue

            give_salary = sum(p.salary for p in combo)
            give_rating = sum(p.rating for p in combo)
            salary_gap = abs(give_salary - want_salary)
            rating_gap = abs(give_rating - want_rating)

            # Higher is better: penalize salary mismatch (in $100k units) and
            # rating imbalance more heavily since it drives fit, not just cap math.
            score = 100.0 - (salary_gap / 100_000.0) - (rating_gap * 2.0)

            reasons = [
                f"Salary ${give_salary:,.0f} vs ${want_salary:,.0f} (gap ${salary_gap:,.0f})",
                f"Rating {give_rating} vs {want_rating} (gap {rating_gap})",
                validation.reason,
            ]
            candidates.append(TradeOffer(give=combo, get=tuple(want), score=round(score, 2), reasons=reasons))

    candidates.sort(key=lambda offer: offer.score, reverse=True)
    return candidates[:top_n]
