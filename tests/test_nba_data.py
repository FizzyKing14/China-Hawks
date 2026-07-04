import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import nba_data  # noqa: E402

SAMPLE_TEAMS = [{"id": "1", "displayName": "Test Team"}]
SAMPLE_ROSTER = [
    {
        "fullName": "Fake Player",
        "age": 27,
        "position": {"abbreviation": "G"},
        "contract": {"salary": 12_000_000, "yearsRemaining": 2},
    },
    {
        "fullName": "No Contract Player",
        "age": 22,
        "position": {"abbreviation": "F"},
        "contract": None,
    },
]


def test_refresh_skips_players_without_salary(monkeypatch, tmp_path):
    monkeypatch.setattr(nba_data, "_fetch_team_list", lambda: SAMPLE_TEAMS)
    monkeypatch.setattr(nba_data, "_fetch_team_roster", lambda team_id: SAMPLE_ROSTER)
    monkeypatch.setattr(nba_data, "_fetch_hoopshype_salaries", lambda team_id: {})
    monkeypatch.setattr(nba_data, "CACHE_PATH", str(tmp_path / "rosters.json"))

    rosters = nba_data.refresh()

    assert list(rosters.keys()) == ["Test Team"]
    players = rosters["Test Team"]
    assert len(players) == 1
    assert players[0].name == "Fake Player"
    assert players[0].salary == 12_000_000
    assert players[0].years_left == 2


def test_refresh_prefers_hoopshype_salary_when_it_disagrees(monkeypatch, tmp_path):
    monkeypatch.setattr(nba_data, "_fetch_team_list", lambda: SAMPLE_TEAMS)
    monkeypatch.setattr(nba_data, "_fetch_team_roster", lambda team_id: SAMPLE_ROSTER)
    monkeypatch.setattr(nba_data, "_fetch_hoopshype_salaries", lambda team_id: {"fake player": 9_000_000.0})
    monkeypatch.setattr(nba_data, "CACHE_PATH", str(tmp_path / "rosters.json"))

    rosters = nba_data.refresh()

    assert rosters["Test Team"][0].salary == 9_000_000.0


def test_load_cached_round_trips(monkeypatch, tmp_path):
    monkeypatch.setattr(nba_data, "_fetch_team_list", lambda: SAMPLE_TEAMS)
    monkeypatch.setattr(nba_data, "_fetch_team_roster", lambda team_id: SAMPLE_ROSTER)
    monkeypatch.setattr(nba_data, "_fetch_hoopshype_salaries", lambda team_id: {})
    monkeypatch.setattr(nba_data, "CACHE_PATH", str(tmp_path / "rosters.json"))

    nba_data.refresh()
    cached = nba_data.load_cached()

    assert cached["Test Team"][0].name == "Fake Player"
    assert cached["Test Team"][0].salary == 12_000_000.0


def test_load_cached_returns_none_when_missing(monkeypatch, tmp_path):
    monkeypatch.setattr(nba_data, "CACHE_PATH", str(tmp_path / "missing.json"))
    assert nba_data.load_cached() is None


def test_normalize_name_folds_suffixes():
    assert nba_data._normalize_name("Jimmy Butler III") == "jimmy butler"
    assert nba_data._normalize_name("Jimmy Butler") == "jimmy butler"
    assert nba_data._normalize_name("Gary Payton II") == "gary payton"
