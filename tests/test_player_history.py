import pandas as pd

from blabla import build_player_history_entry


def test_build_player_history_entry_ranks_and_filters_players() -> None:
    weekly_df = pd.DataFrame(
        [
            {"Navn": "Alice", "Totalt": 150},
            {"Navn": "Bob", "Totalt": 90},
            {"Navn": "Carol", "Totalt": 0},
        ]
    )

    entry = build_player_history_entry(31, 2026, weekly_df, "Totalt")

    assert entry["week"] == 31
    assert entry["year"] == 2026
    assert [row["Navn"] for row in entry["rows"][:2]] == ["Alice", "Bob"]
    assert entry["rows"][0]["Plass"] == 1
    assert entry["rows"][1]["Plass"] == 2
    assert entry["rows"][2]["Navn"] == "Carol"
    assert entry["rows"][2]["Totalt"] == 0
