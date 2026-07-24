import json

import pandas as pd

from blabla import build_player_history_entry, build_player_history_section_html


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


def test_build_player_history_section_html_includes_renderable_chart_data(tmp_path) -> None:
    history_path = tmp_path / "player_history.json"
    history_path.write_text(
        json.dumps(
            {
                "players": {
                    "ALICE": [
                        {"week": 1, "year": 2026, "plass": 1, "totalt": 100, "navn": "Alice"}
                    ]
                }
            }
        ),
        encoding="utf-8",
    )

    html = build_player_history_section_html(history_path)

    assert "ALICE" in html
    assert "const playerHistoryData =" in html
    assert "new Chart" in html
    assert "playerHistoryData[playerName]" in html
