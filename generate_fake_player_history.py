import json
import random
from pathlib import Path


def build_fake_history(output_path: Path, year: int = 2026, weeks: int = 30) -> None:
    random.seed(42)
    player_names = [
        "Anna",
        "Bjørn",
        "Carina",
        "Daniel",
        "Emilie",
        "Fredrik",
        "Guro",
        "Henrik",
        "Ingrid",
        "Jørgen",
    ]

    players: dict[str, list[dict[str, object]]] = {}

    for idx, name in enumerate(player_names):
        base = 110 + idx * 12
        trend = idx * 3
        records: list[dict[str, object]] = []

        for week in range(1, weeks + 1):
            weekly_noise = random.uniform(-14, 18)
            seasonal = (week % 5) * 6
            form = max(0, 18 - abs((week % 7) - 3) * 2)
            total = round(base + trend + week * 8 + seasonal + weekly_noise + form)
            records.append(
                {
                    "week": week,
                    "year": year,
                    "plass": 0,
                    "totalt": total,
                    "navn": name,
                }
            )
        players[name] = records

    for week in range(1, weeks + 1):
        weekly_rows = [
            {"name": name, "totalt": player_rows[week - 1]["totalt"]}
            for name, player_rows in players.items()
        ]
        weekly_rows.sort(key=lambda item: item["totalt"], reverse=True)
        for place, row in enumerate(weekly_rows, start=1):
            for name, player_rows in players.items():
                if player_rows[week - 1]["navn"] == row["name"]:
                    player_rows[week - 1]["plass"] = place
                    break

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump({"players": players}, handle, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    build_fake_history(Path("data/player_history.json"))
    print("Created fake weekly history for weeks 1-30 at data/player_history.json")
