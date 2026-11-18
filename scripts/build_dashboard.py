#!/usr/bin/env python3
"""Build a dependency-free, self-contained HTML dashboard from the supplied CSVs.

Only the Python standard library is used. End users do not need to run this
script: the generated index.html is ready to open in any modern browser.
"""
from __future__ import annotations

import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
TEMPLATE = ROOT / "dashboard_template.html"
OUTPUT = ROOT / "index.html"

FILES = {
    "deaths": "character-deaths.csv",
    "predictions": "character-predictions.csv",
    "battles": "battles.csv",
}
REQUIRED = {
    "deaths": {"Name", "Allegiances", "Death Year", "Book of Death", "Gender", "Nobility"},
    "predictions": {"actual", "pred", "alive", "name", "house", "culture", "isAlive", "isNoble", "popularity"},
    "battles": {"name", "year", "attacker_1", "attacker_outcome", "battle_type", "region", "location"},
}


def read_csv(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл данных: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if not reader.fieldnames:
            raise ValueError(f"В файле отсутствует строка заголовков: {path.name}")
        fields = list(reader.fieldnames)
        rows = [{k: (v or "").strip() for k, v in row.items()} for row in reader]
    return rows, fields


def remove_exact_duplicates(rows: list[dict[str, str]]) -> tuple[list[dict[str, str]], int]:
    seen: set[tuple[tuple[str, str], ...]] = set()
    cleaned: list[dict[str, str]] = []
    for row in rows:
        key = tuple(row.items())
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(row)
    return cleaned, len(rows) - len(cleaned)


def main() -> None:
    payload: dict[str, object] = {}
    stats: dict[str, int] = {}
    duplicate_stats: dict[str, int] = {}
    column_stats: dict[str, list[str]] = {}

    for key, filename in FILES.items():
        rows, columns = read_csv(RAW / filename)
        missing = REQUIRED[key] - set(columns)
        if missing:
            raise ValueError(f"В {filename} отсутствуют обязательные поля: {', '.join(sorted(missing))}")
        clean_rows, duplicate_count = remove_exact_duplicates(rows)
        payload[key] = clean_rows
        stats[key] = len(clean_rows)
        duplicate_stats[key] = duplicate_count
        column_stats[key] = columns

    payload["quality"] = {
        "stats": stats,
        "duplicates_removed": duplicate_stats,
        "columns": column_stats,
    }
    encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # Protect the HTML script-data element from accidental termination by a value.
    encoded = encoded.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    html = TEMPLATE.read_text(encoding="utf-8")
    if "__DATA_JSON__" not in html:
        raise ValueError("В dashboard_template.html не найден маркер __DATA_JSON__")
    OUTPUT.write_text(html.replace("__DATA_JSON__", encoded), encoding="utf-8")

    report = {
        "rows_after_exact_duplicate_removal": stats,
        "exact_duplicates_removed": duplicate_stats,
        "columns": {key: len(cols) for key, cols in column_stats.items()},
    }
    processed = ROOT / "data" / "processed"
    processed.mkdir(parents=True, exist_ok=True)
    (processed / "data_quality.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("Дашборд успешно собран:", OUTPUT)
    print("Размер файла:", f"{OUTPUT.stat().st_size:,}".replace(",", " "), "байт")
    for key, filename in FILES.items():
        print(f"  {filename}: {stats[key]} строк; удалено полных дублей: {duplicate_stats[key]}")
    print("Запуск: откройте index.html в браузере или дважды щелкните ЗАПУСТИТЬ_ДАШБОРД.bat")


if __name__ == "__main__":
    main()
