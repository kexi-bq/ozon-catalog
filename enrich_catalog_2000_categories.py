from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
CATALOG_PATH = APP_DIR / "data" / "catalog_2000.xlsx"
BACKUP_PATH = APP_DIR / "data" / "catalog_2000_before_category_enrich.xlsx"
REPORT_PATH = APP_DIR / "data" / "catalog_2000_category_enrich_report.xlsx"
SOURCE_REPORT_PATH = APP_DIR.parent / "reports" / "ozon_upload_4500_candidates.xlsx"


def clean_str(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def split_aggr_category(text: str) -> tuple[str, str]:
    raw = clean_str(text)
    if not raw:
        return "", ""
    parts = [part.strip() for part in raw.split("/") if part.strip()]
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], parts[0]
    return parts[0], parts[1]


def build_pair_category_map(report: pd.DataFrame) -> dict[tuple[str, str], tuple[str, str]]:
    work = report.copy()
    work["pair_category_id"] = work["description_category_id"].map(clean_str)
    work["pair_type_id"] = work["type_id"].map(clean_str)
    work["local_category_1"] = work["local_category_1"].map(clean_str)
    work["local_category_2"] = work["local_category_2"].map(clean_str)
    work = work[
        (work["pair_category_id"] != "")
        & (work["pair_type_id"] != "")
        & (work["local_category_1"] != "")
        & (work["local_category_2"] != "")
    ].copy()

    result: dict[tuple[str, str], tuple[str, str]] = {}
    grouped = (
        work.groupby(["pair_category_id", "pair_type_id", "local_category_1", "local_category_2"])
        .size()
        .reset_index(name="n")
        .sort_values(["pair_category_id", "pair_type_id", "n"], ascending=[True, True, False])
    )
    for (cat_id, type_id), pair_group in grouped.groupby(["pair_category_id", "pair_type_id"], sort=False):
        top = pair_group.iloc[0]
        result[(cat_id, type_id)] = (top["local_category_1"], top["local_category_2"])
    return result


def heuristic_categories(row: pd.Series) -> tuple[str, str]:
    title = clean_str(row.get("title")).lower()
    group_name = clean_str(row.get("group_name"))
    brand = clean_str(row.get("brand"))

    rules = [
        (("огонь", "фонарь", "лампа"), ("Электрооборудование", "Навигационные огни")),
        (("переключ", "тумблер", "выключател"), ("Электрооборудование", "Переключатели")),
        (("клемм", "кабель", "провод", "розетка", "штекер"), ("Электрооборудование", "Клеммы и разъемы")),
        (("водозабор", "помп", "трюм"), ("Помпы и водоснабжение", "Фитинги и водозабор")),
        (("бак", "канистр", "топлив"), ("Топливная система", "Баки и канистры")),
        (("перчат",), ("Аксессуары", "Перчатки")),
        (("скоба", "ручка", "держател", "петля", "заглуш", "накладка", "зажим", "креплен"), ("Аксессуары", "Крепеж и фурнитура")),
        (("трап", "лестниц"), ("Палубное оборудование", "Трапы и лестницы")),
        (("якор", "шварт", "вертлюг", "карабин"), ("Якорное и швартовое", "Такелаж и швартовка")),
    ]
    for needles, pair in rules:
        if any(needle in title for needle in needles):
            return pair

    if group_name:
        return group_name, brand or group_name
    if brand:
        return "Прочее", brand
    return "Прочее", "Прочее"


def main() -> None:
    if not BACKUP_PATH.exists():
        shutil.copy2(CATALOG_PATH, BACKUP_PATH)

    catalog = pd.read_excel(CATALOG_PATH, dtype=object)
    report = pd.read_excel(SOURCE_REPORT_PATH, sheet_name="upload_candidates", dtype=object)

    report = report[
        ["sku", "local_category_1", "local_category_2", "aggr_final_category", "description_category_id", "type_id"]
    ].copy()
    pair_category_map = build_pair_category_map(report)
    report["sku_key"] = report["sku"].map(clean_str)
    report = report.drop_duplicates(subset=["sku_key"], keep="first")

    catalog["offer_id_key"] = catalog["offer_id"].map(clean_str)
    merged = catalog.merge(
        report,
        left_on="offer_id_key",
        right_on="sku_key",
        how="left",
        suffixes=("", "_src"),
    )

    restored_from_report = 0
    restored_from_pair = 0
    restored_from_aggr = 0
    still_missing = 0

    local1_values: list[str] = []
    local2_values: list[str] = []
    aggr_values: list[str] = []

    for _, row in merged.iterrows():
        local1 = clean_str(row.get("local_category_1"))
        local2 = clean_str(row.get("local_category_2"))
        aggr = clean_str(row.get("aggr_final_category"))

        if local1 and local2:
            restored_from_report += 1
        else:
            pair = (clean_str(row.get("category_id")), clean_str(row.get("type_id")))
            pair_local1, pair_local2 = pair_category_map.get(pair, ("", ""))
            used_pair = False
            if not local1 and pair_local1:
                local1 = pair_local1
                used_pair = True
            if not local2 and pair_local2:
                local2 = pair_local2
                used_pair = True
            if used_pair and local1 and local2:
                restored_from_pair += 1
            else:
                fallback1, fallback2 = split_aggr_category(aggr)
                if not local1:
                    local1 = fallback1
                if not local2:
                    local2 = fallback2 or fallback1
                if not local1 or not local2:
                    fallback1, fallback2 = heuristic_categories(row)
                    if not local1:
                        local1 = fallback1
                    if not local2:
                        local2 = fallback2
                if local1 or local2:
                    restored_from_aggr += 1
                else:
                    still_missing += 1

        local1_values.append(local1)
        local2_values.append(local2)
        aggr_values.append(aggr)

    catalog["local_category_1"] = local1_values
    catalog["local_category_2"] = local2_values
    catalog["aggr_final_category"] = aggr_values
    catalog = catalog.drop(columns=["offer_id_key"], errors="ignore")
    catalog.to_excel(CATALOG_PATH, index=False)

    summary = pd.DataFrame(
        [
            {"metric": "rows", "value": int(len(catalog))},
            {"metric": "restored_from_report", "value": int(restored_from_report)},
            {"metric": "restored_from_pair_fallback", "value": int(restored_from_pair)},
            {"metric": "restored_from_aggr_fallback", "value": int(restored_from_aggr)},
            {"metric": "still_missing_after", "value": int(still_missing)},
            {
                "metric": "local_category_1_missing_after",
                "value": int(catalog["local_category_1"].map(clean_str).eq("").sum()),
            },
            {
                "metric": "local_category_2_missing_after",
                "value": int(catalog["local_category_2"].map(clean_str).eq("").sum()),
            },
        ]
    )

    preview = catalog[
        ["offer_id", "title", "local_category_1", "local_category_2", "aggr_final_category", "category_id", "type_id"]
    ].copy()

    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        preview.to_excel(writer, sheet_name="preview", index=False)

    print(f"UPDATED: {CATALOG_PATH}")
    print(f"BACKUP : {BACKUP_PATH}")
    print(f"REPORT : {REPORT_PATH}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
