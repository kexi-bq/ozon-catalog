from __future__ import annotations

import shutil
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
ROOT = APP_DIR.parent
CATALOG_PATH = APP_DIR / "data" / "catalog_500.xlsx"
BACKUP_PATH = APP_DIR / "data" / "catalog_500_before_ozon_fix.xlsx"
REPORT_PATH = APP_DIR / "data" / "catalog_500_ozon_fix_report.xlsx"
SOURCE_PATH = ROOT / "Новая папка" / "products_with_ozon_price.xlsx"
MASTER_TNVED_PATH = ROOT / "config" / "master_tnved_mapping.xlsx"

MANUAL_TNVED_BY_CODE = {
    "049-0001928": "8302419000",
    "-25": "7326909807",
    "049-0000349": "8536508008",
    "049-0003742": "3926909709",
    "049-0003750": "3926909709",
    "049-0004523": "9405420039",
    "10262341": "8512309009",
    "049-0005110": "9025192000",
    "049-0004039": "9026102909",
    "049-0004012": "9026102909",
    "049-0005117": "8302490009",
    "049-0005270": "8536901000",
    "049-0005268": "8536901000",
    "049-0000295": "8536901000",
    "049-0002259": "8536901000",
    "049-0005203": "8536901000",
    "049-0005161": "7326909807",
    "049-0000351": "8302490009",
    "10267525": "6116930000",
}


def clean_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def clean_brand(value: Any) -> str:
    text = clean_str(value).strip('"').strip("'").strip()
    if not text:
        return ""
    parts = [part.strip() for part in text.split(",") if part.strip()]
    if parts:
        text = parts[0]
    upper = text.upper()
    if upper in {"NAN", "NONE"}:
        return ""
    return text


def normalize_tnved(value: Any) -> str:
    text = clean_str(value)
    if not text:
        return ""
    text = text.replace(".0", "").replace(" ", "")
    try:
        return str(int(float(text)))
    except Exception:
        return "".join(ch for ch in text if ch.isdigit())


def build_pair_tnved_map(source_df: pd.DataFrame) -> dict[tuple[str, str], str]:
    work = source_df.copy()
    work["pair_category_id"] = work["supplier_recomputed_category_id"].map(clean_str)
    work["pair_type_id"] = work["supplier_recomputed_type_id"].map(clean_str)
    work["pair_tnved"] = work["final_tnved"].map(normalize_tnved)
    work = work[(work["pair_category_id"] != "") & (work["pair_type_id"] != "") & (work["pair_tnved"] != "")]
    result: dict[tuple[str, str], str] = {}
    for (cat_id, type_id), group in work.groupby(["pair_category_id", "pair_type_id"]):
        counts = Counter(group["pair_tnved"].tolist())
        result[(cat_id, type_id)] = counts.most_common(1)[0][0]
    return result


def build_master_tnved_map() -> dict[tuple[str, str], str]:
    if not MASTER_TNVED_PATH.exists():
        return {}
    df = pd.read_excel(MASTER_TNVED_PATH, dtype=object)
    df["cat"] = df["ozon_description_category_id"].map(clean_str)
    df["typ"] = df["ozon_type_id"].map(clean_str)
    df["tnved"] = df["selected_tnved"].map(normalize_tnved)
    df = df[(df["cat"] != "") & (df["typ"] != "") & (df["tnved"] != "")]
    return {(row["cat"], row["typ"]): row["tnved"] for _, row in df.iterrows()}


def build_source_lookup(source_df: pd.DataFrame) -> dict[str, pd.Series]:
    lookup: dict[str, pd.Series] = {}
    for _, row in source_df.iterrows():
        code = clean_str(row.get("Код"))
        if code and code not in lookup:
            lookup[code] = row
    return lookup


def generate_description(row: pd.Series) -> str:
    title = clean_str(row.get("title"))
    brand = clean_brand(row.get("brand"))
    group_name = clean_str(row.get("group_name"))
    dims = [clean_str(row.get("length")), clean_str(row.get("width")), clean_str(row.get("height"))]
    dims_text = " x ".join([d for d in dims if d])
    weight = clean_str(row.get("weight"))

    parts = [title.rstrip(".")]
    if brand:
        parts.append(f"Бренд: {brand}.")
    if group_name:
        parts.append(f"Группа: {group_name}.")
    if dims_text:
        parts.append(f"Габариты: {dims_text}.")
    if weight:
        parts.append(f"Вес: {weight} кг.")
    parts.append("Товар из подготовленной локальной базы для загрузки на Ozon.")
    return " ".join(part for part in parts if part).strip()


def infer_manual_tnved(row: pd.Series) -> str:
    code = clean_str(row.get("code"))
    if code in MANUAL_TNVED_BY_CODE:
        return MANUAL_TNVED_BY_CODE[code]

    title = clean_str(row.get("title")).lower()
    url = clean_str(row.get("external_url")).lower()
    haystack = f"{title} {url}"

    if "переключател" in haystack:
        return "8536508008"
    if "заглуш" in haystack:
        return "3926909709"
    if "огонь" in haystack or "light" in haystack:
        return "9405420039"
    if "рында" in haystack:
        return "8512309009"
    if "датчик температур" in haystack:
        return "9025192000"
    if "датчик уровня" in haystack:
        return "9026102909"
    if "клемма" in haystack:
        return "8536901000"
    if "защелк" in haystack or "петл" in haystack:
        return "8302490009"
    if "наконечник" in haystack or "платформ" in haystack:
        return "7326909807"
    if "перчат" in haystack:
        return "6116930000"
    if "водозаборник" in haystack:
        return "8302419000"
    return ""


def make_offer_ids_unique(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    counts = Counter(df["offer_id"].map(clean_str).tolist())
    changed = 0
    seen: Counter[str] = Counter()
    new_values: list[str] = []
    for _, row in df.iterrows():
        offer_id = clean_str(row.get("offer_id"))
        code = clean_str(row.get("code"))
        if counts[offer_id] <= 1:
            new_values.append(offer_id)
            continue
        seen[offer_id] += 1
        suffix = code or str(seen[offer_id])
        new_offer_id = f"{offer_id}-{suffix}"
        new_values.append(new_offer_id)
        if new_offer_id != offer_id:
            changed += 1
    df["offer_id"] = new_values
    return df, changed


def main() -> None:
    if not BACKUP_PATH.exists():
        shutil.copy2(CATALOG_PATH, BACKUP_PATH)

    catalog = pd.read_excel(CATALOG_PATH, dtype=object)
    source = pd.read_excel(SOURCE_PATH, dtype=object)

    source_lookup = build_source_lookup(source)
    pair_tnved_map = build_pair_tnved_map(source)
    master_tnved_map = build_master_tnved_map()

    brand_fixed = 0
    desc_filled = 0
    tnved_filled = 0
    tnved_manual_filled = 0

    for idx, row in catalog.iterrows():
        code = clean_str(row.get("code"))
        source_row = source_lookup.get(code)

        current_brand = clean_brand(row.get("brand"))
        source_brand = ""
        if source_row is not None:
            source_brand = clean_brand(source_row.get("final_brand")) or clean_brand(source_row.get("supplier_recomputed_brand"))
        better_brand = source_brand or current_brand
        if better_brand and better_brand != clean_str(row.get("brand")):
            catalog.at[idx, "brand"] = better_brand
            brand_fixed += 1

        current_desc = clean_str(row.get("description"))
        if not current_desc:
            source_desc = clean_str(source_row.get("supplier_recomputed_description")) if source_row is not None else ""
            new_desc = source_desc or generate_description(catalog.loc[idx])
            if new_desc:
                catalog.at[idx, "description"] = new_desc
                desc_filled += 1

        current_tnved = normalize_tnved(row.get("tnved"))
        if not current_tnved:
            source_tnved = normalize_tnved(source_row.get("final_tnved")) if source_row is not None else ""
            pair = (clean_str(row.get("category_id")), clean_str(row.get("type_id")))
            mapped_tnved = source_tnved or pair_tnved_map.get(pair, "") or master_tnved_map.get(pair, "")
            if not mapped_tnved:
                mapped_tnved = infer_manual_tnved(catalog.loc[idx])
                if mapped_tnved:
                    tnved_manual_filled += 1
            if mapped_tnved:
                catalog.at[idx, "tnved"] = mapped_tnved
                tnved_filled += 1

    catalog, offer_id_fixed = make_offer_ids_unique(catalog)

    stock_num = pd.to_numeric(catalog["stock_qty"], errors="coerce").fillna(0)
    stock_zero_count = int(stock_num.le(0).sum())
    if stock_zero_count:
        catalog.loc[stock_num.le(0), "stock_qty"] = 1
        catalog.loc[stock_num.le(0), "stock_text"] = "В наличии"

    issue_rows: list[dict[str, Any]] = []
    for _, row in catalog.iterrows():
        issues: list[str] = []
        if clean_brand(row.get("brand")) == "":
            issues.append("missing_brand")
        if clean_str(row.get("description")) == "":
            issues.append("missing_description")
        if normalize_tnved(row.get("tnved")) == "":
            issues.append("missing_tnved")
        if clean_str(row.get("offer_id")) == "":
            issues.append("missing_offer_id")
        if clean_str(row.get("primary_image")) == "":
            issues.append("missing_primary_image")
        stock_value = pd.to_numeric(pd.Series([row.get("stock_qty")]), errors="coerce").fillna(0).iloc[0]
        if stock_value <= 0:
            issues.append("bad_stock_qty")
        for dim_col in ("weight", "length", "width", "height", "retail_price", "ozon_price"):
            value = pd.to_numeric(pd.Series([row.get(dim_col)]), errors="coerce").iloc[0]
            if pd.isna(value) or value <= 0:
                issues.append(f"bad_{dim_col}")
        if issues:
            issue_rows.append(
                {
                    "offer_id": clean_str(row.get("offer_id")),
                    "code": clean_str(row.get("code")),
                    "title": clean_str(row.get("title")),
                    "issues": ", ".join(issues),
                }
            )

    catalog.to_excel(CATALOG_PATH, index=False)

    summary = pd.DataFrame(
        [
            {"metric": "rows", "value": int(len(catalog))},
            {"metric": "brand_fixed", "value": int(brand_fixed)},
            {"metric": "description_filled", "value": int(desc_filled)},
            {"metric": "tnved_filled", "value": int(tnved_filled)},
            {"metric": "tnved_manual_filled", "value": int(tnved_manual_filled)},
            {"metric": "offer_id_fixed", "value": int(offer_id_fixed)},
            {"metric": "stock_zero_fixed", "value": int(stock_zero_count)},
            {"metric": "tnved_missing_after", "value": int(catalog["tnved"].map(normalize_tnved).eq("").sum())},
            {"metric": "description_missing_after", "value": int(catalog["description"].map(clean_str).eq("").sum())},
            {"metric": "duplicate_offer_id_after", "value": int(catalog["offer_id"].map(clean_str).duplicated().sum())},
            {"metric": "issue_rows_after", "value": int(len(issue_rows))},
        ]
    )

    missing_after = catalog[
        catalog["tnved"].map(normalize_tnved).eq("") | catalog["description"].map(clean_str).eq("")
    ].copy()

    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        missing_after.to_excel(writer, sheet_name="remaining_missing", index=False)
        pd.DataFrame(issue_rows).to_excel(writer, sheet_name="issue_rows", index=False)

    print(f"UPDATED: {CATALOG_PATH}")
    print(f"BACKUP : {BACKUP_PATH}")
    print(f"REPORT : {REPORT_PATH}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
