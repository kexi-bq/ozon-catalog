from __future__ import annotations

import shutil
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
ROOT = APP_DIR.parent
SOURCE_PATH = ROOT / "Новая папка" / "products_with_ozon_price_photo_recovered_vm.xlsx"
MASTER_TNVED_PATH = ROOT / "config" / "master_tnved_mapping.xlsx"
OUTPUT_PATH = APP_DIR / "data" / "catalog_2000.xlsx"
REPORT_PATH = APP_DIR / "data" / "catalog_2000_build_report.xlsx"
OLD_DATA_PATH = APP_DIR / "data" / "catalog_500.xlsx"
OLD_BACKUP_PATH = APP_DIR / "data" / "catalog_500_backup_before_2000.xlsx"
TARGET_ROWS = 2000

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

GENERIC_FALLBACK_TNVED_BY_PAIR = {
    ("17029003", "970861825"): "7326909807",
}


def clean_str(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
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
    text = clean_str(value).replace(".0", "").replace(" ", "")
    if not text:
        return ""
    try:
        return str(int(float(text)))
    except Exception:
        return "".join(ch for ch in text if ch.isdigit())


def clean_float(value: Any) -> float | None:
    text = clean_str(value)
    if not text:
        return None
    text = (
        text.replace("\u00a0", "")
        .replace("₽", "")
        .replace("руб.", "")
        .replace("р.", "")
        .replace(" ", "")
        .replace(",", ".")
    )
    try:
        return float(text)
    except Exception:
        return None


def choose_image(row: pd.Series) -> tuple[str, str]:
    candidates = [
        ("ready_recovered_image_path", clean_str(row.get("ready_recovered_image_path"))),
        ("supplier_recomputed_image", clean_str(row.get("supplier_recomputed_image"))),
        ("final_image", clean_str(row.get("final_image"))),
        ("ozon_primary_image", clean_str(row.get("ozon_primary_image"))),
    ]
    for source, raw in candidates:
        if not raw:
            continue
        if raw.startswith(("http://", "https://")):
            return raw, source
        path = Path(raw)
        if path.is_absolute() and path.exists():
            return str(path), source
        rel_root = ROOT / raw
        if rel_root.exists():
            return str(rel_root), source
        rel_app = APP_DIR / raw
        if rel_app.exists():
            return str(rel_app), source
    return "", ""


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


def infer_manual_tnved(row: pd.Series) -> str:
    code = clean_str(row.get("Код"))
    if code in MANUAL_TNVED_BY_CODE:
        return MANUAL_TNVED_BY_CODE[code]

    title = clean_str(row.get("Наименование")).lower()
    full_name = clean_str(row.get("Полное наименование")).lower()
    brand = clean_brand(row.get("supplier_recomputed_brand")).lower()
    group_name = clean_str(row.get("Группа")).lower()
    url = clean_str(row.get("supplier_recomputed_external_url")).lower()
    haystack = " ".join(part for part in [title, full_name, brand, group_name, url] if part)

    keyword_rules = [
        (("переключ", "выключател", "тумблер", "кнопк"), "8536508008"),
        (("панель", "приборн"), "8537109800"),
        (("розетк", "usb", "разъем", "штекер", "клемм"), "8536901000"),
        (("рында", "горн", "сирен"), "8512309009"),
        (("огонь", "фонарь", "светильник"), "9405420039"),
        (("датчик температур",), "9025192000"),
        (("датчик уровня", "уровня топлива", "датчик топлива", "датчик давления"), "9026102909"),
        (("заглуш", "колпач", "крышк", "пробк"), "3926909709"),
        (("бак топлив", "канистр", "емкость"), "3926909709"),
        (("перчат",), "6116930000"),
        (("фильтр топлив", "фильтр маслян", "фильтр масля", "фильтр"), "8421230000"),
        (("крыльчат", "импеллер"), "8413910008"),
        (("винт гребн", "гребной винт"), "8487100000"),
        (("наконечник", "платформ", "поруч", "трап", "багор", "коуш", "вертлюг", "карабин", "штыр", "скоба"), "7326909807"),
        (("петл", "защелк", "замок", "держател", "кронштейн", "водозаборник"), "8302490009"),
        (("герморюкзак", "сумк", "чехол"), "4202921900"),
    ]
    for needles, code_value in keyword_rules:
        if any(needle in haystack for needle in needles):
            return code_value
    return ""


def generate_description(row: pd.Series) -> str:
    title = clean_str(row.get("Наименование")) or clean_str(row.get("Полное наименование"))
    brand = clean_brand(row.get("supplier_recomputed_brand")) or clean_brand(row.get("final_brand"))
    group_name = clean_str(row.get("Группа"))
    weight = clean_str(row.get("supplier_recomputed_weight"))
    dims = [
        clean_str(row.get("supplier_recomputed_length")),
        clean_str(row.get("supplier_recomputed_width")),
        clean_str(row.get("supplier_recomputed_height")),
    ]
    dims_text = " x ".join([d for d in dims if d])

    parts = [title.rstrip(".")]
    if brand:
        parts.append(f"Бренд: {brand}.")
    if group_name:
        parts.append(f"Группа: {group_name}.")
    if dims_text:
        parts.append(f"Габариты: {dims_text}.")
    if weight:
        parts.append(f"Вес: {weight} кг.")
    parts.append("Карточка собрана из локальной базы 1С и supplier-enrichment данных.")
    return " ".join(part for part in parts if part).strip()


def select_offer_id(row: pd.Series) -> str:
    for candidate in (
        row.get("final_sku_for_upload"),
        row.get("ready_sku"),
        row.get("Артикул"),
        row.get("Код"),
    ):
        text = clean_str(candidate)
        if text:
            return text
    return clean_str(row.get("Код"))


def make_offer_ids_unique(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    counts = Counter(df["offer_id"].map(clean_str).tolist())
    seen: Counter[str] = Counter()
    changed = 0
    values: list[str] = []
    for _, row in df.iterrows():
        offer_id = clean_str(row.get("offer_id"))
        code = clean_str(row.get("code"))
        if counts[offer_id] <= 1:
            values.append(offer_id)
            continue
        seen[offer_id] += 1
        new_offer = f"{offer_id}-{code or seen[offer_id]}"
        values.append(new_offer)
        if new_offer != offer_id:
            changed += 1
    df["offer_id"] = values
    return df, changed


def main() -> None:
    if OLD_DATA_PATH.exists() and not OLD_BACKUP_PATH.exists():
        shutil.copy2(OLD_DATA_PATH, OLD_BACKUP_PATH)

    source = pd.read_excel(SOURCE_PATH, dtype=object)
    pair_tnved_map = build_pair_tnved_map(source)
    master_tnved_map = build_master_tnved_map()

    scope = source[
        (source["supplier_only_recomputed_ready"].astype(str).str.upper() == "YES")
        & (source["metrazh_flag"].astype(str).str.contains("не метраж", case=False, na=False))
        & (~source["ozon_match_found"].fillna(False).astype(bool))
    ].copy()

    records: list[dict[str, Any]] = []
    image_source_counter: Counter[str] = Counter()
    tnved_source_counter: Counter[str] = Counter()

    for _, row in scope.iterrows():
        image_value, image_source = choose_image(row)
        if not image_value:
            continue

        title = clean_str(row.get("Наименование"))
        full_name = clean_str(row.get("Полное наименование")) or title
        brand = clean_brand(row.get("supplier_recomputed_brand")) or clean_brand(row.get("final_brand"))
        category_id = clean_str(row.get("supplier_recomputed_category_id"))
        type_id = clean_str(row.get("supplier_recomputed_type_id"))
        pair = (category_id, type_id)

        tnved = normalize_tnved(row.get("final_tnved"))
        tnved_source = "final_tnved"
        if not tnved:
            tnved = pair_tnved_map.get(pair, "")
            tnved_source = "pair_map" if tnved else ""
        if not tnved:
            tnved = master_tnved_map.get(pair, "")
            tnved_source = "master_map" if tnved else ""
        if not tnved:
            tnved = infer_manual_tnved(row)
            tnved_source = "manual_rule" if tnved else ""
        if not tnved:
            tnved = GENERIC_FALLBACK_TNVED_BY_PAIR.get(pair, "")
            tnved_source = "generic_pair_fallback" if tnved else ""

        description = clean_str(row.get("supplier_recomputed_description")) or generate_description(row)
        description_source = "supplier_recomputed_description" if clean_str(row.get("supplier_recomputed_description")) else "generated"

        retail_price = clean_float(row.get("Цена (Розница)")) or 0.0
        ozon_price = clean_float(row.get("ozon_price_for_retail")) or retail_price
        stock_qty = int(clean_float(row.get("Количество")) or 0)
        if stock_qty <= 0:
            stock_qty = 1
        weight = clean_float(row.get("supplier_recomputed_weight")) or 0.0
        length = clean_float(row.get("supplier_recomputed_length")) or 0.0
        width = clean_float(row.get("supplier_recomputed_width")) or 0.0
        height = clean_float(row.get("supplier_recomputed_height")) or 0.0

        quality_score = 0
        quality_score += 30 if image_source == "ready_recovered_image_path" else 20
        quality_score += 20 if tnved_source == "final_tnved" else 15 if tnved_source in {"pair_map", "master_map"} else 8 if tnved_source else 0
        quality_score += 10 if description_source == "supplier_recomputed_description" else 5
        quality_score += 10 if brand else 0
        quality_score += 5 if stock_qty > 0 else 0
        quality_score += 5 if image_value.startswith("http") else 8

        records.append(
            {
                "offer_id": select_offer_id(row),
                "code": clean_str(row.get("Код")),
                "title": title,
                "full_name": full_name,
                "brand": brand,
                "group_name": clean_str(row.get("Группа")),
                "category_path": f"{category_id} / {type_id}",
                "category_id": category_id,
                "type_id": type_id,
                "retail_price": retail_price,
                "ozon_price": ozon_price,
                "stock_qty": stock_qty,
                "stock_text": "В наличии" if stock_qty > 0 else "Нет в наличии",
                "description": description,
                "primary_image": image_value,
                "weight": weight,
                "length": length,
                "width": width,
                "height": height,
                "external_url": clean_str(row.get("supplier_recomputed_external_url")) or clean_str(row.get("final_supplier_url")),
                "tnved": tnved,
                "metrazh_flag": clean_str(row.get("metrazh_flag")),
                "image_source": image_source,
                "tnved_source": tnved_source,
                "description_source": description_source,
                "quality_score": quality_score,
            }
        )
        image_source_counter[image_source] += 1
        tnved_source_counter[tnved_source or "missing"] += 1

    catalog = pd.DataFrame(records)
    catalog = catalog[catalog["tnved"].map(clean_str).ne("")].copy()
    catalog = catalog.sort_values(
        ["quality_score", "brand", "title", "offer_id"],
        ascending=[False, True, True, True],
        kind="stable",
    ).head(TARGET_ROWS).reset_index(drop=True)
    catalog, offer_id_fixed = make_offer_ids_unique(catalog)

    final_columns = [
        "offer_id",
        "code",
        "title",
        "full_name",
        "brand",
        "group_name",
        "category_path",
        "category_id",
        "type_id",
        "retail_price",
        "ozon_price",
        "stock_qty",
        "stock_text",
        "description",
        "primary_image",
        "weight",
        "length",
        "width",
        "height",
        "external_url",
        "tnved",
        "metrazh_flag",
    ]
    output = catalog[final_columns].copy()
    output.to_excel(OUTPUT_PATH, index=False)

    summary = pd.DataFrame(
        [
            {"metric": "source_scope_ready_non_metrazh_no_ozon", "value": int(len(scope))},
            {"metric": "records_with_real_image", "value": int(len(records))},
            {"metric": "records_with_tnved_after_fill", "value": int(len(catalog))},
            {"metric": "output_rows", "value": int(len(output))},
            {"metric": "offer_id_fixed", "value": int(offer_id_fixed)},
            {"metric": "missing_brand_after", "value": int(output["brand"].map(clean_brand).eq("").sum())},
            {"metric": "missing_description_after", "value": int(output["description"].map(clean_str).eq("").sum())},
            {"metric": "missing_tnved_after", "value": int(output["tnved"].map(normalize_tnved).eq("").sum())},
            {"metric": "missing_image_after", "value": int(output["primary_image"].map(clean_str).eq("").sum())},
        ]
    )
    image_sources_df = pd.DataFrame(
        [{"image_source": key, "count": int(value)} for key, value in image_source_counter.most_common()]
    )
    tnved_sources_df = pd.DataFrame(
        [{"tnved_source": key, "count": int(value)} for key, value in tnved_source_counter.most_common()]
    )
    preview_df = catalog[
        [
            "offer_id",
            "code",
            "title",
            "brand",
            "group_name",
            "primary_image",
            "image_source",
            "tnved",
            "tnved_source",
            "quality_score",
        ]
    ].copy()

    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        image_sources_df.to_excel(writer, sheet_name="image_sources", index=False)
        tnved_sources_df.to_excel(writer, sheet_name="tnved_sources", index=False)
        preview_df.to_excel(writer, sheet_name="selected_2000", index=False)

    print(f"OUTPUT: {OUTPUT_PATH}")
    print(f"REPORT: {REPORT_PATH}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
