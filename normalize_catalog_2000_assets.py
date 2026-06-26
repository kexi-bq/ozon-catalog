from __future__ import annotations

import shutil
from pathlib import Path

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
CATALOG_PATH = APP_DIR / "data" / "catalog_2000.xlsx"
BACKUP_PATH = APP_DIR / "data" / "catalog_2000_before_asset_normalize.xlsx"
REPORT_PATH = APP_DIR / "data" / "catalog_2000_asset_normalize_report.xlsx"
IMAGES_DIR = APP_DIR / "images"


def is_abs_windows_path(text: str) -> bool:
    return len(text) > 2 and text[1] == ":" and text[2] in ("\\", "/")


def safe_path_part(text: str) -> str:
    cleaned = "".join("_" if ch in '<>:"/\\|?*' else ch for ch in text.strip())
    cleaned = " ".join(cleaned.split())
    return cleaned or "unknown"


def main() -> None:
    if not BACKUP_PATH.exists():
        shutil.copy2(CATALOG_PATH, BACKUP_PATH)

    df = pd.read_excel(CATALOG_PATH, dtype=object)

    copied = 0
    already_relative = 0
    remote_kept = 0
    missing_source = 0
    rewritten = 0

    new_values: list[str] = []
    details: list[dict[str, str]] = []

    for _, row in df.iterrows():
        offer_id = str(row.get("offer_id", "")).strip() or str(row.get("code", "")).strip() or "unknown"
        safe_offer_id = safe_path_part(offer_id)
        value = "" if pd.isna(row.get("primary_image")) else str(row.get("primary_image")).strip()

        if not value:
            new_values.append("")
            details.append({"offer_id": offer_id, "old_value": value, "new_value": "", "status": "empty"})
            continue

        if value.startswith(("http://", "https://")):
            remote_kept += 1
            new_values.append(value)
            details.append({"offer_id": offer_id, "old_value": value, "new_value": value, "status": "remote_kept"})
            continue

        if not is_abs_windows_path(value):
            already_relative += 1
            new_values.append(value.replace("\\", "/"))
            details.append({"offer_id": offer_id, "old_value": value, "new_value": value.replace("\\", "/"), "status": "relative_kept"})
            continue

        source_path = Path(value)
        if not source_path.exists():
            missing_source += 1
            new_values.append(value)
            details.append({"offer_id": offer_id, "old_value": value, "new_value": value, "status": "missing_source"})
            continue

        target_dir = IMAGES_DIR / safe_offer_id
        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = target_dir / source_path.name
        if not target_path.exists():
            shutil.copy2(source_path, target_path)
            copied += 1

        rel_path = target_path.relative_to(APP_DIR).as_posix()
        rewritten += 1
        new_values.append(rel_path)
        details.append({"offer_id": offer_id, "old_value": value, "new_value": rel_path, "status": "copied_to_repo"})

    df["primary_image"] = new_values
    df.to_excel(CATALOG_PATH, index=False)

    summary = pd.DataFrame(
        [
            {"metric": "rows", "value": int(len(df))},
            {"metric": "rewritten_to_relative", "value": int(rewritten)},
            {"metric": "files_copied", "value": int(copied)},
            {"metric": "remote_kept", "value": int(remote_kept)},
            {"metric": "already_relative", "value": int(already_relative)},
            {"metric": "missing_source", "value": int(missing_source)},
        ]
    )

    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        pd.DataFrame(details).to_excel(writer, sheet_name="details", index=False)

    print(f"UPDATED: {CATALOG_PATH}")
    print(f"BACKUP : {BACKUP_PATH}")
    print(f"REPORT : {REPORT_PATH}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
