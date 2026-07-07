from __future__ import annotations

from pathlib import Path

import pandas as pd

APP_DIR = Path(__file__).resolve().parent
DATA_PATH = APP_DIR / "data" / "catalog_500_FINAL_FOR_SITE.xlsx"
OUTPUT_PATH = APP_DIR / "data" / "image_paths_audit.xlsx"
IMAGE_DIRS = [
    APP_DIR / "static" / "images_refreshed",
    APP_DIR / "images_refreshed",
    APP_DIR / "data" / "images_refreshed",
]


def clean_str(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def normalize_image_source(value: str) -> str:
    text = clean_str(value).replace("\\", "/").strip()
    if not text:
        return ""
    if text.startswith(("http://", "https://")):
        return text
    if "static/images_refreshed/" in text:
        return text.split("static/images_refreshed/", 1)[1].lstrip("/\\")
    if "images_refreshed/" in text:
        return text.split("images_refreshed/", 1)[1].lstrip("/\\")
    return text.lstrip("./")


def resolve_path(value: str) -> Path | None:
    text = normalize_image_source(value)
    if not text:
        return None
    if text.startswith(("http://", "https://")):
        return None
    candidate = Path(text)
    if candidate.is_absolute() and candidate.exists():
        return candidate
    for root in IMAGE_DIRS:
        path = root / text
        if path.exists():
            return path
    for root in IMAGE_DIRS:
        path = root / Path(text).name
        if path.exists():
            return path
    return None


def main() -> None:
    df = pd.read_excel(DATA_PATH, dtype=object)
    df["primary_image"] = df["primary_image"].fillna("")
    df["normalized_image"] = df["primary_image"].astype(str).map(normalize_image_source)
    df["image_exists"] = df["normalized_image"].map(lambda x: bool(resolve_path(x)))
    df["resolved_path"] = df["normalized_image"].map(lambda x: str(resolve_path(x)) if resolve_path(x) else "")

    summary = {
        "rows": len(df),
        "filled_primary_image": int(df["normalized_image"].astype(bool).sum()),
        "files_found": int(df["image_exists"].sum()),
        "files_missing": int((~df["image_exists"]).sum()),
    }

    missing = df[~df["image_exists"]][["offer_id", "code", "primary_image", "normalized_image"]].head(20)

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="summary", index=False)
        missing.to_excel(writer, sheet_name="missing_images", index=False)

    print("Image audit completed")
    print(summary)
    print(f"Report saved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
