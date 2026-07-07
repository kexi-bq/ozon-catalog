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
    APP_DIR / "images",
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

    if len(text) > 2 and text[1] == ":" and text[2] == "/":
        if "static/images_refreshed/" in text:
            text = text.split("static/images_refreshed/", 1)[1]
        elif "images_refreshed/" in text:
            text = text.split("images_refreshed/", 1)[1]
        elif "/images/" in text:
            text = "images/" + text.split("/images/", 1)[1].lstrip("/\\")
        elif "catalog/import/" in text:
            text = text.split("catalog/import/", 1)[1].lstrip("/\\")
        else:
            text = text.split(":", 1)[1].lstrip("/\\")

    if "static/images_refreshed/" in text:
        return text.split("static/images_refreshed/", 1)[1].lstrip("/\\")
    if "images_refreshed/" in text:
        return text.split("images_refreshed/", 1)[1].lstrip("/\\")
    if "/images/" in text and not text.startswith("images/"):
        return "images/" + text.split("/images/", 1)[1].lstrip("/\\")
    return text.lstrip("./")


def _image_candidates_for_root(root: Path, suffix: str) -> list[Path]:
    candidates: list[Path] = []
    suffix_path = Path(suffix)
    candidates.append(root / suffix_path)

    for depth in range(1, len(root.parts) + 1):
        prefix = Path(*root.parts[-depth:])
        if suffix_path.parts[: len(prefix.parts)] == prefix.parts:
            try:
                rel = suffix_path.relative_to(prefix)
            except Exception:
                continue
            candidates.append(root / rel)
    return candidates


def _find_image_by_filename(text: str) -> Path | None:
    filename = Path(text).name
    if not filename:
        return None
    for root in IMAGE_DIRS:
        for candidate in root.rglob(filename):
            if candidate.exists():
                return candidate
    return None


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
        for path in _image_candidates_for_root(root, text):
            if path.exists():
                return path

    if len(Path(text).parts) >= 2:
        folder_name = Path(text).parts[-2]
        filename = Path(text).name
        for root in IMAGE_DIRS:
            candidate = root / folder_name / filename
            if candidate.exists():
                return candidate

    filename_match = _find_image_by_filename(text)
    if filename_match:
        return filename_match

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
