from __future__ import annotations

import ast
import base64
import html
import math
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


APP_DIR = Path(__file__).resolve().parent
ROOT = APP_DIR.parent
DATA_PATH = APP_DIR / "data" / "catalog_500_exact_match.xlsx"
PAGE_SIZE = 25


def clean_str(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def clean_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
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


def first_image(value: Any) -> str:
    text = clean_str(value)
    if not text:
        return ""
    if text.startswith("[") and text.endswith("]"):
        try:
            parsed = ast.literal_eval(text)
            if isinstance(parsed, list) and parsed:
                return clean_str(parsed[0])
        except Exception:
            pass
    return text


def resolve_image(value: Any) -> str:
    text = first_image(value)
    if not text:
        return ""
    if text.startswith(("http://", "https://")):
        return text
    path = APP_DIR / text
    if path.exists():
        return str(path)
    root_path = ROOT / text
    if root_path.exists():
        return str(root_path)
    raw_path = Path(text)
    if raw_path.is_absolute() and raw_path.exists():
        return str(raw_path)
    return ""


@st.cache_data(show_spinner=False)
def image_to_src(path_or_url: str) -> str:
    text = clean_str(path_or_url)
    if not text:
        return ""
    if text.startswith(("http://", "https://")):
        return text
    path = Path(text)
    if not path.exists():
        return ""
    suffix = path.suffix.lower()
    mime = "image/jpeg"
    if suffix == ".png":
        mime = "image/png"
    elif suffix == ".webp":
        mime = "image/webp"
    data = base64.b64encode(path.read_bytes()).decode("utf-8")
    return f"data:{mime};base64,{data}"


def get_col(df: pd.DataFrame, name: str, default: Any = "") -> pd.Series:
    if name in df.columns:
        return df[name]
    return pd.Series([default] * len(df), index=df.index)


@st.cache_data(show_spinner=True)
def load_catalog() -> pd.DataFrame:
    df = pd.read_excel(DATA_PATH, dtype=object)

    df["offer_id"] = get_col(df, "offer_id", "").map(clean_str)
    if df["offer_id"].eq("").all() and "sku" in df.columns:
        df["offer_id"] = df["sku"].map(clean_str)

    df["code"] = get_col(df, "code", "").map(clean_str)
    if df["code"].eq("").all() and "code_1c" in df.columns:
        df["code"] = df["code_1c"].map(clean_str)

    df["title"] = get_col(df, "title", "").map(clean_str)
    if df["title"].eq("").all() and "name" in df.columns:
        df["title"] = df["name"].map(clean_str)

    df["full_name"] = get_col(df, "full_name", "").map(clean_str)
    if df["full_name"].eq("").all() and "name" in df.columns:
        df["full_name"] = df["name"].map(clean_str)

    df["brand"] = get_col(df, "brand", "").map(clean_str)
    df["brand_clean"] = df["brand"].map(clean_str)

    df["category_nav"] = get_col(df, "local_category_1", "").map(clean_str)
    df["subcategory_nav"] = get_col(df, "local_category_2", "").map(clean_str)

    df["group_name"] = get_col(df, "group_name", "").map(clean_str)
    if df["group_name"].eq("").all():
        for group_col in ["Группа", "group", "ready_source_group"]:
            if group_col in df.columns:
                df["group_name"] = df[group_col].map(clean_str)
                break

    df["material"] = get_col(df, "material", "").map(clean_str)
    if df["material"].eq("").all():
        for material_col in ["Материал", "supplier_recomputed_material"]:
            if material_col in df.columns:
                df["material"] = df[material_col].map(clean_str)
                break

    df["color"] = get_col(df, "color", "").map(clean_str)
    if df["color"].eq("").all():
        for color_col in ["Цвет", "supplier_recomputed_color"]:
            if color_col in df.columns:
                df["color"] = df[color_col].map(clean_str)
                break

    if "ozon_category_name" in df.columns:
        df["ozon_category_display"] = df["ozon_category_name"].map(clean_str)
    elif "aggr_final_category" in df.columns:
        df["ozon_category_display"] = df["aggr_final_category"].map(clean_str)
    elif "mapped_ozon_category_name" in df.columns:
        df["ozon_category_display"] = df["mapped_ozon_category_name"].map(clean_str)
    else:
        df["ozon_category_display"] = ""

    df["retail_price_num"] = get_col(df, "retail_price", 0).map(clean_float).fillna(0.0)
    if df["retail_price_num"].eq(0).all() and "price" in df.columns:
        df["retail_price_num"] = df["price"].map(clean_float).fillna(0.0)

    df["ozon_price_num"] = get_col(df, "ozon_price", 0).map(clean_float).fillna(df["retail_price_num"])
    if df["ozon_price_num"].eq(0).all() and "price" in df.columns:
        df["ozon_price_num"] = df["price"].map(clean_float).fillna(df["retail_price_num"])
    df["price_source"] = get_col(df, "price_source", "").map(clean_str)

    df["stock_qty_num"] = get_col(df, "stock_qty", 0).map(clean_float).fillna(0)
    if df["stock_qty_num"].eq(0).all() and "stock" in df.columns:
        df["stock_qty_num"] = df["stock"].map(clean_float).fillna(0)

    image_col = "primary_image" if "primary_image" in df.columns else "main_image"
    df["primary_image"] = get_col(df, image_col, "").map(clean_str)
    df["primary_image_resolved"] = df["primary_image"].map(resolve_image)
    df["has_image"] = df["primary_image_resolved"].astype(str).str.strip().ne("")

    df["stock_text"] = df["stock_qty_num"].apply(
        lambda x: f"В наличии: {int(x)}" if x > 0 else "Нет в наличии"
    )

    df["weight"] = get_col(df, "weight", "").map(clean_str)
    df["length"] = get_col(df, "length", "").map(clean_str)
    df["width"] = get_col(df, "width", "").map(clean_str)
    df["height"] = get_col(df, "height", "").map(clean_str)

    if "tnved" in df.columns:
        df["tnved"] = df["tnved"].map(clean_str)
    elif "aggr_final_tnved" in df.columns:
        df["tnved"] = df["aggr_final_tnved"].map(clean_str)
    else:
        df["tnved"] = ""

    df["description"] = get_col(df, "description", "").map(clean_str)
    df["external_url"] = get_col(df, "external_url", "").map(clean_str)
    if df["external_url"].eq("").all() and "source_url" in df.columns:
        df["external_url"] = df["source_url"].map(clean_str)

    df["search_blob"] = (
        df["offer_id"].fillna("").astype(str)
        + " "
        + df["code"].fillna("").astype(str)
        + " "
        + df["title"].fillna("").astype(str)
        + " "
        + df["full_name"].fillna("").astype(str)
        + " "
        + df["brand_clean"].fillna("").astype(str)
        + " "
        + df["category_nav"].fillna("").astype(str)
        + " "
        + df["subcategory_nav"].fillna("").astype(str)
        + " "
        + df["ozon_category_display"].fillna("").astype(str)
        + " "
        + df["group_name"].fillna("").astype(str)
    ).str.lower()

    df = df.sort_values(
        ["category_nav", "subcategory_nav", "brand", "title", "offer_id"],
        kind="stable",
    ).reset_index(drop=True)

    df["catalog_number"] = range(1, len(df) + 1)
    return df


def format_price(value: float | int | None) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    return f"{int(math.ceil(float(value))):,} ₽".replace(",", " ")


def format_dimensions(row: pd.Series) -> str:
    length = clean_str(row["length"])
    width = clean_str(row["width"])
    height = clean_str(row["height"])
    parts = [x for x in [length, width, height] if x]
    if len(parts) == 3:
        return f"{length} × {width} × {height}"
    return "—"


def format_weight(value: Any) -> str:
    weight = clean_float(value)
    if weight is None or weight <= 0:
        return "—"
    if weight >= 1000:
        kg = weight / 1000
        kg_text = f"{kg:.2f}".rstrip("0").rstrip(".")
        gram_text = f"{int(round(weight)):,}".replace(",", " ")
        return f"{kg_text} кг / {gram_text} г"
    if weight == int(weight):
        return f"{int(weight)} г"
    return f"{weight:g} г"


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: #f4f7fb; }
        .block-container { max-width: 1380px; padding-top: 1rem; padding-bottom: 2rem; }

        .topbar {
            background: linear-gradient(135deg, #005bff 0%, #2b7cff 60%, #8cb6ff 100%);
            border-radius: 24px; padding: 20px 24px; color: white; margin-bottom: 18px;
        }
        .topbar h1 { margin: 0; font-size: 34px; font-weight: 800; }
        .topbar p { margin: 8px 0 0 0; font-size: 14px; color: rgba(255,255,255,.9); }

        .metric {
            background: white; border: 1px solid #dfe7f2; border-radius: 18px; padding: 16px 18px;
        }
        .metric-label { color: #607087; font-size: 12px; margin-bottom: 6px; }
        .metric-value { color: #1f2d3d; font-size: 24px; font-weight: 800; }

        .category-title { font-size: 24px; font-weight: 800; margin-top: 18px; margin-bottom: 10px; color: #1f2d3d; }

        .product-link { text-decoration: none !important; color: inherit !important; display: block; }
        .product-card {
            background: white; border: 1px solid #dfe7f2; border-radius: 18px; padding: 12px;
            min-height: 100%; transition: transform .12s ease, box-shadow .12s ease, border-color .12s ease;
        }
        .product-card:hover {
            transform: translateY(-2px); box-shadow: 0 10px 26px rgba(15, 35, 80, .12); border-color: #8cb6ff;
        }
        .product-image {
            width: 100%; aspect-ratio: 1 / 1; object-fit: contain; background: #f8fbff;
            border-radius: 14px; display: block; margin-bottom: 10px;
        }
        .num-badge {
            display: inline-block; background: #eef4ff; color: #005bff; border: 1px solid #cfe0ff;
            border-radius: 999px; padding: 4px 9px; font-size: 12px; font-weight: 800; margin-bottom: 8px;
        }
        .brand { font-size: 11px; color: #005bff; font-weight: 700; text-transform: uppercase; }
        .title { font-size: 15px; line-height: 1.35; font-weight: 700; color: #162033; min-height: 62px; margin: 8px 0 10px 0; }
        .price { font-size: 26px; font-weight: 800; color: #f91155; margin: 8px 0 2px 0; }
        .old { font-size: 13px; color: #8a97a8; }
        .meta { font-size: 12px; color: #607087; line-height: 1.55; margin-top: 10px; }
        .full-info {
            border-top: 1px solid #edf2f8; margin-top: 12px; padding-top: 10px;
            font-size: 12px; color: #41536b; line-height: 1.55; overflow-wrap: anywhere; word-break: break-word;
        }
        .full-info-row { margin-bottom: 5px; }
        .full-info-label { color: #718198; font-weight: 700; }
        .full-description {
            margin-top: 8px; padding: 9px 10px; border-radius: 12px; background: #f8fbff;
            color: #304057; white-space: pre-wrap;
        }

        .badge {
            display: inline-block; padding: 5px 10px; border-radius: 999px; font-size: 12px; font-weight: 700;
            margin-top: 8px; background: #eaf8ef; color: #138a43;
        }
        .badge.out { background: #fff2e2; color: #a45d00; }

        .detail { background: white; border: 1px solid #dfe7f2; border-radius: 24px; padding: 24px; }
        .spec { background: #f8fbff; border: 1px solid #e4edf9; border-radius: 18px; padding: 14px 16px; margin-bottom: 10px; }
        .spec-label { font-size: 12px; color: #718198; margin-bottom: 4px; }
        .spec-value { font-size: 15px; color: #18253a; font-weight: 600; word-break: break-word; }

        div[data-testid="stButton"] button { border-radius: 14px; min-height: 42px; white-space: normal; }
        img { border-radius: 14px; }

        *, *::before, *::after { box-sizing: border-box; }
        html, body, .stApp { overflow-x: hidden; }

        @media (max-width: 768px) {
            .block-container {
                padding-left: 12px !important;
                padding-right: 12px !important;
                padding-top: .6rem;
                max-width: 100% !important;
            }
            .topbar { border-radius: 18px; padding: 16px; margin-bottom: 12px; }
            .topbar h1 { font-size: 25px; line-height: 1.15; }
            .topbar p { font-size: 12px; line-height: 1.35; }
            .metric { padding: 12px 13px; border-radius: 14px; }
            .metric-value { font-size: 20px; }
            .category-title { font-size: 20px; }
            .title { min-height: auto; font-size: 14px; }
            .price { font-size: 22px; }
            .detail { padding: 16px; border-radius: 18px; }
            .product-card {
                width: 100%;
                max-width: 100%;
                padding: 10px;
                border-radius: 14px;
                overflow: hidden;
            }
            .product-link { width: 100%; max-width: 100%; }
            .product-image { max-width: 100%; height: auto; }
            .meta, .title, .brand, .spec-value { overflow-wrap: anywhere; word-break: break-word; }
            div[data-testid="stHorizontalBlock"] {
                flex-wrap: wrap !important;
                gap: .6rem !important;
            }
            div[data-testid="column"] {
                width: auto !important;
                min-width: calc(50% - .3rem) !important;
                flex: 1 1 calc(50% - .3rem) !important;
            }
        }

        @media (max-width: 480px) {
            .block-container {
                padding-left: 10px !important;
                padding-right: 10px !important;
            }
            div[data-testid="column"] {
                width: 100% !important;
                min-width: 100% !important;
                flex: 1 1 100% !important;
            }
            .topbar h1 { font-size: 23px; }
            .price { font-size: 21px; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_metrics(df: pd.DataFrame) -> None:
    c1, c2, c3, c4 = st.columns(4)
    values = [
        ("Товаров", f"{len(df):,}".replace(",", " ")),
        ("С фото", f"{int(df['has_image'].sum()):,}".replace(",", " ")),
        ("Брендов", f"{int(df['brand_clean'].nunique()):,}".replace(",", " ")),
        ("Средняя цена", format_price(df["ozon_price_num"].mean() if not df.empty else 0)),
    ]
    for col, (label, value) in zip((c1, c2, c3, c4), values):
        with col:
            st.markdown(f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div></div>', unsafe_allow_html=True)


def render_category_shortcuts(df: pd.DataFrame) -> None:
    category_counts = (
        df[df["category_nav"].astype(str).str.strip().ne("")]
        .groupby("category_nav").size().sort_values(ascending=False).head(12)
    )
    if category_counts.empty:
        return

    st.markdown('<div class="category-title">Категории</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    for idx, (category_name, count) in enumerate(category_counts.items()):
        with cols[idx % 4]:
            if st.button(f"{category_name} ({count})", key=f"cat_shortcut_{idx}", use_container_width=True):
                st.query_params["category"] = category_name
                st.query_params.pop("subcategory", None)
                st.query_params["page"] = "1"
                st.rerun()


def render_subcategory_shortcuts(df: pd.DataFrame, selected_category: str) -> None:
    if selected_category == "Все категории":
        return

    pool = df[df["category_nav"] == selected_category]
    subcategory_counts = (
        pool[pool["subcategory_nav"].astype(str).str.strip().ne("")]
        .groupby("subcategory_nav").size().sort_values(ascending=False).head(16)
    )
    if subcategory_counts.empty:
        return

    st.markdown("#### Подкатегории")
    cols = st.columns(4)
    for idx, (subcategory_name, count) in enumerate(subcategory_counts.items()):
        with cols[idx % 4]:
            if st.button(f"{subcategory_name} ({count})", key=f"subcat_shortcut_{idx}", use_container_width=True):
                st.query_params["subcategory"] = subcategory_name
                st.query_params["page"] = "1"
                st.rerun()


def filter_catalog(df: pd.DataFrame) -> tuple[pd.DataFrame, str, str]:
    st.sidebar.header("Фильтры")
    query = st.sidebar.text_input("Поиск", placeholder="Артикул, код, название, бренд, категория")

    category_options = sorted([x for x in df["category_nav"].dropna().unique().tolist() if clean_str(x)])
    category_from_query = clean_str(st.query_params.get("category", ""))
    category_index = category_options.index(category_from_query) + 1 if category_from_query in category_options else 0

    selected_category = st.sidebar.selectbox("Категория", ["Все категории"] + category_options, index=category_index)

    subcategory_pool = df.copy()
    if selected_category != "Все категории":
        subcategory_pool = subcategory_pool[subcategory_pool["category_nav"] == selected_category]

    subcategory_options = sorted([x for x in subcategory_pool["subcategory_nav"].dropna().unique().tolist() if clean_str(x)])
    subcategory_from_query = clean_str(st.query_params.get("subcategory", ""))
    subcategory_index = subcategory_options.index(subcategory_from_query) + 1 if subcategory_from_query in subcategory_options else 0

    selected_subcategory = st.sidebar.selectbox("Подкатегория", ["Все подкатегории"] + subcategory_options, index=subcategory_index)

    if selected_category != "Все категории":
        st.query_params["category"] = selected_category
    else:
        st.query_params.pop("category", None)

    if selected_subcategory != "Все подкатегории":
        st.query_params["subcategory"] = selected_subcategory
    else:
        st.query_params.pop("subcategory", None)

    selected_brands = st.sidebar.multiselect("Бренд", sorted([x for x in df["brand_clean"].dropna().unique().tolist() if clean_str(x)]))
    only_in_stock = st.sidebar.checkbox("Только в наличии")
    only_with_photo = st.sidebar.checkbox("Только с фото", value=True)

    min_price = int(df["ozon_price_num"].min()) if not df.empty else 0
    max_price = int(df["ozon_price_num"].max()) if not df.empty else 0
    price_range = st.sidebar.slider("Цена Ozon", min_price, max_price, (min_price, max_price)) if max_price > min_price else (min_price, max_price)

    sort_by = st.sidebar.selectbox("Сортировка", ["По умолчанию", "Номер: сначала", "Номер: конец", "Цена: ниже", "Цена: выше", "Название: А-Я", "Название: Я-А"])

    result = df.copy()
    if query.strip():
        result = result[result["search_blob"].str.contains(query.lower(), na=False)]
    if selected_category != "Все категории":
        result = result[result["category_nav"] == selected_category]
    if selected_subcategory != "Все подкатегории":
        result = result[result["subcategory_nav"] == selected_subcategory]
    if selected_brands:
        result = result[result["brand_clean"].isin(selected_brands)]
    if only_in_stock:
        result = result[result["stock_qty_num"] > 0]
    if only_with_photo:
        result = result[result["has_image"]]

    result = result[(result["ozon_price_num"] >= price_range[0]) & (result["ozon_price_num"] <= price_range[1])]

    if sort_by == "Номер: сначала":
        result = result.sort_values(["catalog_number"], kind="stable")
    elif sort_by == "Номер: конец":
        result = result.sort_values(["catalog_number"], ascending=False, kind="stable")
    elif sort_by == "Цена: ниже":
        result = result.sort_values(["ozon_price_num", "title"], kind="stable")
    elif sort_by == "Цена: выше":
        result = result.sort_values(["ozon_price_num", "title"], ascending=[False, True], kind="stable")
    elif sort_by == "Название: А-Я":
        result = result.sort_values(["title", "brand"], kind="stable")
    elif sort_by == "Название: Я-А":
        result = result.sort_values(["title", "brand"], ascending=[False, True], kind="stable")

    return result.reset_index(drop=True), selected_category, selected_subcategory


def card_info_row(label: str, value: Any) -> str:
    text = clean_str(value)
    if not text:
        return ""
    return (
        '<div class="full-info-row">'
        f'<span class="full-info-label">{html.escape(label)}:</span> {html.escape(text)}'
        "</div>"
    )


def render_card(row: pd.Series, show_full: bool = False) -> None:
    image_src = image_to_src(row["primary_image_resolved"])
    category = html.escape(clean_str(row["category_nav"]) or "—")
    subcategory = html.escape(clean_str(row["subcategory_nav"]) or "—")
    title = html.escape(clean_str(row["title"]))
    brand = html.escape(clean_str(row["brand"]))
    offer_id = html.escape(clean_str(row["offer_id"]))
    code = html.escape(clean_str(row["code"]))
    display_number = int(row.get("display_number", row["catalog_number"]))

    old_price = ""
    if row["retail_price_num"] and row["retail_price_num"] != row["ozon_price_num"]:
        old_price = f'<div class="old">Цена источника: {html.escape(format_price(row["retail_price_num"]))}</div>'

    image_html = f'<img class="product-image" src="{html.escape(image_src)}" alt="{title}">' if image_src else ""
    full_info = ""
    if show_full:
        dimensions = format_dimensions(row)
        weight = format_weight(row["weight"])
        full_rows = [
            card_info_row("Полное наименование", row["full_name"]),
            card_info_row("Группа", row["group_name"]),
            card_info_row("Ozon-категория", row["ozon_category_display"]),
            card_info_row("Вес", weight if weight != "—" else ""),
            card_info_row("Габариты Д×Ш×В", dimensions if dimensions != "—" else ""),
            card_info_row("ТН ВЭД", row["tnved"]),
            card_info_row("Материал", row["material"]),
            card_info_row("Цвет", row["color"]),
            card_info_row("Источник", row["external_url"]),
        ]
        description = clean_str(row["description"])
        description_html = (
            f'<div class="full-description">{html.escape(description)}</div>' if description else ""
        )
        full_info = f'<div class="full-info">{"".join(full_rows)}{description_html}</div>'

    st.markdown(
        f"""
        <a class="product-link" href="?offer_id={offer_id}" target="_self">
            <div class="product-card">
                <div class="num-badge">№ {display_number}</div>
                {image_html}
                <div class="brand">{brand}</div>
                <div class="title">{title}</div>
                <div class="price">{html.escape(format_price(row["ozon_price_num"]))}</div>
                {old_price}
                <div class="meta">
                    Категория: {category}<br>
                    Подкатегория: {subcategory}<br>
                    Артикул: {offer_id}<br>
                    Код: {code}
                </div>
                {full_info}
            </div>
        </a>
        """,
        unsafe_allow_html=True,
    )


def render_pagination(page: int, total_pages: int, key_prefix: str) -> None:
    left, mid, right = st.columns([1, 2, 1])
    with left:
        if page > 1 and st.button("← Назад", key=f"{key_prefix}_prev", use_container_width=True):
            st.session_state["scroll_to_catalog_top"] = True
            st.query_params["page"] = str(page - 1)
            st.rerun()
    with mid:
        st.markdown(
            f"<div style='text-align:center;padding-top:8px;'>Страница {page} из {total_pages}</div>",
            unsafe_allow_html=True,
        )
    with right:
        if page < total_pages and st.button("Вперёд →", key=f"{key_prefix}_next", use_container_width=True):
            st.session_state["scroll_to_catalog_top"] = True
            st.query_params["page"] = str(page + 1)
            st.rerun()


def render_catalog_page(df: pd.DataFrame) -> None:
    df = df.copy().reset_index(drop=True)
    df["display_number"] = range(1, len(df) + 1)

    total_pages = max(1, math.ceil(len(df) / PAGE_SIZE))
    page = int(st.query_params.get("page", "1"))
    page = max(1, min(page, total_pages))

    start = (page - 1) * PAGE_SIZE
    end = start + PAGE_SIZE

    _, toggle_col = st.columns([3, 1])
    with toggle_col:
        show_full_cards = st.checkbox("Вся инфа в карточках", value=False, key="show_full_cards")

    st.markdown('<div id="catalog-top-anchor"></div>', unsafe_allow_html=True)
    if st.session_state.pop("scroll_to_catalog_top", False):
        components.html(
            """
            <script>
            const anchor = window.parent.document.getElementById("catalog-top-anchor");
            if (anchor) {
                anchor.scrollIntoView({behavior: "auto", block: "start"});
            }
            </script>
            """,
            height=0,
        )

    render_pagination(page, total_pages, "top")

    items = list(df.iloc[start:end].iterrows())

    # Важно: рисуем карточки СТРОКАМИ по 4 штуки, а не одной колонкой через idx % 4.
    # Тогда на широком экране порядок: 1 2 3 4 / 5 6 7 8.
    # На вертикальном телефоне Streamlit сложит колонки сверху вниз: 1,2,3,4,5,6...
    # Старый вариант давал на телефоне 1,5,9... из-за вертикального заполнения колонок.
    for row_start in range(0, len(items), 4):
        cols = st.columns(4)
        for col, (_, row) in zip(cols, items[row_start:row_start + 4]):
            with col:
                render_card(row, show_full=show_full_cards)

    if total_pages > 1:
        st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
        render_pagination(page, total_pages, "bottom")


def render_product_page(df: pd.DataFrame, offer_id: str) -> None:
    item = df[df["offer_id"].astype(str) == str(offer_id)]
    if item.empty:
        st.error("Товар не найден")
        return

    row = item.iloc[0]
    st.markdown("[← Назад в каталог](./)")
    st.markdown('<div class="detail">', unsafe_allow_html=True)

    category = clean_str(row["category_nav"])
    subcategory = clean_str(row["subcategory_nav"])
    ozon_category = clean_str(row["ozon_category_display"])
    dimensions = format_dimensions(row)

    st.markdown(f'<div class="num-badge">№ {int(row["catalog_number"])}</div>', unsafe_allow_html=True)
    st.caption(" / ".join([x for x in [category, subcategory] if x]))
    st.title(row["title"])

    if clean_str(row["full_name"]) and clean_str(row["full_name"]) != clean_str(row["title"]):
        st.write(row["full_name"])

    left, right = st.columns([1.05, 1])
    with left:
        if row["primary_image_resolved"]:
            st.image(row["primary_image_resolved"], width="stretch")

    with right:
        st.markdown(f"### {format_price(row['ozon_price_num'])}")
        if row["retail_price_num"] and row["retail_price_num"] != row["ozon_price_num"]:
            price_source = clean_str(row.get("price_source"))
            source_suffix = f" ({price_source})" if price_source else ""
            st.markdown(f"Цена источника{source_suffix}: {format_price(row['retail_price_num'])}")
        st.markdown(f"**Наличие:** {row['stock_text']}")
        if clean_str(row["external_url"]):
            st.link_button("Открыть источник", clean_str(row["external_url"]))

        specs = [
            ("Порядковый номер", int(row["catalog_number"])),
            ("Категория", category or "—"),
            ("Подкатегория", subcategory or "—"),
            ("Ozon-категория", ozon_category or "—"),
            ("Артикул", row["offer_id"]),
            ("Код", row["code"]),
            ("Бренд", row["brand"]),
            ("Вес", format_weight(row["weight"])),
            ("Габариты Д×Ш×В", dimensions),
            ("ТН ВЭД", clean_str(row["tnved"]) or "—"),
        ]

        sc1, sc2 = st.columns(2)
        for idx, (label, value) in enumerate(specs):
            with (sc1 if idx % 2 == 0 else sc2):
                st.markdown(f'<div class="spec"><div class="spec-label">{label}</div><div class="spec-value">{value}</div></div>', unsafe_allow_html=True)

    if clean_str(row["description"]):
        st.markdown("### Описание")
        st.write(clean_str(row["description"]))

    st.markdown("</div>", unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title="Catalog 500 Exact", page_icon="🛒", layout="wide")
    inject_css()
    df = load_catalog()

    st.markdown(
        """
        <div class="topbar">
            <h1>Catalog 500 Exact</h1>
            <p>Витрина: 500 не метражных товаров с точным совпадением по артикулу/коду, фото, категориями и supplier-ценами.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    offer_id = st.query_params.get("offer_id", "")
    if offer_id:
        render_product_page(df, offer_id)
        return

    filtered, selected_category, selected_subcategory = filter_catalog(df)
    render_metrics(filtered)
    render_category_shortcuts(df)
    render_subcategory_shortcuts(df, selected_category)

    st.write(f"Показано товаров: {len(filtered):,}".replace(",", " "))
    render_catalog_page(filtered)


if __name__ == "__main__":
    main()
