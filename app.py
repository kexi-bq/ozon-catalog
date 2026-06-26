from __future__ import annotations

import math
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_PATH = APP_DIR / "data" / "catalog_500.xlsx"
PAGE_SIZE = 24


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


def resolve_image(value: str) -> str:
    text = clean_str(value)
    if not text:
        return ""
    if text.startswith(("http://", "https://")):
        return text
    path = APP_DIR / text
    if path.exists():
        return str(path)
    return ""


@st.cache_data(show_spinner=True)
def load_catalog() -> pd.DataFrame:
    df = pd.read_excel(DATA_PATH, dtype=object)
    df["retail_price_num"] = df["retail_price"].map(clean_float).fillna(0.0)
    df["ozon_price_num"] = df["ozon_price"].map(clean_float).fillna(df["retail_price_num"])
    df["stock_qty_num"] = df["stock_qty"].map(clean_float).fillna(0)
    df["primary_image_resolved"] = df["primary_image"].map(resolve_image)
    df["has_image"] = df["primary_image_resolved"].astype(str).str.strip().ne("")
    df["search_blob"] = (
        df["offer_id"].fillna("").astype(str)
        + " "
        + df["code"].fillna("").astype(str)
        + " "
        + df["title"].fillna("").astype(str)
        + " "
        + df["full_name"].fillna("").astype(str)
        + " "
        + df["brand"].fillna("").astype(str)
        + " "
        + df["group_name"].fillna("").astype(str)
    ).str.lower()
    return df.sort_values(["brand", "title", "offer_id"], kind="stable").reset_index(drop=True)


def format_price(value: float | int | None) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    return f"{int(math.ceil(float(value))):,} ₽".replace(",", " ")


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
        .brand { font-size: 11px; color: #005bff; font-weight: 700; text-transform: uppercase; }
        .title { font-size: 15px; line-height: 1.35; font-weight: 700; color: #162033; min-height: 62px; margin: 8px 0 10px 0; }
        .price { font-size: 26px; font-weight: 800; color: #f91155; margin: 8px 0 2px 0; }
        .old { font-size: 13px; color: #8a97a8; text-decoration: line-through; }
        .meta { font-size: 12px; color: #607087; line-height: 1.55; margin-top: 10px; }
        .badge {
            display: inline-block; padding: 5px 10px; border-radius: 999px; font-size: 12px; font-weight: 700;
            margin-top: 8px; background: #eaf8ef; color: #138a43;
        }
        .badge.out { background: #fff2e2; color: #a45d00; }
        .detail { background: white; border: 1px solid #dfe7f2; border-radius: 24px; padding: 24px; }
        .spec { background: #f8fbff; border: 1px solid #e4edf9; border-radius: 18px; padding: 14px 16px; margin-bottom: 10px; }
        .spec-label { font-size: 12px; color: #718198; margin-bottom: 4px; }
        .spec-value { font-size: 15px; color: #18253a; font-weight: 600; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_metrics(df: pd.DataFrame) -> None:
    c1, c2, c3, c4 = st.columns(4)
    values = [
        ("Товаров", f"{len(df):,}".replace(",", " ")),
        ("С фото", f"{int(df['has_image'].sum()):,}".replace(",", " ")),
        ("Брендов", f"{int(df['brand'].nunique()):,}".replace(",", " ")),
        ("Средняя цена", format_price(df["ozon_price_num"].mean() if not df.empty else 0)),
    ]
    for col, (label, value) in zip((c1, c2, c3, c4), values):
        with col:
            st.markdown(
                f'<div class="metric"><div class="metric-label">{label}</div><div class="metric-value">{value}</div></div>',
                unsafe_allow_html=True,
            )


def filter_catalog(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header("Фильтры")
    query = st.sidebar.text_input("Поиск", placeholder="Артикул, код, название, бренд, группа")
    selected_brands = st.sidebar.multiselect("Бренд", sorted(df["brand"].dropna().unique().tolist()))
    selected_groups = st.sidebar.multiselect("Группа", sorted(df["group_name"].dropna().unique().tolist()))
    only_in_stock = st.sidebar.checkbox("Только в наличии")
    min_price = int(df["ozon_price_num"].min()) if not df.empty else 0
    max_price = int(df["ozon_price_num"].max()) if not df.empty else 0
    price_range = st.sidebar.slider("Цена Ozon", min_price, max_price, (min_price, max_price)) if max_price > min_price else (min_price, max_price)
    sort_by = st.sidebar.selectbox("Сортировка", ["По умолчанию", "Цена: ниже", "Цена: выше", "Название: А-Я", "Название: Я-А"])

    result = df.copy()
    if query.strip():
        result = result[result["search_blob"].str.contains(query.lower(), na=False)]
    if selected_brands:
        result = result[result["brand"].isin(selected_brands)]
    if selected_groups:
        result = result[result["group_name"].isin(selected_groups)]
    if only_in_stock:
        result = result[result["stock_qty_num"] > 0]
    result = result[(result["ozon_price_num"] >= price_range[0]) & (result["ozon_price_num"] <= price_range[1])]

    if sort_by == "Цена: ниже":
        result = result.sort_values(["ozon_price_num", "title"], kind="stable")
    elif sort_by == "Цена: выше":
        result = result.sort_values(["ozon_price_num", "title"], ascending=[False, True], kind="stable")
    elif sort_by == "Название: А-Я":
        result = result.sort_values(["title", "brand"], kind="stable")
    elif sort_by == "Название: Я-А":
        result = result.sort_values(["title", "brand"], ascending=[False, True], kind="stable")
    return result.reset_index(drop=True)


def render_card(row: pd.Series) -> None:
    with st.container(border=True):
        if row["primary_image_resolved"]:
            st.image(row["primary_image_resolved"], width="stretch")
        st.markdown(f'<div class="brand">{row["brand"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="title">{row["title"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="price">{format_price(row["ozon_price_num"])}</div>', unsafe_allow_html=True)
        if row["retail_price_num"] and row["retail_price_num"] != row["ozon_price_num"]:
            st.markdown(f'<div class="old">{format_price(row["retail_price_num"])}</div>', unsafe_allow_html=True)
        badge_class = "badge" if row["stock_qty_num"] > 0 else "badge out"
        st.markdown(f'<div class="{badge_class}">{row["stock_text"]}</div>', unsafe_allow_html=True)
        st.markdown(
            f'<div class="meta">Артикул: {row["offer_id"]}<br>Код: {row["code"]}<br>Группа: {row["group_name"]}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(f"[Открыть карточку](?offer_id={row['offer_id']})")


def render_catalog_page(df: pd.DataFrame) -> None:
    total_pages = max(1, math.ceil(len(df) / PAGE_SIZE))
    page = int(st.query_params.get("page", "1"))
    page = max(1, min(page, total_pages))
    start = (page - 1) * PAGE_SIZE
    end = start + PAGE_SIZE

    left, mid, right = st.columns([1, 2, 1])
    with left:
        if page > 1 and st.button("← Назад"):
            st.query_params["page"] = str(page - 1)
            st.rerun()
    with mid:
        st.markdown(f"<div style='text-align:center;padding-top:8px;'>Страница {page} из {total_pages}</div>", unsafe_allow_html=True)
    with right:
        if page < total_pages and st.button("Вперёд →"):
            st.query_params["page"] = str(page + 1)
            st.rerun()

    items = df.iloc[start:end]
    cols = st.columns(4)
    for idx, (_, row) in enumerate(items.iterrows()):
        with cols[idx % 4]:
            render_card(row)


def render_product_page(df: pd.DataFrame, offer_id: str) -> None:
    item = df[df["offer_id"].astype(str) == str(offer_id)]
    if item.empty:
        st.error("Товар не найден")
        return

    row = item.iloc[0]
    st.markdown("[← Назад в каталог](./)")
    st.markdown('<div class="detail">', unsafe_allow_html=True)
    st.caption(row["group_name"])
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
            st.markdown(f"Старая цена: {format_price(row['retail_price_num'])}")
        st.markdown(f"**Наличие:** {row['stock_text']}")
        if clean_str(row["external_url"]):
            st.link_button("Открыть источник", clean_str(row["external_url"]))

        specs = [
            ("Артикул", row["offer_id"]),
            ("Код", row["code"]),
            ("Бренд", row["brand"]),
            ("Группа", row["group_name"]),
            ("Вес", f"{row['weight']} кг" if clean_str(row["weight"]) else "—"),
            ("Длина", clean_str(row["length"]) or "—"),
            ("Ширина", clean_str(row["width"]) or "—"),
            ("Высота", clean_str(row["height"]) or "—"),
            ("TN VED", clean_str(row["tnved"]) or "—"),
        ]
        sc1, sc2 = st.columns(2)
        for idx, (label, value) in enumerate(specs):
            with (sc1 if idx % 2 == 0 else sc2):
                st.markdown(
                    f'<div class="spec"><div class="spec-label">{label}</div><div class="spec-value">{value}</div></div>',
                    unsafe_allow_html=True,
                )

    if clean_str(row["description"]):
        st.markdown("### Описание")
        st.write(clean_str(row["description"]))
    st.markdown("</div>", unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(page_title="Catalog 500", page_icon="🛒", layout="wide")
    inject_css()
    df = load_catalog()

    st.markdown(
        """
        <div class="topbar">
            <h1>Catalog 500</h1>
            <p>Git-ready витрина: 500 не метражных товаров с фото из локальной базы.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    offer_id = st.query_params.get("offer_id", "")
    if offer_id:
        render_product_page(df, offer_id)
        return

    filtered = filter_catalog(df)
    render_metrics(filtered)
    st.write(f"Показано товаров: {len(filtered):,}".replace(",", " "))
    render_catalog_page(filtered)


if __name__ == "__main__":
    main()
