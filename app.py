from __future__ import annotations

from html import escape
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Utilities CHF Streamlit", layout="wide")

DATA_PATH = Path(__file__).resolve().parent / "Market data utilities.xlsx"
LOGO_PATH = Path(__file__).resolve().parent / "Tradition_LOGO.png"

TARGET_ISSUERS = [
    "IWB IND WERKE BASEL",
    "AEW ENERGIE AG",
    "LIMECO",
    "WWZ AG",
    "SIG",
    "BKW AG",
    "GROUPE E SA",
    "KRAFTWERKE OBERHASLI",
    "ENGADINER KRAFTWERKE",
    "NANT DE DRANCE SA",
    "PRIMEO HOLDING AG",
    "GRANDE DIXENCE",
]

RATING_ORDER = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-"]
GROUP_ORDER = ["AA", "A", "BBB+"]
GROUP_LINE_COLORS = {
    "AA": "#9ca3af",
    "A": "#4b5563",
    "BBB+": "#d6dae1",
}
PALETTE = [
    "#2563eb", "#dc2626", "#16a34a", "#f59e0b", "#7c3aed", "#06b6d4",
    "#ef4444", "#84cc16", "#ec4899", "#0f766e", "#9333ea", "#ea580c",
    "#0891b2", "#db2777", "#65a30d", "#1d4ed8", "#0ea5e9", "#b45309",
    "#be123c", "#4f46e5", "#059669", "#c2410c", "#0f766e", "#a855f7",
    "#3b82f6", "#22c55e", "#f97316", "#e11d48", "#14b8a6", "#6366f1",
]


def rating_group(rating: str) -> str | None:
    if rating.startswith("AA"):
        return "AA"
    if rating.startswith("A"):
        return "A"
    if rating == "BBB+":
        return "BBB+"
    return None


def rating_sort_key(rating: str) -> tuple[int, str]:
    try:
        return (RATING_ORDER.index(rating), rating)
    except ValueError:
        return (len(RATING_ORDER), rating)


def fit_line(xs: pd.Series, ys: pd.Series) -> tuple[float, float]:
    if len(xs) < 2:
        return 0.0, float(ys.iloc[0])
    x_mean = float(xs.mean())
    y_mean = float(ys.mean())
    denom = float(((xs - x_mean) ** 2).sum())
    if denom == 0:
        return 0.0, y_mean
    slope = float(((xs - x_mean) * (ys - y_mean)).sum()) / denom
    intercept = y_mean - slope * x_mean
    return slope, intercept


@st.cache_data
def load_data(path: str, file_mtime: float) -> tuple[pd.DataFrame, dict[str, str], dict[str, str]]:
    raw = pd.read_excel(path)
    df = pd.DataFrame(
        {
            "entity": raw.iloc[:, 2],
            "rating_zkb": raw.iloc[:, 3],
            "residual_duration": raw.iloc[:, 14],
            "spread_market": raw.iloc[:, 18],
        }
    )
    df = df.dropna(subset=["entity", "rating_zkb", "residual_duration", "spread_market"])
    df["entity"] = df["entity"].astype(str).str.strip()
    df["rating_zkb"] = df["rating_zkb"].astype(str).str.strip()
    df["residual_duration"] = pd.to_numeric(df["residual_duration"], errors="coerce")
    df["spread_market"] = pd.to_numeric(df["spread_market"], errors="coerce")
    df = df.dropna(subset=["residual_duration", "spread_market"])
    df["rating_group"] = df["rating_zkb"].map(rating_group)
    df = df.dropna(subset=["rating_group"]).reset_index(drop=True)

    issuer_ratings = df.groupby("entity")["rating_zkb"].agg(lambda s: s.mode().iloc[0]).to_dict()
    all_issuers = sorted(df["entity"].unique().tolist(), key=lambda x: (rating_sort_key(issuer_ratings.get(x, "")), x))
    issuer_colors = {issuer: PALETTE[i % len(PALETTE)] for i, issuer in enumerate(all_issuers)}
    df["issuer_color"] = df["entity"].map(issuer_colors)
    return df, issuer_ratings, issuer_colors


def make_legend_df(df: pd.DataFrame, issuer_ratings: dict[str, str], issuer_colors: dict[str, str]) -> pd.DataFrame:
    issuers = sorted(df["entity"].unique().tolist(), key=lambda x: (rating_sort_key(issuer_ratings.get(x, "")), x))
    return pd.DataFrame(
        {
            "entity": issuers,
            "rating_zkb": [issuer_ratings.get(i, "") for i in issuers],
            "color": [issuer_colors.get(i, "#94a3b8") for i in issuers],
        }
    )


def make_line_data(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    trend_rows: list[dict[str, float | str]] = []
    label_rows: list[dict[str, float | str]] = []
    for group in GROUP_ORDER:
        subset = df.loc[df["rating_group"] == group].sort_values("residual_duration")
        if subset.empty:
            continue
        slope, intercept = fit_line(subset["residual_duration"], subset["spread_market"])
        x1 = float(subset["residual_duration"].min())
        x2 = float(subset["residual_duration"].max())
        y1 = slope * x1 + intercept
        y2 = slope * x2 + intercept
        trend_rows.extend(
            [
                {"group": group, "x": x1, "y": y1, "color": GROUP_LINE_COLORS[group]},
                {"group": group, "x": x2, "y": y2, "color": GROUP_LINE_COLORS[group]},
            ]
        )
        label_rows.append({"group": group, "x": x2, "y": y2, "color": GROUP_LINE_COLORS[group]})
    return pd.DataFrame(trend_rows), pd.DataFrame(label_rows)


if LOGO_PATH.exists():
    st.image(str(LOGO_PATH), width=360)

st.title("Comparatif des spreads sur le marché secondaire obligataire : entreprises d'énergie")

if not DATA_PATH.exists():
    st.error(f"Fichier introuvable : {DATA_PATH}")
    st.stop()

file_mtime = DATA_PATH.stat().st_mtime
all_df, issuer_ratings, issuer_colors = load_data(str(DATA_PATH), file_mtime)

focus_shortlist = st.toggle("Focus shortlist", value=False, help="Affiche uniquement la shortlist d'émetteurs mise en avant.")

if focus_shortlist:
    df = all_df.loc[all_df["entity"].isin(TARGET_ISSUERS)].copy()
else:
    df = all_df.copy()

if df.empty:
    st.warning("Aucune donnée exploitable pour les émetteurs sélectionnés.")
    st.stop()

legend_df = make_legend_df(df, issuer_ratings, issuer_colors)
trend_source, label_df = make_line_data(df)

chart_df = df[["entity", "rating_zkb", "residual_duration", "spread_market"]].rename(
    columns={
        "entity": "Entite",
        "rating_zkb": "Rating ZKB",
        "residual_duration": "Duration residuelle",
        "spread_market": "Spread marche",
    }
)

x_max = max(8, int(df["residual_duration"].max().round()))
y_min = int(max(20, (df["spread_market"].min() // 5) * 5))
y_max = int(((df["spread_market"].max() + 4) // 5 + 1) * 5)

base = alt.Chart(chart_df).encode(
    x=alt.X(
        "Duration residuelle:Q",
        title="Duration residuelle",
        scale=alt.Scale(domain=[0, x_max], nice=False),
        axis=alt.Axis(values=list(range(0, x_max + 1)), labelExpr="datum.value + 'y'", labelFontSize=16, titleFontSize=19),
    ),
    y=alt.Y(
        "Spread marche:Q",
        title="Spread marche",
        scale=alt.Scale(domain=[y_min, y_max], nice=False),
        axis=alt.Axis(values=list(range(y_min, y_max + 1, 5)), labelFontSize=16, titleFontSize=19),
    ),
)

points = base.mark_circle(size=115, opacity=0.95).encode(
    color=alt.Color(
        "Entite:N",
        scale=alt.Scale(domain=legend_df["entity"].tolist(), range=legend_df["color"].tolist()),
        legend=None,
    ),
    tooltip=[
        alt.Tooltip("Entite:N", title="Entite"),
        alt.Tooltip("Rating ZKB:N", title="Rating ZKB"),
        alt.Tooltip("Duration residuelle:Q", title="Duration residuelle", format=".2f"),
        alt.Tooltip("Spread marche:Q", title="Spread marche", format=".2f"),
    ],
)

lines = alt.Chart(trend_source).mark_line(strokeWidth=3.8).encode(
    x="x:Q",
    y="y:Q",
    detail="group:N",
    color=alt.Color("color:N", scale=None, legend=None),
)

line_labels = alt.Chart(label_df).mark_text(dx=12, dy=-6, fontSize=15, fontWeight="bold").encode(
    x="x:Q",
    y="y:Q",
    text="group:N",
    color=alt.Color("color:N", scale=None, legend=None),
)

chart = (
    (points + lines + line_labels)
    .properties(width=1150, height=650)
     .configure_axis(gridColor="#d9dee7", domainColor="#475569", tickColor="#475569", labelColor="#475569", titleColor="#475569", labelFontSize=16, titleFontSize=19)
    .configure_view(stroke=None)
    .interactive()
)

left, right = st.columns([4.8, 1.8], gap="large")
with left:
    st.altair_chart(chart, width="stretch")

with right:
    st.subheader("Émetteurs")
    for _, row in legend_df.iterrows():
        bullet = row["color"]
        st.markdown(
            f"""
            <div style="display:flex; align-items:center; gap:10px; margin:8px 0;">
                <div style="width:14px; height:14px; border-radius:50%; background:{bullet}; flex:0 0 14px;"></div>
                <div style="flex:1; font-size:14px;">{escape(row['entity'])}</div>
                <div style="font-size:13px; color:#475569; min-width:28px; text-align:right;">{escape(row['rating_zkb'])}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.caption("Colonnes utilisées : C = Entité, D = Rating ZKB, O = Duration résiduelle, S = Spread marché.")



