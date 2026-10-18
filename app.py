"""Painel Streamlit Devolução - TCv2 alimentado pelo arquivo data.xlsx."""

from __future__ import annotations

import html
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

from data import (
    AGE_BUCKETS,
    PBIX_ENTRY_AGE_BUCKETS,
    PBIX_ENTRY_AGE_AS_OF,
    VALUE_BANDS,
    load_source_data,
    powerbi_aging_bucket,
)


ROOT = Path(__file__).resolve().parent
LOGO_PATH = ROOT / "images" / "jadlog_logo.webp"
ICON_PATH = ROOT / "images" / "jadlog_icon.webp"
DATA_PATH = ROOT / "data.xlsx"
RED = "#e5003d"
PAGES = ["BACKLOG", "ENTRADA", "CONTROLE CLIENTES", "CONTROLE LOJAS", "FAIXA DE VALOR", "Lista CTE"]
BUCKET_LABELS = AGE_BUCKETS

st.set_page_config(
    page_title="Devolução | TC Reversas Jad-SAO",
    page_icon=str(ICON_PATH),
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_data(show_spinner=False)
def load_dashboard_data(
    source_path: str, modified_ns: int, file_size: int
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    # File metadata is part of the cache key so replacing the workbook refreshes the app.
    del modified_ns, file_size
    return load_source_data(source_path)


if not DATA_PATH.is_file():
    st.error(f"Base não encontrada: {DATA_PATH}. Coloque data.xlsx junto de app.py.")
    st.stop()
source_stat = DATA_PATH.stat()
BASE, CONTROL_CLIENTS, CONTROL_STORES = load_dashboard_data(
    str(DATA_PATH), source_stat.st_mtime_ns, source_stat.st_size
)


def category_options(column: str) -> list[str]:
    return sorted(BASE[column].dropna().astype(str).unique().tolist())


SHIPPER_NAMES = category_options("SHIPPER")
EMITTER_NAMES = category_options("EMISSOR")
UNIT_NAMES = category_options("PROXIMA_UNIDADE")
STATUS_NAMES = category_options("STATUS_ATUAL")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: Roboto, Arial, sans-serif; color:#202939; }
    .stApp { background:#f5f6fa; }
    [data-testid="stHeader"] { background:transparent; }
    [data-testid="stToolbar"] { right:12px; }
    [data-testid="stSidebar"] { background:#fff; border-right:1px solid #e7e9ef; min-width:238px; max-width:238px; }
    [data-testid="stSidebar"] > div { padding:20px 14px 14px; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { margin:0; }
    [data-testid="stSidebar"] h3 { font-size:12px; color:#8a91a1; font-weight:700; letter-spacing:.09em; margin:18px 10px 8px; text-transform:uppercase; }
    [data-testid="stSidebar"] div[role="radiogroup"] { gap:4px; }
    [data-testid="stSidebar"] div[role="radiogroup"] label { border-radius:9px; padding:8px 10px; min-height:39px; border:1px solid transparent; font-size:13px; transition:all .15s ease; }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover { background:#f7f8fb; }
    [data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child { display:none; }
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) { background:#fff1f4; border-color:#ffe0e8; color:#c90036; font-weight:650; }
    [data-testid="stSidebar"] div[role="radiogroup"] label p { font-size:13px; }
    .block-container { max-width:1580px; padding:1.35rem 1.55rem 1.4rem; }
    .report-header { display:grid; grid-template-columns:1fr 165px; align-items:center; gap:16px; min-height:76px; border:1px solid #eaecf1; border-radius:15px; background:#fff; padding:12px 18px; margin:0 0 17px; box-shadow:0 2px 8px #20293908; }
    .report-kicker { color:#8a91a1; font-size:10px; font-weight:650; letter-spacing:.08em; text-transform:uppercase; margin-bottom:4px; }
    .report-title { display:flex; align-items:baseline; justify-content:flex-start; color:#182230; line-height:1.15; font-weight:600; white-space:nowrap; }
    .report-title .page-name { color:#c90036; font-size:25px; font-weight:700; letter-spacing:-.02em; }
    .report-title .divider { color:#cdd2db; font-size:20px; margin:0 10px; }
    .report-title .branch { color:#647084; font-size:15px; font-weight:500; }
    .report-subtitle { color:#8790a0; text-align:right; font-size:10px; line-height:1.45; }
    .report-subtitle strong { display:inline-block; color:#263246; font-size:12px; font-weight:600; background:#f4f6f9; border-radius:20px; padding:3px 8px; margin-top:2px; }
    .section-title { color:#202939; font-size:14px; font-weight:650; padding:0 0 5px; margin:2px 0 7px; }
    .visual-card { min-width:0; height:100%; background:#fff; border:1px solid #eaecf1; border-radius:13px; padding:12px; overflow:hidden; box-shadow:0 2px 8px #20293908; }
    .panel-title { color:#253147; font-size:14px; font-weight:650; line-height:1.25; margin:0 0 8px; text-transform:none; }
    div[data-testid="stButton"] { width:100%; display:flex; justify-content:flex-end; }
    div[data-testid="stButton"] button { width:16px !important; min-width:16px !important; height:16px !important; min-height:16px !important; margin:0; padding:0; border:0 !important; border-radius:0; background:transparent !important; box-shadow:none !important; color:#566176; font-size:14px; line-height:16px; transform:none; }
    div[data-testid="stButton"] button:hover, div[data-testid="stButton"] button:focus { border:0 !important; background:transparent !important; box-shadow:none !important; color:#c90036; }
    div[data-testid="stMetric"] { background:transparent; border:0; padding:0; min-height:0; }
    div[data-testid="stMetricLabel"] { color:#788295 !important; font-size:10px !important; }
    div[data-testid="stMetricValue"] { color:#202939; font-size:21px !important; font-weight:700; letter-spacing:-.02em; white-space:nowrap; }
    .kpi-icon { display:grid; place-items:center; width:38px; height:38px; background:#fff0f4; color:#c90036; font-size:19px; font-weight:700; border-radius:12px; }
    div[data-testid="stSelectbox"] label, div[data-testid="stMultiSelect"] label, div[data-testid="stDateInput"] label, div[data-testid="stTextInput"] label, div[data-testid="stCheckbox"] label, div[data-testid="stRadio"] > label, div[data-testid="stSegmentedControl"] label { color:#687386 !important; font-size:10px !important; font-weight:600 !important; text-transform:none; }
    div[data-testid="stSelectbox"], div[data-testid="stMultiSelect"], div[data-testid="stDateInput"], div[data-testid="stTextInput"] { margin:0 0 4px; }
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div, div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div { min-height:38px; border-radius:9px; border-color:#e1e5ec; background:#fff; }
    div[data-testid="stDateInput"] input, div[data-testid="stTextInput"] input { min-height:37px; font-size:12px; border-radius:9px; }
    div[data-testid="stDateInput"] label { margin-bottom:2px; }
    div[data-testid="stDateInput"] [data-baseweb="input"] { min-height:37px; border-radius:9px; }
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] { border-radius:7px; background:#fff1f4; color:#bd0031; }
    div[data-testid="stSegmentedControl"] { margin:0; }
    div[data-testid="stSegmentedControl"] button { min-height:35px; padding:5px 10px; font-size:11px; }
    div[data-testid="stRadio"] div[role="radiogroup"] { gap:3px; }
    div[data-testid="stRadio"] div[role="radiogroup"] label { background:#fff; border:1px solid #e5e8ee; border-radius:9px; padding:5px 9px; font-size:11px; min-height:32px; }
    [data-testid="stVerticalBlockBorderWrapper"] { border:1px solid #eaecf1 !important; border-radius:13px !important; box-shadow:0 2px 8px #20293908; background:#fff; }
    [data-testid="stDataFrame"] { border:1px solid #eef0f4; border-radius:10px; overflow:hidden; }
    [data-testid="stDataFrame"] iframe { border-radius:10px; }
    .filter-box { min-height:69px; border:1px solid #eaecf1; border-radius:10px; padding:7px 9px; }
    .filter-title { color:#687386; font-size:10px; font-weight:600; margin:0 0 3px; }
    .trend-card { background:#fff; border:1px solid #eaecf1; border-radius:13px; padding:12px; box-shadow:0 2px 8px #20293908; }
    .micro-note { color:#788295; font-size:10px; padding:4px; }
    .segmented-wrap { display:flex; justify-content:flex-end; }
    hr { margin:.5rem 0; border-color:#edf0f4; }
    @media (max-width:1100px) { .block-container { padding:1rem; } .report-header { grid-template-columns:1fr 120px; gap:8px; padding:10px 12px; } .report-title .page-name { font-size:19px; } .report-title .branch { font-size:12px; } .report-title .divider { margin:0 5px; } .report-subtitle { font-size:9px; } div[data-testid="stMetricValue"] { font-size:17px !important; } .kpi-icon { width:32px; height:32px; } }
    </style>
    """,
    unsafe_allow_html=True,
)


def number_br(value: int | float) -> str:
    return f"{int(value):,}".replace(",", ".")


def decimal_br(value: float, digits: int = 2) -> str:
    return f"{value:,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def money_short(value: float) -> str:
    if value >= 1_000_000:
        return f"R$ {value / 1_000_000:.2f} Mi".replace(".", ",")
    return f"R$ {value / 1_000:.3f} Mil".replace(".", ",")


def count_short(value: int | float) -> str:
    return f"{value / 1_000:.3f} Mil".replace(".", ",")


def esc(value: object) -> str:
    return html.escape("" if pd.isna(value) else str(value))


def is_total_row(frame: pd.DataFrame) -> bool:
    if frame.empty or not len(frame.columns):
        return False
    label = frame.iloc[-1, 0]
    return isinstance(label, str) and label.strip().casefold() == "total"


def sticky_table_html(frame: pd.DataFrame, *, bar_column: str | None = None) -> str:
    columns = list(frame.columns)

    def cell_value(value: object, column: str) -> str:
        if value is None or (not isinstance(value, tuple) and pd.isna(value)):
            return ""
        if bar_column and column == bar_column:
            try:
                raw_percentage, label = value if isinstance(value, tuple) else (value, None)
                percentage = max(0.0, min(float(raw_percentage), 100.0))
            except (TypeError, ValueError):
                percentage = 0.0
                label = None
            visible_label = esc(label) if label is not None else f"{decimal_br(percentage, 2)}%"
            return (
                '<div class="combined-progress"><div class="progress"><span style="width:'
                f'{percentage:.2f}%"></span></div><span class="progress-label">'
                f'{visible_label}</span></div>'
            )
        if pd.api.types.is_number(value):
            if float(value).is_integer():
                return number_br(value)
            return decimal_br(float(value), 2)
        return esc(value)

    def render_row(row: pd.Series, *, total: bool = False) -> str:
        cells = []
        for column, value in row.items():
            numeric = (
                (bar_column is not None and column == bar_column)
                or pd.api.types.is_numeric_dtype(frame[column])
            )
            classes = []
            if numeric:
                classes.append("numeric")
            if bar_column is not None and column == bar_column:
                classes.append("distribution")
            class_name = f' class="{" ".join(classes)}"' if classes else ""
            cells.append(f"<td{class_name}>{cell_value(value, str(column))}</td>")
        row_class = ' class="total-row"' if total else ""
        return f"<tr{row_class}>{''.join(cells)}</tr>"

    header = "".join(
        f'<th class="distribution">{esc(column)}</th>'
        if bar_column is not None and column == bar_column
        else f"<th>{esc(column)}</th>"
        for column in columns
    )
    body = "".join(render_row(row) for _, row in frame.iloc[:-1].iterrows())
    total = "".join(render_row(row, total=True) for _, row in frame.tail(1).iterrows())
    return f"""
    <style>
      * {{ box-sizing: border-box; }}
      html, body {{ width: 100%; height: 100%; margin: 0; padding: 0; background: #fff; font-family: Roboto, Arial, sans-serif; }}
      .scroll-area {{ height: 100%; max-height: 100%; overflow: auto; border: 1px solid #eef0f4; border-radius: 9px; scrollbar-color: #c7ccd5 transparent; scrollbar-width: thin; }}
      table {{ width: max-content; min-width: 100%; border-collapse: separate; border-spacing: 0; color: #30394a; font-size: 11px; line-height: 1.25; }}
      th, td {{ padding: 5px 8px; border-right: 1px solid #e8ebf0; border-bottom: 1px solid #e8ebf0; white-space: nowrap; }}
      th.distribution, td.distribution {{ width: 124px; min-width: 124px; max-width: 124px; }}
      th {{ position: sticky; top: 0; z-index: 2; background: #f7f8fa; color: #778092; font-weight: 600; text-align: left; }}
      td {{ background: #fff; }}
      tbody tr:nth-child(even) td {{ background: #fafbfc; }}
      td.numeric {{ text-align: right; font-variant-numeric: tabular-nums; }}
      tfoot td {{ position: sticky; bottom: 0; z-index: 3; background: #f1f3f7; color: #263246; font-weight: 700; border-top: 1px solid #d9dee7; box-shadow: 0 -2px 4px #2029390b; }}
      .combined-progress {{ display: flex; align-items: center; justify-content: flex-end; gap: 7px; min-width: 106px; }}
      .progress {{ display: inline-block; width: 56px; height: 7px; flex: 0 0 56px; overflow: hidden; vertical-align: middle; border-radius: 10px; background: #fff0f4; }}
      .progress span {{ display: block; height: 100%; background: #e5003d; border-radius: inherit; }}
      .progress-label {{ min-width: 43px; text-align: right; }}
    </style>
    <div class="scroll-area"><table><thead><tr>{header}</tr></thead>
      <tbody>{body}</tbody><tfoot>{total}</tfoot>
    </table></div>
    """


@st.dialog("Modo foco", width="large")
def show_table_focus(
    title: str,
    frame: pd.DataFrame,
    bar_column: str | None = None,
    *,
    native: bool = False,
) -> None:
    st.markdown(f"**{esc(title)}**")
    if native:
        st.dataframe(frame, use_container_width=True, hide_index=True, height=700)
        if "CTE" in frame.columns:
            st.caption(f"Total · {number_br(frame['CTE'].notna().sum())}")
        return
    focus_height = min(720, max(150, 54 + len(frame) * 32))
    components.html(
        sticky_table_html(frame, bar_column=bar_column),
        height=focus_height,
        scrolling=False,
    )


def table_panel(
    title: str,
    frame: pd.DataFrame,
    *,
    height: int = 220,
    dark_header: bool = False,
    bar_column: str | None = None,
    numeric_columns: tuple[str, ...] = (),
    total_note: str | None = None,
    max_rows: int | None = None,
) -> None:
    view = frame.copy()
    has_total = is_total_row(view)
    if max_rows is not None and len(view) > max_rows:
        view = pd.concat(
            [view.head(max_rows - 1), view.tail(1)] if has_total else [view.head(max_rows)],
            ignore_index=True,
        )
    if bar_column is None and "BARRA" in view.columns:
        bar_column = "BARRA"
    if bar_column and bar_column in view:
        if "%" in view.columns:
            bar_values = view[bar_column].tolist()
            percent_labels = view["%"].astype(str).tolist()
            view = view.drop(columns=[bar_column, "%"])
            view["%"] = list(zip(bar_values, percent_labels))
        else:
            view = view.rename(columns={bar_column: "%"})
        bar_column = "%"
    config: dict[str, object] = {}
    for column in view.columns:
        if column == bar_column or column in config:
            continue
        if column in numeric_columns or pd.api.types.is_numeric_dtype(view[column]):
            config[column] = st.column_config.NumberColumn(str(column), format="%d", width="small")
    with st.container(border=True):
        title_col, focus_col = st.columns([1, .1], gap="small", vertical_alignment="center")
        with title_col:
            st.markdown(f'<div class="panel-title">{esc(title)}</div>', unsafe_allow_html=True)
        with focus_col:
            focus_clicked = st.button("⛶", key=f"focus_{title.casefold().replace(' ', '_')}", help="Abrir em modo foco")
        if focus_clicked:
            show_table_focus(title, view, bar_column)
        if has_total and not view.empty:
            components.html(
                sticky_table_html(view, bar_column=bar_column),
                height=height,
                scrolling=False,
            )
        else:
            st.dataframe(view, use_container_width=True, hide_index=True, height=height, column_config=config)
        if total_note:
            st.caption(f"Total · {total_note}")


def kpi(label: str, value: str, icon: str, note: str = "") -> None:
    with st.container(border=True):
        icon_col, metric_col = st.columns([.22, 1], gap="small", vertical_alignment="center")
        with icon_col:
            st.markdown(f'<div class="kpi-icon">{esc(icon)}</div>', unsafe_allow_html=True)
        with metric_col:
            st.metric(label=label, value=value, help=note or None)


def metric_cards(columns: list, df: pd.DataFrame, *, show_open: bool = True) -> None:
    amount = float(df["VALOR_NF"].sum()) if not df.empty else 0.0
    count = int(df["CTE"].notna().sum()) if not df.empty else 0
    open_pct = float(df["FINALIZADO?"].eq("EM ABERTO").mean()) if count else 0.0
    with columns[0]:
        kpi("Soma de VALOR_NF", money_short(amount), "＄", "")
    with columns[1]:
        kpi("Contagem de CTE" if show_open else "Quantidade", count_short(count), "▤", "")
    if show_open:
        with columns[2]:
            kpi("% Em Aberto", f"{open_pct:.2%}".replace(".", ","), "↗", "")


def header(page: str, *, backlog_type: bool = False) -> None:
    page_title = "LISTA DE CTE" if page == "Lista CTE" else page.upper()
    data_cutoff = BASE["DATA_ATUAL"].max()
    cutoff_label = data_cutoff.strftime("%d/%m/%Y") if pd.notna(data_cutoff) else "Sem data"
    header_markup = (
        '<div class="report-header"><div><div class="report-kicker">JADLOG · OPERAÇÕES DE REVERSA</div>'
        f'<div class="report-title"><span class="page-name">{esc(page_title)}</span>'
        '<span class="divider">/</span><span class="branch">TC REVERSAS JAD-SAO</span></div></div>'
        f'<div class="report-subtitle"><b>Dados até</b><br><strong>{cutoff_label}</strong></div></div>'
    )
    if backlog_type:
        title_col, group_col = st.columns([6.0, 1.0], gap="small", vertical_alignment="center")
        with title_col:
            st.markdown(header_markup, unsafe_allow_html=True)
        with group_col:
            st.segmented_control("Grupo", ["Cliente", "Lojas e PAs", "Rua"], default=None, selection_mode="single", key="backlog_grupo", label_visibility="collapsed")
    else:
        st.markdown(header_markup, unsafe_allow_html=True)


def kpi_row(df: pd.DataFrame, *, show_open: bool = False, fixed_value: float | None = None) -> None:
    cols = st.columns([1, 1, 1] if show_open else [1, 1], gap="small")
    metric_cards(cols, df, show_open=show_open)


def period_picker(df: pd.DataFrame, field: str, label: str, key: str, *, default: tuple[date, date] | None = None) -> tuple[date, date] | None:
    values = pd.to_datetime(df[field], errors="coerce").dropna()
    if values.empty:
        return None
    lo, hi = values.min().date(), values.max().date()
    selected_default = default or (lo, hi)
    selected_default = (max(lo, selected_default[0]), min(hi, selected_default[1]))
    value = st.date_input(label, value=selected_default, min_value=lo, max_value=hi, format="DD/MM/YYYY", key=key)
    if isinstance(value, tuple) and len(value) == 2:
        return value
    return None


def apply_period(df: pd.DataFrame, field: str, period: tuple[date, date] | None) -> pd.DataFrame:
    if period is None:
        return df
    dates = pd.to_datetime(df[field], errors="coerce").dt.date
    return df.loc[dates.between(period[0], period[1])]


def filtered(df: pd.DataFrame, fields: dict[str, list[str]]) -> pd.DataFrame:
    out = df
    for field, values in fields.items():
        if values:
            out = out[out[field].astype(str).isin(values)]
    return out


def filter_select(label: str, options: list[str], key: str) -> list[str]:
    return st.multiselect(label, options, key=key, placeholder="Todos", label_visibility="visible")


def summary_table(df: pd.DataFrame, field: str, *, limit: int = 12) -> pd.DataFrame:
    grouped = df.groupby(field, dropna=False).size().rename("QTDE").sort_values(ascending=False).head(limit)
    total_count = len(df)
    base = max(total_count, 1)
    out = grouped.rename_axis(field).reset_index()
    out["%"] = out["QTDE"].map(lambda val: 100 * val / max(base, 1))
    # The bar and printed percentage use the same whole-table denominator.
    out["BARRA"] = out["%"]
    out["%"] = out["%"].map(lambda val: f"{decimal_br(val, 2)}%")
    total_pct = 100.0 if total_count else 0.0
    total = pd.DataFrame([{field: "Total", "QTDE": total_count, "%": f"{decimal_br(total_pct, 2)}%", "BARRA": total_pct}])
    return pd.concat([out, total], ignore_index=True)


def matrix_table(
    df: pd.DataFrame,
    row_field: str,
    bucket_field: str,
    *,
    limit: int = 12,
    order: list[str] | None = None,
    bucket_labels: list[str] | None = None,
) -> pd.DataFrame:
    labels = bucket_labels or BUCKET_LABELS
    if df.empty:
        return pd.DataFrame(columns=[row_field, *labels, "Total"])
    pivot = pd.crosstab(df[row_field].fillna("(Em branco)"), df[bucket_field].fillna("(Em branco)"))
    columns = order or labels
    pivot = pivot.reindex(columns=columns, fill_value=0)
    pivot["Total"] = pivot.sum(axis=1)
    all_rows = pivot.copy()
    pivot = pivot.sort_values("Total", ascending=False).head(limit)
    grand = {name: int(all_rows[name].sum()) for name in columns + ["Total"]}
    grand[row_field] = "Total"
    pivot = pivot.reset_index()
    pivot = pd.concat([pivot, pd.DataFrame([grand])], ignore_index=True)
    pivot = pivot.rename(columns={row_field: row_field.replace("_", " ")})
    return pivot


def aging_totals(df: pd.DataFrame, label: str, bucket_field: str) -> pd.DataFrame:
    row = {"STATUS ATUAL": label}
    counts = df[bucket_field].value_counts(dropna=False)
    row.update({bucket: int(counts.get(bucket, 0)) for bucket in BUCKET_LABELS})
    row["Total"] = len(df)
    return pd.DataFrame([row])


def footer_detail(df: pd.DataFrame) -> None:
    """Show only the CTE detail as a visible, scrollable native Streamlit dataframe."""
    columns = ["CTE", "PEDIDO", "STATUS_ATUAL", "SOLICITACAO", "DESCRICAO_ATUAL", "STATUS_TRATATIVA"]
    chosen = df.loc[:, columns].copy()
    chosen["SOLICITACAO"] = pd.to_datetime(chosen["SOLICITACAO"], errors="coerce").dt.strftime("%d/%m/%Y")

    with st.container(border=True):
        title_col, focus_col = st.columns([1, .1], gap="small", vertical_alignment="center")
        with title_col:
            st.markdown('<div class="panel-title">DETALHES DAS CTEs</div>', unsafe_allow_html=True)
        with focus_col:
            focus_clicked = st.button("⛶", key="focus_detail_cte", help="Abrir tabela geral em modo foco")
        if focus_clicked:
            show_table_focus("DETALHES DAS CTEs", chosen, native=True)

        st.dataframe(
            chosen,
            use_container_width=True,
            hide_index=True,
            height=280,
            column_config={
                "CTE": st.column_config.TextColumn("CTE", width="medium"),
                "PEDIDO": st.column_config.TextColumn("PEDIDO", width="medium"),
                "STATUS_ATUAL": st.column_config.TextColumn("STATUS_ATUAL", width="medium"),
                "SOLICITACAO": st.column_config.TextColumn("SOLICITACAO", width="small"),
                "DESCRICAO_ATUAL": st.column_config.TextColumn("DESCRICAO_ATUAL", width="medium"),
                "STATUS_TRATATIVA": st.column_config.TextColumn("STATUS_TRATATIVA", width="large"),
            },
        )
        st.caption(f"Total · {number_br(chosen['CTE'].notna().sum())}")


def backlog_page() -> None:
    selection = st.session_state.get("backlog_grupo")
    df = BASE.copy()
    if selection:
        df = df[df["TIPO"].eq(selection)]

    # First row mirrors the small Power BI cards and slicers.
    top = st.columns([1.05, 1.05, .78, .82, 1.18, 1.18, .86], gap="small", vertical_alignment="center")
    with top[0]:
        kpi("Soma de VALOR_NF", money_short(float(df["VALOR_NF"].sum())), "＄", "")
    with top[1]:
        kpi("Quantidade", count_short(len(df)), "▤", "")
    with top[2]:
        shipper_filter = filter_select("SHIPPER", SHIPPER_NAMES, "backlog_shipper")
    with top[3]:
        emitter_filter = filter_select("EMISSOR", EMITTER_NAMES, "backlog_emissor")
    with top[4]:
        request_period = period_picker(df, "SOLICITACAO", "DATA_SOLICITAÇÃO", "backlog_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    with top[5]:
        event_period = period_picker(df, "DATA_ATUAL", "DATA_EVENTO", "backlog_evento", default=(date(2024, 7, 10), date(2026, 9, 7)))
    with top[6]:
        status_filter = filter_select("STATUS_ATUAL", STATUS_NAMES, "backlog_status")

    df = apply_period(df, "SOLICITACAO", request_period)
    df = apply_period(df, "DATA_ATUAL", event_period)
    df = filtered(df, {"SHIPPER": shipper_filter, "EMISSOR": emitter_filter, "STATUS_ATUAL": status_filter})

    left, mid, right = st.columns([1.0, 1.0, .9], gap="small")
    with left:
        emitters = df[df["TIPO"].eq("Lojas e PAs")]
        table_panel("LOJAS E PA'S", summary_table(emitters, "EMISSOR", limit=12), height=224, dark_header=True, bar_column="BARRA", numeric_columns=("QTDE",), total_note=number_br(emitters.shape[0]))
    with mid:
        clients = df[df["TIPO"].eq("Cliente")]
        table_panel("CLIENTES", summary_table(clients, "SHIPPER", limit=12), height=224, dark_header=True, bar_column="BARRA", numeric_columns=("QTDE",), total_note=number_br(clients.shape[0]))
    with right:
        table_panel("STATUS", summary_table(df, "CIDADE_ORIGEM", limit=12), height=224, dark_header=True, numeric_columns=("QTDE",), total_note=number_br(df.shape[0]))

    dates = pd.to_datetime(df["DATA_ATUAL"], errors="coerce")
    trend = df.assign(DT=dates.dt.floor("D")).groupby("DT").size().rename("QTD").reset_index()
    trend = trend.sort_values("DT")
    with st.container(border=True):
        st.markdown('<div class="panel-title">DT_EVENTO X QTD</div>', unsafe_allow_html=True)
        if trend.empty:
            st.info("Sem eventos para o período selecionado.")
        else:
            trend["ROTULO"] = trend["QTD"].where(trend["QTD"] >= trend["QTD"].quantile(.94), "")
            chart_end = trend["DT"].max()
            chart_start = max(trend["DT"].min(), chart_end - pd.DateOffset(months=8))
            if chart_start >= chart_end:
                chart_start = chart_end - pd.DateOffset(days=1)
            fig = px.bar(trend, x="DT", y="QTD", text="ROTULO", color_discrete_sequence=[RED])
            fig.update_traces(textposition="outside", cliponaxis=False, marker_line_width=0)
            fig.update_layout(
                height=250, margin=dict(l=5, r=8, t=7, b=3), plot_bgcolor="#fff", paper_bgcolor="#fff",
                showlegend=False, font=dict(family="Roboto,Arial", size=8, color="#555"),
                xaxis=dict(
                    title=None, showgrid=True, gridcolor="#eee", tickfont=dict(size=8), nticks=12,
                    range=[chart_start, chart_end],
                    rangeslider=dict(visible=True, thickness=0.11, bgcolor="#f2f4f7", bordercolor="#dfe3ea", borderwidth=1),
                ),
                yaxis=dict(title=None, showgrid=True, gridcolor="#e7e7e7", tickfont=dict(size=8)),
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key="backlog_trend")
    footer_detail(df if not df.empty else BASE)


def top_common_controls(df: pd.DataFrame, key: str, *, include_aging: bool = False) -> pd.DataFrame:
    cols = st.columns([1, 1, .75, .75, 1] if include_aging else [1, 1, 1, .9], gap="small", vertical_alignment="center")
    selected = df.copy()
    # KPI widgets sit first, with date range slicers occupying the right half.
    metric_cards(cols[:3] if include_aging else cols[:2], selected, show_open=include_aging)
    if include_aging:
        with cols[3]:
            period = period_picker(selected, "DATA_ATUAL", "DATA ATUAL", f"{key}_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
        with cols[4]:
            sol = period_picker(selected, "SOLICITACAO", "DATA SOLICITAÇÃO", f"{key}_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    else:
        with cols[2]:
            period = period_picker(selected, "DATA_ATUAL", "DATA ATUAL", f"{key}_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
        with cols[3]:
            sol = period_picker(selected, "SOLICITACAO", "DATA SOLICITAÇÃO", f"{key}_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    selected = apply_period(selected, "DATA_ATUAL", period)
    selected = apply_period(selected, "SOLICITACAO", sol)
    return selected


def entrada_page() -> None:
    top = st.columns([1, 1, 1, 1.05, 1.05, .62], gap="small", vertical_alignment="center")
    metric_cards(top[:3], BASE)
    with top[3]:
        at_period = period_picker(BASE, "DATA_ATUAL", "DATA ATUAL", "entrada_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
    with top[4]:
        sol_period = period_picker(BASE, "SOLICITACAO", "DATA SOLICITAÇÃO", "entrada_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    with top[5]:
        finalized = st.multiselect("FINALIZADO?", ["EM ABERTO", "FINALIZADO"], key="entrada_final", placeholder="Todos")
    selected = apply_period(BASE, "DATA_ATUAL", at_period)
    selected = apply_period(selected, "SOLICITACAO", sol_period)
    selected = filtered(selected, {"FINALIZADO?": finalized})

    a, b = st.columns([1.45, 1.0], gap="small")
    with a:
        manifest_blank = selected["MANIFESTO"].isna() | selected["MANIFESTO"].eq("")
        city_not_closed = (
            selected["CIDADE_ORIGEM"]
            .fillna("")
            .str.strip()
            .str.upper()
            .ne("ENCERRADO")
        )
        cli = selected[
            selected["TIPO"].eq("Cliente")
            & selected["STATUS_ATUAL"].isin(
                ["EM DEVOLUCAO - ENTRADA", "ENTRADA"]
            )
            & manifest_blank
            & city_not_closed
        ].copy()
        cli["AgingEntradaPBIX"] = powerbi_aging_bucket(
            cli["DATA_ATUAL"], as_of=PBIX_ENTRY_AGE_AS_OF
        )
        client_shipper_count = int(cli["SHIPPER"].nunique(dropna=False))
        client_matrix = matrix_table(
            cli,
            "SHIPPER",
            "AgingEntradaPBIX",
            limit=max(11, client_shipper_count),
            bucket_labels=PBIX_ENTRY_AGE_BUCKETS,
        )
        # Power BI hides empty age buckets in this matrix (D+0 and D+2 here).
        empty_buckets = [
            bucket
            for bucket in PBIX_ENTRY_AGE_BUCKETS
            if bucket in client_matrix.columns
            and not client_matrix[bucket].astype(int).sum()
        ]
        client_matrix = client_matrix.drop(columns=empty_buckets)
        client_numeric_columns = tuple(
            [bucket for bucket in PBIX_ENTRY_AGE_BUCKETS if bucket in client_matrix.columns]
            + ["Total"]
        )
        table_panel(
            "AGING ENTRADA - CLIENTE",
            client_matrix,
            height=236,
            numeric_columns=client_numeric_columns,
        )
    with b:
        cust = selected[
            selected["STATUS_ATUAL"].isin(["CUSTODIA", "DEVOLUCAO RECUSADA"])
            & manifest_blank
            & city_not_closed
        ].copy()
        cust["AgingEntradaPBIX"] = powerbi_aging_bucket(
            cust["DATA_ATUAL"], as_of=PBIX_ENTRY_AGE_AS_OF
        )
        custody_description_count = int(cust["DESCRICAO_ATUAL"].nunique(dropna=False))
        custody_matrix = matrix_table(
            cust,
            "DESCRICAO_ATUAL",
            "AgingEntradaPBIX",
            limit=max(9, custody_description_count),
            bucket_labels=PBIX_ENTRY_AGE_BUCKETS,
        )
        # Hide empty buckets as Power BI does for this visual.
        empty_custody_buckets = [
            bucket
            for bucket in PBIX_ENTRY_AGE_BUCKETS
            if bucket in custody_matrix.columns
            and not custody_matrix[bucket].astype(int).sum()
        ]
        custody_matrix = custody_matrix.drop(columns=empty_custody_buckets)
        custody_numeric_columns = tuple(
            [bucket for bucket in PBIX_ENTRY_AGE_BUCKETS if bucket in custody_matrix.columns]
            + ["Total"]
        )
        table_panel(
            "AGING - CUSTODIA",
            custody_matrix,
            height=236,
            numeric_columns=custody_numeric_columns,
        )

    c, d = st.columns([1.45, 1.0], gap="small")
    with c:
        stores = selected[
            selected["TIPO"].eq("Lojas e PAs")
            & selected["STATUS_ATUAL"].isin(
                ["EM DEVOLUCAO - ENTRADA", "ENTRADA"]
            )
            & manifest_blank
            & city_not_closed
        ].copy()
        stores["AgingEntradaPBIX"] = powerbi_aging_bucket(
            stores["DATA_ATUAL"], as_of=PBIX_ENTRY_AGE_AS_OF
        )
        store_emitter_count = int(stores["EMISSOR"].nunique(dropna=False))
        stores_matrix = matrix_table(
            stores,
            "EMISSOR",
            "AgingEntradaPBIX",
            limit=max(12, store_emitter_count),
            bucket_labels=PBIX_ENTRY_AGE_BUCKETS,
        )
        empty_store_buckets = [
            bucket
            for bucket in PBIX_ENTRY_AGE_BUCKETS
            if bucket in stores_matrix.columns
            and not stores_matrix[bucket].astype(int).sum()
        ]
        stores_matrix = stores_matrix.drop(columns=empty_store_buckets)
        stores_numeric_columns = tuple(
            [bucket for bucket in PBIX_ENTRY_AGE_BUCKETS if bucket in stores_matrix.columns]
            + ["Total"]
        )
        table_panel(
            "AGING ENTRADA - LOJAS E PA'S",
            stores_matrix,
            height=212,
            numeric_columns=stores_numeric_columns,
        )
    with d:
        stuck_records = selected[
            selected["STATUS_ATUAL"].eq("TRAVADO")
        ].copy()
        closed_records = selected[
            selected["CIDADE_ORIGEM"].eq("ENCERRADO")
        ].copy()
        stuck_records["AgingEntradaPBIX"] = powerbi_aging_bucket(
            stuck_records["DATA_ATUAL"], as_of=PBIX_ENTRY_AGE_AS_OF
        )
        closed_records["AgingEntradaPBIX"] = powerbi_aging_bucket(
            closed_records["DATA_ATUAL"], as_of=PBIX_ENTRY_AGE_AS_OF
        )
        def status_aging_row(label: str, records: pd.DataFrame) -> dict[str, object]:
            counts = records["AgingEntradaPBIX"].value_counts(dropna=False)
            return {
                "GRUPO": label,
                **{bucket: int(counts.get(bucket, 0)) for bucket in PBIX_ENTRY_AGE_BUCKETS},
                "Total": int(records["CTE"].notna().sum()),
            }

        status_stuck = pd.DataFrame([status_aging_row("Total", stuck_records)])
        status_closed = pd.DataFrame([status_aging_row("Total", closed_records)])
        empty_status_buckets = [
            bucket
            for bucket in PBIX_ENTRY_AGE_BUCKETS
            if not status_stuck[bucket].astype(int).sum()
            and not status_closed[bucket].astype(int).sum()
        ]
        status_stuck = status_stuck.drop(columns=empty_status_buckets)
        status_closed = status_closed.drop(columns=empty_status_buckets)
        status_numeric_columns = tuple(
            [bucket for bucket in PBIX_ENTRY_AGE_BUCKETS if bucket in status_stuck.columns]
            + ["Total"]
        )
        table_panel(
            "AGING - TRAVADO",
            status_stuck,
            height=66,
            numeric_columns=status_numeric_columns,
        )
        table_panel(
            "AGING - ENCERRADO",
            status_closed,
            height=66,
            numeric_columns=status_numeric_columns,
        )
    footer_detail(selected if not selected.empty else BASE)


def control_clients_page() -> None:
    top = st.columns([1, 1, 1, 1.1, 1.1, .65], gap="small", vertical_alignment="center")
    metric_cards(top[:3], CONTROL_CLIENTS)
    with top[3]:
        data_period = period_picker(CONTROL_CLIENTS, "DATA_ATUAL", "DATA ATUAL", "cc_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
    with top[4]:
        sol_period = period_picker(CONTROL_CLIENTS, "SOLICITACAO", "DATA SOLICITAÇÃO", "cc_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    with top[5]:
        finalized = st.multiselect("FINALIZADO?", ["EM ABERTO", "FINALIZADO"], key="cc_final", placeholder="Todos")
    df = apply_period(CONTROL_CLIENTS, "DATA_ATUAL", data_period)
    df = apply_period(df, "SOLICITACAO", sol_period)
    df = filtered(df, {"FINALIZADO?": finalized})

    aging_cols = st.columns([1, 1.05, 1.25, 1], gap="small", vertical_alignment="center")
    with aging_cols[0]:
        aging = st.selectbox("AgingExpedição", ["Todos", *AGE_BUCKETS], key="cc_aging", label_visibility="visible")
    with aging_cols[1]:
        extra_period = period_picker(CONTROL_CLIENTS, "DATA_ATUAL", "AGING EXPEDIÇÃO", "cc_aging_dates", default=(date(2026, 6, 17), date(2026, 9, 8)))
    with aging_cols[2]:
        status_filter = st.multiselect("STATUS", STATUS_NAMES, key="cc_status", placeholder="Todos")
    with aging_cols[3]:
        shipper_filter = st.multiselect("SHIPPER", SHIPPER_NAMES, key="cc_shipper", placeholder="Todos")
    if aging != "Todos":
        df = df[df["AgingExpedição"].eq(aging)]
    df = apply_period(df, "DATA_ATUAL", extra_period)
    df = filtered(df, {"STATUS_ATUAL": status_filter, "SHIPPER": shipper_filter})

    left, right = st.columns([1.15, 1], gap="small")
    with left:
        wait = df[df["STATUS_TRATATIVA"].eq("AGUARDANDO AGENDAMENTO")]
        table_panel("AGUARDANDO AGENDAMENTO", matrix_table(wait, "SHIPPER", "AgingExpedição", limit=11), height=206, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
        shipped = df[df["STATUS_TRATATIVA"].eq("EXPEDIDO")]
        table_panel("EXPEDIDO", matrix_table(shipped, "SHIPPER", "AgingExpedição", limit=11), height=198, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    with right:
        status_matrix = matrix_table(df, "STATUS_ATUAL", "AgingExpedição", limit=8)
        table_panel("STATUS", status_matrix, height=170, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
        scheduled = df[df["STATUS_TRATATIVA"].eq("AGENDADO")]
        table_panel("AGENDADO", matrix_table(scheduled, "SHIPPER", "AgingExpedição", limit=5), height=134, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
        table_panel("CONSOLIDADO", matrix_table(df, "CIDADE_ORIGEM", "AgingExpedição", limit=5), height=134, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    footer_detail(df if not df.empty else CONTROL_CLIENTS)


def control_stores_page() -> None:
    top = st.columns([1, 1, 1, 1.1, 1.1, .65], gap="small", vertical_alignment="center")
    metric_cards(top[:3], CONTROL_STORES)
    with top[3]:
        data_period = period_picker(CONTROL_STORES, "DATA_ATUAL", "DATA ATUAL", "cl_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
    with top[4]:
        sol_period = period_picker(CONTROL_STORES, "SOLICITACAO", "DATA SOLICITAÇÃO", "cl_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    with top[5]:
        finalized = st.multiselect("FINALIZADO?", ["EM ABERTO", "FINALIZADO"], key="cl_final", placeholder="Todos")
    df = apply_period(CONTROL_STORES, "DATA_ATUAL", data_period)
    df = apply_period(df, "SOLICITACAO", sol_period)
    df = filtered(df, {"FINALIZADO?": finalized})
    a, b = st.columns([1.25, 1.0], gap="small")
    with a:
        waiting = df[df["STATUS_TRATATIVA"].eq("LOJAS -TRANSF")]
        table_panel("AGUARDANDO RETIRADA", matrix_table(waiting, "EMISSOR", "AgingEntrada", limit=27), height=477, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    with b:
        shipped = df[df["STATUS_TRATATIVA"].eq("EXPEDIDO")]
        table_panel("EXPEDIDO", matrix_table(shipped, "EMISSOR", "AgingExpedição", limit=25), height=477, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    footer_detail(df if not df.empty else CONTROL_STORES)


def faixa_valor_page() -> None:
    top = st.columns([1, 1, 1, 1.1, 1.1, .65], gap="small", vertical_alignment="center")
    metric_cards(top[:3], BASE)
    with top[3]:
        date_filter = period_picker(BASE, "DATA_ATUAL", "DATA ATUAL", "fv_data")
    with top[4]:
        shipper = st.multiselect("SHIPPER", SHIPPER_NAMES, key="fv_shipper", placeholder="Todos")
    with top[5]:
        tipo = st.multiselect("TIPO", ["Cliente", "Lojas e PAs", "Rua"], key="fv_tipo", placeholder="Todos")
    df = apply_period(BASE.copy(), "DATA_ATUAL", date_filter)
    df = filtered(df, {"SHIPPER": shipper, "TIPO": tipo})
    a, b, c = st.columns([1.35, .95, .95], gap="small")
    with a:
        table_panel("FAIXA DE VALOR", summary_table(df, "Faixa_Valor", limit=8), height=225, dark_header=True, numeric_columns=("QTDE",))
    with b:
        table_panel("CLIENTES", summary_table(df[df["TIPO"].eq("Cliente")], "SHIPPER", limit=10), height=225, dark_header=True, numeric_columns=("QTDE",))
    with c:
        table_panel("LOJAS E PA'S", summary_table(df[df["TIPO"].eq("Lojas e PAs")], "EMISSOR", limit=10), height=225, dark_header=True, numeric_columns=("QTDE",))
    d, e = st.columns([1.15, 1], gap="small")
    with d:
        table_panel("STATUS", matrix_table(df, "STATUS_ATUAL", "AgingEntrada", limit=9), height=235, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    with e:
        table_panel("OCORRÊNCIA", matrix_table(df, "DESCRICAO_ATUAL", "AgingEntrada", limit=9), height=235, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    footer_detail(df if not df.empty else BASE)


def lista_cte_page() -> None:
    sidebar, main = st.columns([.78, 2.2], gap="small")
    with sidebar:
        with st.container(border=True):
            st.markdown('<div class="panel-title" style="font-size:17px;text-align:center">FILTROS</div>', unsafe_allow_html=True)
            df = BASE.copy()
            dates_sol = pd.to_datetime(df["SOLICITACAO"])
            dates_emi = pd.to_datetime(df["DT_EMISSAO"])
            p1 = st.date_input("DATA SOLICITAÇÃO", value=(date(2026, 7, 28), date(2026, 8, 19)), min_value=dates_sol.min().date(), max_value=dates_sol.max().date(), format="DD/MM/YYYY", key="lista_sol")
            p2 = st.date_input("DATA EMISSÃO", value=(date(2026, 6, 16), date(2026, 8, 30)), min_value=dates_emi.min().date(), max_value=dates_emi.max().date(), format="DD/MM/YYYY", key="lista_emissao")
            for field, label, choices in [
                ("SHIPPER", "CLIENTE", SHIPPER_NAMES), ("EMISSOR", "LOJAS E PA'S", EMITTER_NAMES),
                ("STATUS_TRATATIVA", "STATUS_TT", category_options("STATUS_TRATATIVA")),
                ("TIPO", "TIPO", category_options("TIPO")),
                ("MANIFESTO", "MANIFESTO", ["(Em branco)", *sorted(BASE["MANIFESTO"].dropna().astype(str).unique())[:30]]),
                ("PROXIMA_UNIDADE", "PROXIMA_UNIDADE", UNIT_NAMES), ("Faixa_Valor", "Faixa_Valor", VALUE_BANDS),
            ]:
                selected = st.multiselect(label, choices, key=f"lista_{field}", placeholder="Pesquisar")
                if selected:
                    if field == "MANIFESTO" and "(Em branco)" in selected:
                        df = df[df[field].isna() | df[field].astype(str).isin(selected)]
                    else:
                        df = df[df[field].astype(str).isin(selected)]
            search = st.text_input("PESQUISAR CTE / PEDIDO", key="lista_busca", placeholder="Pesquisar")
            if search.strip():
                mask = df["CTE"].astype(str).str.contains(search, case=False) | df["PEDIDO"].astype(str).str.contains(search, case=False)
                df = df[mask]
            if isinstance(p1, tuple) and len(p1) == 2:
                df = apply_period(df, "SOLICITACAO", p1)
            if isinstance(p2, tuple) and len(p2) == 2:
                df = apply_period(df, "DT_EMISSAO", p2)
    with main:
        with st.container(border=True):
            st.markdown(f'<div class="panel-title">RESULTADO DA CONSULTA · {number_br(len(df))} REGISTROS</div>', unsafe_allow_html=True)
            view = df.loc[:, ["CTE", "PEDIDO", "STATUS_ATUAL", "SOLICITACAO", "DESCRICAO_ATUAL", "STATUS_TRATATIVA", "EMISSOR", "SOLICITANTE", "PROXIMA_UNIDADE", "SHIPPER", "MANIFESTO", "VALOR_NF"]].copy()
            view["SOLICITACAO"] = pd.to_datetime(view["SOLICITACAO"]).dt.strftime("%d/%m/%Y")
            view["VALOR_NF"] = view["VALOR_NF"].map(lambda val: "—" if pd.isna(val) else f"R$ {decimal_br(val)}")
            st.dataframe(view, use_container_width=True, hide_index=True, height=535)
with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=148)
    st.caption("REVERSAS · TC JAD-SAO")
    st.markdown("### Páginas")
    selected_page = st.radio("Navegação do relatório", PAGES, index=0, label_visibility="collapsed", key="report_page")
    st.caption("Fonte: data.xlsx")

header(selected_page, backlog_type=(selected_page == "BACKLOG"))
if selected_page == "BACKLOG":
    backlog_page()
elif selected_page == "ENTRADA":
    entrada_page()
elif selected_page == "CONTROLE CLIENTES":
    control_clients_page()
elif selected_page == "CONTROLE LOJAS":
    control_stores_page()
elif selected_page == "FAIXA DE VALOR":
    faixa_valor_page()
else:
    lista_cte_page()
