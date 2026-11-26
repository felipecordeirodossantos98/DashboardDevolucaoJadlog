"""Painel Streamlit Devolução - TCv2 com base padrão e importação Excel."""

from __future__ import annotations

import hashlib
import html
import base64
from io import BytesIO
from datetime import date
from pathlib import Path
import unicodedata
import uuid

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook

from data import (
    AGE_BUCKETS,
    PBIX_ENTRY_AGE_BUCKETS,
    PBIX_ENTRY_AGE_AS_OF,
    VALUE_BANDS,
    load_source_data,
    powerbi_aging_bucket,
)


ROOT = Path(__file__).resolve().parent
lazy_export_table = components.declare_component(
    "jadlog_lazy_export_table",
    path=str(ROOT / "components" / "lazy_export_table"),
)
LOGO_PATH = ROOT / "images" / "jadlog_logo.webp"
ICON_PATH = ROOT / "images" / "jadlog_icon.webp"
RED = "#e5003d"
VENCIMENTO_FILTER_OPTIONS = ["No prazo", "Vencido"]
PAGES = ["BACKLOG", "ENTRADA", "CONTROLE LOJAS", "CONTROLE CLIENTES", "FAIXA DE VALOR", "Lista CTE"]
PAGE_FILTER_KEYS = {
    "BACKLOG": [
        "backlog_emissor", "backlog_shipper", "backlog_status", "backlog_vencimento",
        "backlog_sol", "backlog_evento", "backlog_emissao",
    ],
    "ENTRADA": ["entrada_aging", "entrada_vencimento", "entrada_data", "entrada_sol"],
    "CONTROLE CLIENTES": [
        "cc_aging_filter_v3", "cc_status_v2", "cc_shipper_v2", "cc_data_v2", "cc_sol_v2",
    ],
    "CONTROLE LOJAS": [
        "cl_data", "cl_sol", "cl_aging", "cl_status", "cl_emissor",
    ],
    "FAIXA DE VALOR": [
        "fv_faixa_valor", "fv_emissor", "fv_shipper", "fv_tipo", "fv_vencimento", "fv_data",
    ],
    "Lista CTE": [
        "lista_vencimento", "lista_SHIPPER", "lista_EMISSOR", "lista_STATUS_ATUAL",
        "lista_STATUS_TRATATIVA",
        "lista_TIPO", "lista_MANIFESTO", "lista_PROXIMA_UNIDADE", "lista_Faixa_Valor",
        "lista_sol", "lista_emissao", "lista_busca",
    ],
}
CLIENT_ALLOWLIST = [
    'MERCADO LIVRE',
    'AMAZON',
    'SHEIN',
    'RENNER',
    'F.M TRANSPORTES',
    'AMAZON-PARP',
    'O BOTICARIO',
    'SHOPEE',
    'PHOENEX CARGO',
    'CARREFOUR',
    'SHEIN IN GLOW',
    'GURGEL MIX',
    'LOREAL',
    'TELEFONICA',
    'CAMICADO',
    'OLIST',
    'COBASI COMERCIO DE PRODUTOS BASICOS E INDUSTRIALIZADOS S.A.',
    'VIA VAREJO',
    'OBJETIVO 1220 COMERCIO DE COSMETICOS LTDA',
    'EDITORA JUSPODIVM',
    'TRAMONTINA SUDESTE',
    'HOPE',
    'MULTILASER',
    'ON RUNNING',
    'AMERICANAS',
    'UNIDOCKS ASSESSORIA E LOGISTICA DE MATERIAIS',
    'NATURA&',
]
EMISSOR_ALLOWLIST = [
    'LJ BARUERI 04',
    'PA OSASCO 01',
    'TC EMISSAO TECA',
    'LJ SANTO ANDRE 02',
    'PA OSASCO 03',
    'PA JANDIRA 01',
    'PA EXTREMA 01',
    'LJ SAO PAULO 22',
    'PA EXTREMA 07',
    'LJ SAO PAULO 20',
    'LJ ITAPEVI 01',
    'PA SAO PAULO 21',
    'PA BARUERI 02',
    'LJ SAO PAULO 26',
    'LJ SAO PAULO 03',
    'PA EXTREMA 06',
    'PA SAO PAULO 22',
    'LJ OSASCO 01',
    'LJ SAO PAULO 02',
    'LJ SAO PAULO 12',
    'LJ BARUERI 03',
    'LJ SAO PAULO 35',
    'LJ SAO BERNARDO DO CAMPO 02',
    'LJ GUARULHOS 03',
    'LJ SAO PAULO 10',
    'TC DEPOT 1 - FREGUESIA',
    'LJ SAO PAULO 01',
    'LJ SAO PAULO 29',
    'LJ SAO PAULO 07',
    'LJ SAO PAULO 04',
    'LJ SAO PAULO 16',
    'PA SAO PAULO 30',
    'LJ SAO PAULO 18',
    'LJ TABOAO DA SERRA 01',
    'PA SAO PAULO 07',
    'PA SAO PAULO 02',
    'PA SAO PAULO 05',
    'PA BARUERI 09',
    'LJ EMBU DAS ARTES 01',
    'LJ SAO PAULO 13',
    'LJ SAO PAULO 09',
    'PA EXTREMA 08',
    'PA BARUERI 07',
    'LJ SAO PAULO 14',
    'LJ JUNDIAI 01',
    'LJ SAO PAULO 06',
    'CO RIO DE JANEIRO 07',
    'LJ SAO PAULO',
    'CO SAMAMBAIA 01',
    'LJ SAO CAETANO DO SUL 01',
    'LJ SAO PAULO 17',
    'PA BELO HORIZONTE 01',
    'PA-JAD-PICKUP-CWB',
    'CO FLORIANOPOLIS 01',
    'CO RIO DE JANEIRO 01',
    'LJ BELO HORIZONTE 03',
    'CO JUIZ DE FORA 02',
    'CO MARINGA 02',
    'PA SAO PAULO 26',
    'CO NITEROI 02',
    'CO RIO DE JANEIRO 06',
    'CO SAO PAULO 12',
    'CO VILA VELHA 02',
    'CO BETIM 02',
    'CO DUQUE DE CAXIAS 01',
    'CO RIO DE JANEIRO 04',
    'CO RIO DE JANEIRO 05',
    'CO BEBEDOURO 01',
    'CO CONSELHEIRO LAFAIETE 02',
    'CO GUAPIMIRIM 02',
    'CO ITABORAI 01',
    'CO ITAPERUNA 02',
    'CO JOAO PESSOA 02 FL',
    'CO JOINVILLE 02',
    'CO NOVO HAMBURGO 02',
    'CO PAULINIA 02',
    'CO QUEIMADOS',
    'CO RECIFE 01',
    'CO SANTOS 02',
    'CO SAO CARLOS 01',
    'CO SAO GONCALO 01',
    'CO SAO PAULO 16',
    'CO SERRA 01',
    'CO TAGUATINGA 01',
    'CO TAUBATE 01',
    'CO UBERLANDIA 03',
    'LJ BELO HORIZONTE 04',
    'LJ BELO HORIZONTE 05',
    'LJ SAO PAULO 11',
    'PA SAO PAULO 10',
    'CO ARACATUBA 01',
    'CO ARARAQUARA 02',
    'CO ASSIS 03',
    'CO BELO HORIZONTE 02',
    'CO BOTUCATU 02',
    'CO BRAGANCA PAULISTA 02',
    'CO BRASILIA 01',
    'CO CACHOEIRO DE ITAPEMIRIM 02',
    'CO CAMPO MOURAO 01',
    'CO CAXIAS DO SUL 03',
    'CO GOV VALADARES 01',
    'CO GURUPI 02',
    'CO INDAIATUBA 02',
    'CO ITAJAI 02',
    'CO ITAJUBA 01',
    'CO JABOATAO DOS GUARARAPES 03',
    'CO JAD SALVADOR',
    'CO LAGES 01',
    'CO LENCOIS PAULISTA 01',
    'CO MAR DE ESPANHA 01',
    'CO MONTE MOR 02',
    'CO MONTE SIAO 01',
    'CO PASSO FUNDO 02',
    'CO PETROPOLIS 02',
    'CO PIRASSUNUNGA 01',
    'CO PORTO ALEGRE 03',
    'CO RIO CLARO 02',
    'CO RIO DAS OSTRAS 01',
    'CO RIO DE JANEIRO 03',
    'CO RIO DE JANEIRO 08',
    'CO RIO DE JANEIRO 13',
    'CO SANTA MARIA',
    'CO SAO JOAO DA BOA VISTA 01',
    'CO SAO JOAO DE MERITI 01',
    'CO SAO JOSE DOS CAMPOS 01',
    'CO SAO PAULO 01',
    'CO SOROCABA 01',
    'CO UBA 01',
    'CO VALINHOS 02',
    'CO VICOSA 02',
    'CO VINHEDO 01',
    'CO VOLTA REDONDA 01',
    'LJ BLUMENAU 01',
    'LJ CARUARU 01',
    'LJ CHAPECO 02',
    'LJ GUARULHOS 02',
    'LJ SAO JOSE DOS CAMPOS 02',
    'PA CARIACICA 01',
    'PA EXTREMA 02',
]
BUCKET_LABELS = [bucket for bucket in AGE_BUCKETS if bucket != "Sem dados"]

st.set_page_config(
    page_title="Devolução | TC Reversas Jad-SAO",
    page_icon=str(ICON_PATH),
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_data(show_spinner=False, max_entries=4)
def load_dashboard_data(
    source_fingerprint: str, _file_bytes: bytes, _source_name: str
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, tuple[str, ...]]:
    # The fingerprint keys the cache; large workbook bytes are deliberately
    # excluded from Streamlit's per-rerun argument hashing.
    del source_fingerprint
    base, clients, stores = load_source_data(_file_bytes)
    original_columns = tuple(base.attrs.get("original_columns", ()))
    if not original_columns:
        raise ValueError("Não foi possível identificar as colunas originais da planilha.")
    return base, clients, stores, original_columns


with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=148)
    st.caption("REVERSAS · TC JAD-SAO")
    st.markdown("### Importar base")
    uploaded_file = st.file_uploader(
        "Arquivo Excel (.xlsx ou .xlsm)",
        type=["xlsx", "xlsm"],
        key="dashboard_source_upload_v2",
    )

if uploaded_file is None:
    st.info("Importe uma base para continuar.")
    st.stop()

source_bytes = uploaded_file.getvalue()
active_source = uploaded_file.name
upload_identity = (
    getattr(uploaded_file, "file_id", None),
    uploaded_file.name,
    uploaded_file.size,
)
if (
    upload_identity[0] is not None
    and st.session_state.get("_dashboard_upload_identity") == upload_identity
    and st.session_state.get("_dashboard_source_fingerprint")
):
    source_fingerprint = st.session_state["_dashboard_source_fingerprint"]
else:
    source_fingerprint = hashlib.sha256(source_bytes).hexdigest()
    st.session_state["_dashboard_upload_identity"] = upload_identity
if st.session_state.get("_dashboard_source_fingerprint") != source_fingerprint:
    filter_prefixes = (
        "backlog_", "entrada_", "cc_", "cl_", "fv_", "lista_", "focus_",
        "_lazy_export_pending_", "_multiselect_cleared_", "_preferred_period_",
    )
    for state_key in list(st.session_state):
        if state_key.startswith(filter_prefixes):
            del st.session_state[state_key]
    st.session_state["_dashboard_source_fingerprint"] = source_fingerprint

try:
    (
        _RAW_BASE,
        _RAW_CONTROL_CLIENTS,
        _RAW_CONTROL_STORES,
        _RAW_SOURCE_COLUMNS,
    ) = load_dashboard_data(source_fingerprint, source_bytes, active_source)
except Exception as error:
    st.error(f"Não foi possível importar {active_source}: {error}")
    st.stop()

DETAIL_SOURCE_COLUMNS = list(_RAW_SOURCE_COLUMNS)


def category_options(column: str, frame: pd.DataFrame | None = None) -> list[str]:
    source = BASE if frame is None else frame
    return sorted(source[column].dropna().astype(str).unique().tolist())


def normalize_dimension_name(value: object) -> str:
    if pd.isna(value):
        return ""
    decomposed = unicodedata.normalize("NFKD", str(value)).casefold()
    unaccented = "".join(char for char in decomposed if not unicodedata.combining(char))
    return " ".join(unaccented.split())


CLIENT_ALLOWLIST_KEYS = {normalize_dimension_name(name) for name in CLIENT_ALLOWLIST}
EMISSOR_ALLOWLIST_KEYS = {normalize_dimension_name(name) for name in EMISSOR_ALLOWLIST}
CLASSIFICATION_RULES_FINGERPRINT = hashlib.sha256(
    "\x1f".join(CLIENT_ALLOWLIST + EMISSOR_ALLOWLIST).encode("utf-8")
).hexdigest()


def allowed_client_mask(frame: pd.DataFrame) -> pd.Series:
    return frame["SHIPPER"].map(normalize_dimension_name).isin(CLIENT_ALLOWLIST_KEYS)


def allowed_emitter_mask(frame: pd.DataFrame) -> pd.Series:
    values = frame["EMISSOR"].map(normalize_dimension_name)
    matches = values.isin(EMISSOR_ALLOWLIST_KEYS)
    if "(em branco)" in EMISSOR_ALLOWLIST_KEYS:
        matches |= values.eq("")
    return matches


@st.cache_data(show_spinner=False, max_entries=8)
def prepare_dashboard_frames(
    classification_fingerprint: str,
    assignment_pairs: tuple[tuple[str, str], ...],
    _raw_base: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Cache allowlist classification and labels for this source and rule set."""
    del classification_fingerprint
    base = _raw_base.copy()
    client_names = base["SHIPPER"].map(normalize_dimension_name)
    emitter_names = base["EMISSOR"].map(normalize_dimension_name)
    client_rows = client_names.isin(CLIENT_ALLOWLIST_KEYS)
    emitter_rows = emitter_names.isin(EMISSOR_ALLOWLIST_KEYS)

    base["TIPO"] = "Rua"
    base.loc[emitter_rows, "TIPO"] = "Lojas e PAs"
    base.loc[client_rows, "TIPO"] = "Cliente"
    conflicts = client_rows & emitter_rows
    base.loc[conflicts, "TIPO"] = "Conflito"
    pair_keys = client_names + "\x1f" + emitter_names
    destinations = pair_keys.map(dict(assignment_pairs)).fillna("CLIENTES")
    base.loc[conflicts & destinations.eq("CLIENTES"), "TIPO"] = "Cliente"
    base.loc[conflicts & destinations.eq("LOJAS E PA'S"), "TIPO"] = "Lojas e PAs"
    base.loc[~base["TIPO"].isin(["Cliente", "Lojas e PAs"]), "TIPO"] = "Rua"

    client_labels = {normalize_dimension_name(name): name for name in CLIENT_ALLOWLIST}
    emitter_labels = {
        normalize_dimension_name(name): name for name in EMISSOR_ALLOWLIST
        if normalize_dimension_name(name) != "(em branco)"
    }
    base["SHIPPER"] = client_names.map(client_labels).fillna(base["SHIPPER"])
    base["EMISSOR"] = emitter_names.map(emitter_labels).fillna(base["EMISSOR"])
    blank_emitter = base["EMISSOR"].isna() | base["EMISSOR"].astype("string").str.strip().eq("")
    base.loc[blank_emitter, "EMISSOR"] = "(Em branco)"

    clients = base.loc[base["TIPO"].eq("Cliente")].copy()
    stores = base.loc[base["TIPO"].eq("Lojas e PAs")].copy()
    classified = base.loc[base["TIPO"].isin(["Cliente", "Lojas e PAs"])].copy()
    return base, clients, stores, classified


_conflict_pair_assignments = st.session_state.get("backlog_conflict_pair_assignments_v2", {})
_assignment_pairs = tuple(sorted(_conflict_pair_assignments.items()))
BASE, CONTROL_CLIENTS, CONTROL_STORES, CLASSIFIED_BASE = prepare_dashboard_frames(
    f"{source_fingerprint}:{CLASSIFICATION_RULES_FINGERPRINT}",
    _assignment_pairs,
    _RAW_BASE,
)
XLSX_CACHE_EPOCH = hashlib.sha256(
    f"{source_fingerprint}|{CLASSIFICATION_RULES_FINGERPRINT}|{pd.Timestamp.today().date()}|{_assignment_pairs}".encode()
).hexdigest()


def _backlog_category_masks(frame: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
    """Return exclusive client/store masks; overlaps are handled separately."""
    clients = allowed_client_mask(frame)
    emitters = allowed_emitter_mask(frame)
    stores = emitters & ~clients
    return clients, stores


def backlog_conflict_mask(frame: pd.DataFrame) -> pd.Series:
    """Rows whose SHIPPER and EMISSOR both match their fixed lists."""
    return allowed_client_mask(frame) & allowed_emitter_mask(frame)


def backlog_group_mask(frame: pd.DataFrame, group: str) -> pd.Series:
    """Return rows currently assigned to a fixed group."""
    if "TIPO" in frame.columns and group in {"Cliente", "Lojas e PAs", "Rua"}:
        return frame["TIPO"].eq(group)
    return pd.Series(True, index=frame.index, dtype=bool)


_present_clients = {
    normalize_dimension_name(name)
    for name in BASE["SHIPPER"].dropna().astype(str)
}
_present_emitters = {
    normalize_dimension_name(name)
    for name in BASE["EMISSOR"].dropna().astype(str)
}
CLIENT_FILTER_NAMES = sorted(
    name for name in CLIENT_ALLOWLIST
    if normalize_dimension_name(name) in _present_clients
)
EMISSOR_FILTER_NAMES = sorted(
    name for name in EMISSOR_ALLOWLIST
    if normalize_dimension_name(name) in _present_emitters
)
SHIPPER_NAMES = CLIENT_FILTER_NAMES
EMITTER_NAMES = EMISSOR_FILTER_NAMES
UNIT_NAMES = category_options("PROXIMA_UNIDADE", CLASSIFIED_BASE)
STATUS_NAMES = category_options("STATUS_ATUAL", CLASSIFIED_BASE)
BACKLOG_STATUS_NAMES = category_options("STATUS_ATUAL", BASE)
STATUS_DIMENSION = "CIDADE_ORIGEM" if BASE["CIDADE_ORIGEM"].notna().any() else "STATUS_ATUAL"
STATUS_DIMENSION_NAMES = sorted(set(category_options(STATUS_DIMENSION)))
if BASE[STATUS_DIMENSION].isna().any():
    STATUS_DIMENSION_NAMES = sorted([*STATUS_DIMENSION_NAMES, "(Em branco)"])

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: Roboto, Arial, sans-serif; color:#202939; }
    .stApp { background:#f5f6fa; }
    [data-testid="stHeader"] { background:transparent; }
    [data-testid="stToolbar"] { right:12px; }
    [data-testid="stSidebar"] { background:#fff; border-right:1px solid #e7e9ef; transition:width .46s cubic-bezier(.4,0,.2,1); }
    [data-testid="stSidebar"] > div { padding:20px 14px 14px; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { margin:0; }
    [data-testid="stSidebar"] h3 { font-size:12px; color:#8a91a1; font-weight:700; letter-spacing:.09em; margin:18px 10px 8px; text-transform:uppercase; }
    [data-testid="stSidebar"] div[role="radiogroup"] { gap:4px; }
    [data-testid="stSidebar"] div[role="radiogroup"] label { border-radius:9px; padding:8px 10px; min-height:39px; border:1px solid transparent; font-size:13px; transition:all .15s ease; }
    [data-testid="stSidebar"] div[role="radiogroup"] label:hover { background:#f7f8fb; }
    [data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child { display:none; }
    [data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) { background:#fff1f4; border-color:#ffe0e8; color:#c90036; font-weight:650; }
    [data-testid="stSidebar"] div[role="radiogroup"] label p { font-size:13px; }
    .block-container { width:100%; max-width:none; padding:1.35rem 1.55rem 1.4rem; }
    .report-header { display:grid; grid-template-columns:1fr 165px; align-items:center; gap:16px; width:calc(100% - 20px); box-sizing:border-box; min-height:76px; border:1px solid #eaecf1; border-radius:15px; background:#fff; padding:12px 18px; margin:0 20px 17px 0; box-shadow:0 2px 8px #20293908; }
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
    .st-key-page-filter-clear div[data-testid="stButton"] { justify-content:flex-end; }
    .st-key-page-filter-clear div[data-testid="stButton"] button { width:auto !important; min-width:118px !important; height:36px !important; min-height:36px !important; padding:0 12px !important; border:1px solid #e1e5ec !important; border-radius:9px !important; background:#fff !important; box-shadow:none !important; color:#566176 !important; font-size:11px !important; line-height:normal !important; }
    .st-key-page-filter-clear div[data-testid="stButton"] button:hover { border-color:#c90036 !important; color:#c90036 !important; }
    div[data-testid="stMetric"] { background:transparent; border:0; padding:0; min-height:0; }
    div[data-testid="stMetricLabel"] { color:#788295 !important; font-size:10px !important; }
    div[data-testid="stMetricValue"] { color:#202939; font-size:21px !important; font-weight:700; letter-spacing:-.02em; white-space:nowrap; }
    .kpi-card { box-sizing:border-box; display:flex; align-items:center; gap:10px; width:173px; min-width:173px; max-width:173px; height:84px; min-height:84px; padding:10px 11px; border:1px solid #d1d5dd; border-radius:9px; background:#fff; overflow:hidden; }
    .kpi-copy { flex:1; min-width:0; }
    .kpi-label { overflow:hidden; color:#30343b; font-size:10px; line-height:1.3; white-space:nowrap; text-overflow:ellipsis; }
    .kpi-value { overflow:hidden; color:#202939; font-size:16px; font-weight:700; line-height:1.2; letter-spacing:-.02em; white-space:nowrap; text-overflow:ellipsis; }
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
    @media (prefers-reduced-motion: reduce) { [data-testid="stSidebar"] { transition:none !important; } }
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


def dataframe_xlsx_cache_key(
    frame: pd.DataFrame, *, stable_source_rows: bool = False
) -> str:
    if stable_source_rows:
        data_hashes = pd.util.hash_pandas_object(frame.index, index=False).values.tobytes()
    else:
        data_hashes = pd.util.hash_pandas_object(frame, index=True).values.tobytes()
    key_material = b"\x1f".join(
        [
            XLSX_CACHE_EPOCH.encode("ascii"),
            "\x1f".join(str(column) for column in frame.columns).encode("utf-8"),
            data_hashes,
        ]
    )
    return hashlib.sha256(key_material).hexdigest()


@st.cache_data(show_spinner=False, max_entries=512)
def _cached_dataframe_to_xlsx(cache_key: str, _frame: pd.DataFrame) -> bytes:
    del cache_key
    frame = _frame
    output = BytesIO()
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("Dados")
    worksheet.append([str(column) for column in frame.columns])
    for row in frame.itertuples(index=False, name=None):
        values = []
        for value in row:
            if pd.isna(value):
                values.append(None)
            elif isinstance(value, pd.Timestamp):
                values.append(value.to_pydatetime())
            elif hasattr(value, "item") and not isinstance(value, (str, bytes)):
                values.append(value.item())
            else:
                values.append(value)
        worksheet.append(values)
    workbook.save(output)
    return output.getvalue()


def dataframe_to_xlsx(
    frame: pd.DataFrame, *, cache_key: str | None = None
) -> bytes:
    resolved_key = cache_key or dataframe_xlsx_cache_key(frame)
    return _cached_dataframe_to_xlsx(resolved_key, frame)


def detail_export_frame(records: pd.DataFrame) -> pd.DataFrame:
    """Return original imported columns plus the dashboard type and aging."""
    source_columns = [
        column for column in DETAIL_SOURCE_COLUMNS
        if column not in {"TIPO", "AGING"}
    ]
    detail_columns: dict[str, pd.Series] = {}
    for column in source_columns:
        if column in records.columns:
            detail_columns[column] = records[column]
        elif column in _RAW_BASE.columns:
            detail_columns[column] = _RAW_BASE[column].reindex(records.index)
        else:
            detail_columns[column] = pd.Series(pd.NA, index=records.index)
    details = pd.DataFrame(detail_columns, index=records.index)

    if "TIPO" in records.columns:
        details["TIPO"] = records["TIPO"]
    else:
        details["TIPO"] = BASE["TIPO"].reindex(records.index)

    age_column = next(
        (
            column for column in ("DATA_ULTIMA_DATA", "DT_ENTRADA_REVERSAS", "DATA_EVENTO")
            if column in records.columns or column in _RAW_BASE.columns
        ),
        None,
    )
    if age_column is None:
        details["AGING"] = pd.NA
    else:
        age_source = records if age_column in records.columns else _RAW_BASE
        age_dates = pd.to_datetime(age_source[age_column], errors="coerce")
        age_dates = age_dates.reindex(records.index)
        aging = powerbi_aging_bucket(age_dates)
        details["AGING"] = aging.where(age_dates.notna(), pd.NA)
    return details


def is_total_row(frame: pd.DataFrame) -> bool:
    if frame.empty or not len(frame.columns):
        return False
    label = frame.iloc[-1, 0]
    return isinstance(label, str) and label.strip().casefold() == "total"


def sticky_table_html(
    frame: pd.DataFrame,
    *,
    bar_column: str | None = None,
    cell_links: dict[tuple[int, str], str] | None = None,
) -> str:
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

    def render_row(
        row: pd.Series,
        row_index: int,
        *,
        total: bool = False,
    ) -> str:
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
            content = cell_value(value, str(column))
            cell_link = (cell_links or {}).get((row_index, str(column)))
            if cell_link is not None:
                content = (
                    f'<a class="table-export-link" href="#" '
                    f'data-cell-id="{esc(cell_link)}">{content}</a>'
                )
            cells.append(f"<td{class_name}>{content}</td>")
        row_class = ' class="total-row"' if total else ""
        return f"<tr{row_class}>{''.join(cells)}</tr>"

    header = "".join(
        f'<th class="distribution">{esc(column)}</th>'
        if bar_column is not None and column == bar_column
        else f"<th>{esc(column)}</th>"
        for column in columns
    )
    has_total = is_total_row(frame)
    body_frame = frame.iloc[:-1] if has_total else frame
    body = "".join(
        render_row(row, row_index)
        for row_index, (_, row) in enumerate(body_frame.iterrows())
    )
    total = (
        render_row(frame.iloc[-1], len(body_frame), total=True)
        if has_total
        else ""
    )
    shipper_summary = columns[:3] in (["SHIPPER", "QTDE", "%"], ["CLIENTES", "QTDE", "%"])
    table_class = "shipper-summary" if shipper_summary else ""
    return f"""
    <style>
      * {{ box-sizing: border-box; }}
      html, body {{ width: 100%; height: 100%; margin: 0; padding: 0; background: #fff; font-family: Roboto, Arial, sans-serif; }}
      .scroll-area {{ height: 100%; max-height: 100%; overflow: auto; border: 1px solid #eef0f4; border-radius: 9px; scrollbar-color: #c7ccd5 transparent; scrollbar-width: thin; }}
      table {{ width: max-content; min-width: 100%; border-collapse: separate; border-spacing: 0; color: #30394a; font-size: 11px; line-height: 1.25; }}
      table.shipper-summary {{ width: 100%; table-layout: fixed; }}
      table.shipper-summary th:nth-child(1), table.shipper-summary td:nth-child(1) {{ width: 62%; }}
      table.shipper-summary th:nth-child(2), table.shipper-summary td:nth-child(2) {{ width: 10%; }}
      table.shipper-summary th:nth-child(3), table.shipper-summary td:nth-child(3) {{ width: 28%; }}
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
      a.table-export-link {{ color: inherit; text-decoration: underline; text-decoration-color: #aeb7c5; text-underline-offset: 2px; cursor: pointer; }}
      a.table-export-link:hover {{ color: #e5003d; text-decoration-color: #e5003d; }}
    </style>
    <div class="scroll-area"><table class="{table_class}"><thead><tr>{header}</tr></thead>
      <tbody>{body}</tbody>{f'<tfoot>{total}</tfoot>' if has_total else ''}
    </table></div>
    """


@st.dialog("Modo foco", width="large")
def show_table_focus(
    title: str,
    frame: pd.DataFrame,
    bar_column: str | None = None,
    *,
    native: bool = False,
    cell_export_context: tuple[pd.DataFrame, str | None, str | None] | None = None,
) -> None:
    st.markdown(f"**{esc(title)}**")
    if native:
        st.dataframe(frame, use_container_width=True, hide_index=True, height=700)
        if "CTE" in frame.columns:
            st.caption(f"Total · {number_br(frame['CTE'].notna().sum())}")
        return
    if cell_export_context is not None:
        records, row_field, bucket_field = cell_export_context
        _linked_table(
            frame,
            title=f"{title}_focus",
            height=720,
            bar_column=bar_column,
            export_records=records,
            row_field=row_field,
            bucket_field=bucket_field,
        )
        return
    focus_height = min(720, max(150, 54 + len(frame) * 32))
    components.html(
        sticky_table_html(frame, bar_column=bar_column),
        height=focus_height,
        scrolling=False,
    )


def _linked_table(
    frame: pd.DataFrame,
    *,
    title: str,
    height: int,
    bar_column: str | None,
    export_records: pd.DataFrame,
    row_field: str | None,
    bucket_field: str | None,
) -> None:
    """Render the original HTML table and generate each XLSX after its click."""
    row_labels = (
        export_records[row_field].fillna("(Em branco)").astype(str)
        if row_field is not None
        else pd.Series("Total", index=export_records.index, dtype="string")
    )
    rows_by_label = row_labels.groupby(row_labels, sort=False).indices
    buckets_by_label: dict[str, object] = {}
    cells_by_pair: dict[tuple[str, str], object] = {}
    if bucket_field is not None:
        bucket_labels = export_records[bucket_field].fillna("(Em branco)").astype(str)
        buckets_by_label = bucket_labels.groupby(bucket_labels, sort=False).indices
        pairs = pd.DataFrame(
            {"row": row_labels.to_numpy(), "bucket": bucket_labels.to_numpy()}
        )
        cells_by_pair = pairs.groupby(["row", "bucket"], sort=False).indices

    all_positions = list(range(len(export_records)))
    table_slug = "".join(
        character.lower() if character.isalnum() else "_"
        for character in title
    ).strip("_")

    def export_positions(row_name: str, column: str) -> object | None:
        if row_name == "Total" or row_field is None:
            if bucket_field is not None and column != "Total":
                return buckets_by_label.get(str(column))
            return all_positions
        if bucket_field is not None and column != "Total":
            return cells_by_pair.get((row_name, str(column)))
        return rows_by_label.get(row_name)

    body = frame.iloc[:-1] if is_total_row(frame) else frame
    rows = list(body.iterrows())
    if is_total_row(frame):
        rows.append((len(frame) - 1, frame.iloc[-1]))
    cell_exports: dict[str, tuple[tuple[int, ...], str]] = {}
    cell_links: dict[tuple[int, str], str] = {}
    for display_index, (_, row) in enumerate(rows):
        row_name = str(row.iloc[0])
        for column_index, (column, value) in enumerate(row.items()):
            if not pd.api.types.is_number(value) or isinstance(value, bool):
                continue
            positions = export_positions(row_name, str(column))
            if positions is None or not len(positions):
                continue
            row_slug = "".join(
                character.lower() if character.isalnum() else "_"
                for character in row_name
            ).strip("_") or "total"
            column_slug = "".join(
                character.lower() if character.isalnum() else "_"
                for character in str(column)
            ).strip("_")
            filename = f"{table_slug}_{row_slug}_{column_slug}.xlsx"
            cell_id = f"{display_index}_{column_index}"
            cell_links[(display_index, str(column))] = cell_id
            cell_exports[cell_id] = (
                tuple(int(position) for position in positions),
                filename,
            )
    table_html = sticky_table_html(
        frame,
        bar_column=bar_column,
        cell_links=cell_links,
    )
    state_key = f"_lazy_export_pending_{table_slug}"
    pending_download = st.session_state.get(state_key)
    action = lazy_export_table(
        table_html=table_html,
        height=int(height),
        download=pending_download,
        key=f"lazy_export_table_{table_slug}",
        default=None,
    )
    if isinstance(action, dict):
        if action.get("type") == "export" and pending_download is None:
            selected = cell_exports.get(str(action.get("cell_id", "")))
            if selected is not None:
                positions, filename = selected
                details = detail_export_frame(
                    export_records.iloc[list(positions)]
                )
                cache_key = dataframe_xlsx_cache_key(
                    details,
                    stable_source_rows=True,
                )
                xlsx_bytes = dataframe_to_xlsx(details, cache_key=cache_key)
                st.session_state[state_key] = {
                    "token": uuid.uuid4().hex,
                    "filename": filename,
                    "base64": base64.b64encode(xlsx_bytes).decode("ascii"),
                }
                st.rerun()
        elif action.get("type") == "downloaded" and pending_download is not None:
            if action.get("token") == pending_download.get("token"):
                del st.session_state[state_key]
                st.rerun()


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
    key_suffix: str = "",
    export_name: str | None = None,
    export_frame: pd.DataFrame | None = None,
    cell_export_context: tuple[pd.DataFrame, str | None, str | None] | None = None,
) -> None:
    # Older cached matrix frames may still carry this bucket; hide it at the
    # shared table renderer so standard, linked, and focus tables stay aligned.
    view = frame.drop(columns=["Sem dados"], errors="ignore").copy()
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
        if export_name:
            title_col, actions_col = st.columns([1, .24], gap="small", vertical_alignment="center")
        else:
            title_col, focus_col = st.columns([1, .1], gap="small", vertical_alignment="center")
        with title_col:
            st.markdown(f'<div class="panel-title">{esc(title)}</div>', unsafe_allow_html=True)
        focus_key = f"focus_{title.casefold().replace(' ', '_')}"
        if key_suffix:
            focus_key = f"{focus_key}_{key_suffix}"
        if export_name:
            source_export = export_frame if export_frame is not None else view

            def build_export() -> bytes:
                export_details = detail_export_frame(source_export)
                export_cache_key = dataframe_xlsx_cache_key(
                    export_details, stable_source_rows=export_frame is not None
                )
                return dataframe_to_xlsx(export_details, cache_key=export_cache_key)

            with actions_col:
                buttons_col, _right_padding = st.columns([1, .08], gap=None)
                with buttons_col:
                    with st.container(horizontal=True, horizontal_alignment="right", vertical_alignment="center", gap="small"):
                        st.download_button(
                            "⇩",
                            data=build_export,
                            file_name=export_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"export_{key_suffix or title.casefold().replace(' ', '_')}",
                            help="Exportar tabela em XLSX",
                            on_click="ignore",
                        )
                        focus_clicked = st.button("⛶", key=focus_key, help="Abrir em modo foco")
        else:
            with focus_col:
                focus_clicked = st.button("⛶", key=focus_key, help="Abrir em modo foco")
        if focus_clicked:
            show_table_focus(
                title,
                view,
                bar_column,
                cell_export_context=cell_export_context,
            )
        if cell_export_context is not None:
            export_records, row_field, bucket_field = cell_export_context
            _linked_table(
                view,
                height=height,
                bar_column=bar_column,
                export_records=export_records,
                row_field=row_field,
                bucket_field=bucket_field,
                title=f"{title}_{key_suffix}" if key_suffix else title,
            )
        elif has_total and not view.empty:
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
    st.markdown(
        '<div class="kpi-card">'
        f'<div class="kpi-icon">{esc(icon)}</div>'
        '<div class="kpi-copy">'
        f'<div class="kpi-label" title="{esc(label)}">{esc(label)}</div>'
        f'<div class="kpi-value" title="{esc(value)}">{esc(value)}</div>'
        '</div></div>',
        unsafe_allow_html=True,
    )


def metric_cards(columns: list, df: pd.DataFrame) -> None:
    amount = float(df["VALOR_NF"].sum()) if not df.empty else 0.0
    count = int(df["CTE"].notna().sum()) if not df.empty else 0
    with columns[0]:
        kpi("Soma de VALOR_NF", money_short(amount), "＄", "")
    with columns[1]:
        kpi("Quantidade", count_short(count), "▤", "")


def header(page: str, *, backlog_type: bool = False) -> None:
    page_title = "LISTA DE CTE" if page == "Lista CTE" else page.upper()
    header_base = BASE if page == "BACKLOG" else CLASSIFIED_BASE
    data_cutoff = header_base["DATA_ATUAL"].max()
    cutoff_label = data_cutoff.strftime("%d/%m/%Y") if pd.notna(data_cutoff) else "Sem data"
    header_markup = (
        '<div class="report-header"><div><div class="report-kicker">JADLOG · OPERAÇÕES DE REVERSA</div>'
        f'<div class="report-title"><span class="page-name">{esc(page_title)}</span>'
        '<span class="divider">/</span><span class="branch">TC REVERSAS JAD-SAO</span></div></div>'
        f'<div class="report-subtitle"><b>Dados até</b><br><strong>{cutoff_label}</strong></div></div>'
    )
    header_columns = st.columns([6, 1], gap="small", vertical_alignment="center")
    with header_columns[0]:
        st.markdown(header_markup, unsafe_allow_html=True)
    with header_columns[1]:
        with st.container(key="page-filter-clear"):
            st.button(
                "Limpar filtros",
                key=f"clear_filters_{page.lower().replace(' ', '_')}",
                on_click=clear_page_filters,
                args=(page,),
                help="Restaurar os filtros desta aba",
            )


def clear_page_filters(page: str) -> None:
    """Reset only the current page's filter widgets."""
    date_filter_keys = {
        "backlog_sol", "backlog_evento", "backlog_emissao",
        "entrada_data", "entrada_sol", "cc_data_v2", "cc_sol_v2",
        "cl_data", "cl_sol", "fv_data", "lista_sol", "lista_emissao",
    }
    for key in PAGE_FILTER_KEYS.get(page, []):
        if key in date_filter_keys:
            st.session_state[key] = None
        elif key == "lista_busca":
            st.session_state[key] = ""
        else:
            # Explicitly clear widget state in the button callback so Streamlit
            # cannot restore the prior selection when the widget is recreated.
            st.session_state[key] = []
        for state_key in (f"_preferred_period_{key}", f"_multiselect_cleared_{key}"):
            if state_key in st.session_state:
                del st.session_state[state_key]


def kpi_row(df: pd.DataFrame, *, fixed_value: float | None = None) -> None:
    cols = st.columns([1, 1], gap="small")
    metric_cards(cols, df)


def period_picker(df: pd.DataFrame, field: str, label: str, key: str, *, default: tuple[date, date] | None = None) -> tuple[date, date] | None:
    values = pd.to_datetime(df[field], errors="coerce").dropna()
    if values.empty:
        return None
    lo, hi = values.min().date(), values.max().date()
    selected_default = default or (lo, hi)
    selected_default = (max(lo, selected_default[0]), hi)
    value = st.date_input(label, value=selected_default, min_value=lo, max_value=hi, format="DD/MM/YYYY", key=key)
    if isinstance(value, tuple) and len(value) == 2:
        return value
    return None


def apply_period(
    df: pd.DataFrame,
    field: str,
    period: tuple[date, date] | None,
    *,
    include_missing_when_full: bool = False,
) -> pd.DataFrame:
    if period is None:
        return df
    dates = pd.to_datetime(df[field], errors="coerce").dt.date
    mask = dates.between(period[0], period[1])
    valid_dates = dates.dropna()
    if (
        include_missing_when_full
        and not valid_dates.empty
        and period[0] <= valid_dates.min()
        and period[1] >= valid_dates.max()
    ):
        mask |= dates.isna()
    return df.loc[mask]


def filtered(df: pd.DataFrame, fields: dict[str, list[str]]) -> pd.DataFrame:
    out = df
    for field, values in fields.items():
        if values:
            include_blank = "(Em branco)" in values
            selected = [value for value in values if value != "(Em branco)"]
            series = out[field].astype("string")
            mask = series.isin(selected)
            if include_blank:
                mask |= series.isna() | series.fillna("").str.strip().eq("")
            out = out[mask]
    return out


def filter_vencimento(df: pd.DataFrame, selected: list[str]) -> pd.DataFrame:
    if not selected:
        return df
    vencimento = df["VENCIMENTO"]
    if not pd.api.types.is_datetime64_any_dtype(vencimento):
        vencimento = pd.to_datetime(vencimento, errors="coerce")
    today = pd.Timestamp.today().normalize()
    mask = pd.Series(False, index=df.index)
    if "No prazo" in selected:
        mask |= vencimento.ge(today)
    if "Vencido" in selected:
        mask |= vencimento.lt(today)
    return df.loc[mask]


def filter_select(label: str, options: list[str], key: str) -> list[str]:
    return st.multiselect(
        label,
        options,
        key=key,
        placeholder="Todos",
        label_visibility="visible",
        on_change=_remember_multiselect_change,
        args=(key,),
    )


def _selection_values(value: object) -> list[str]:
    if isinstance(value, list):
        return [str(item) for item in value]
    return [str(value)] if value else []


def _multiselect_clear_marker(key: str) -> str:
    return f"_multiselect_cleared_{key}"


def _remember_multiselect_change(key: str) -> None:
    """Record an explicit Clear all action before cascading filters rerun."""
    st.session_state[_multiselect_clear_marker(key)] = not bool(
        _selection_values(st.session_state.get(key, []))
    )


def _consume_multiselect_clear(key: str) -> bool:
    marker = _multiselect_clear_marker(key)
    cleared = bool(st.session_state.get(marker, False))
    if marker in st.session_state:
        del st.session_state[marker]
    if cleared:
        st.session_state[key] = []
    return cleared


def _preferred_period_key(key: str) -> str:
    return f"_preferred_period_{key}"


def _remember_user_period_change(key: str) -> None:
    """Keep the user's date selection separate from temporary synchronized bounds."""
    value = st.session_state.get(key)
    if not isinstance(value, (tuple, list)) or len(value) != 2:
        return
    try:
        start, end = pd.Timestamp(value[0]), pd.Timestamp(value[1])
    except (TypeError, ValueError, OverflowError):
        return
    if pd.isna(start) or pd.isna(end):
        return
    st.session_state[_preferred_period_key(key)] = (start.date(), end.date())


def _apply_filter_spec(frame: pd.DataFrame, spec: dict, selected: list[str]) -> pd.DataFrame:
    if spec.get("kind") == "vencimento":
        return filter_vencimento(frame, selected)
    return filtered(frame, {spec["field"]: selected})


def _apply_filter_specs(
    frame: pd.DataFrame,
    specs: list[dict],
    selections: dict[str, list[str]],
    *,
    exclude_group: str | None = None,
    exclude_key: str | None = None,
) -> pd.DataFrame:
    """Apply ordinary facets and combine customer/store facets within TIPO."""
    out = frame
    handled_groups: set[str] = set()
    for spec in specs:
        if spec["key"] == exclude_key:
            continue
        group = spec.get("selection_group")
        if group:
            if group == exclude_group or group in handled_groups:
                continue
            handled_groups.add(group)
            members = [item for item in specs if item.get("selection_group") == group]
            active = [item for item in members if selections.get(item["key"])]
            if active and "TIPO" in out.columns:
                mask = pd.Series(False, index=out.index)
                for item in active:
                    category_mask = out["TIPO"].eq(item["classification_type"])
                    values = out[item["field"]].astype("string")
                    selected = selections[item["key"]]
                    item_mask = values.isin([value for value in selected if value != "(Em branco)"])
                    if "(Em branco)" in selected:
                        item_mask |= values.isna() | values.fillna("").str.strip().eq("")
                    mask |= category_mask & item_mask
                out = out.loc[mask]
            elif active:
                # Frames without classification metadata retain the usual
                # field-by-field behavior.
                for item in active:
                    out = _apply_filter_spec(out, item, selections[item["key"]])
            continue
        out = _apply_filter_spec(out, spec, selections.get(spec["key"], []))
    return out


def synchronized_multiselects(
    frame: pd.DataFrame,
    specs: list[dict],
    columns: list | None = None,
) -> dict[str, list[str]]:
    """Render cascading filters whose options reflect every other active filter."""
    selections = {}
    for spec in specs:
        key = spec["key"]
        _consume_multiselect_clear(key)
        selections[key] = _selection_values(st.session_state.get(key, []))
    for index, spec in enumerate(specs):
        candidates = _apply_filter_specs(
            frame,
            specs,
            selections,
            exclude_group=spec.get("selection_group"),
            exclude_key=spec["key"] if not spec.get("selection_group") else None,
        )

        if spec.get("kind") == "vencimento":
            options = [
                value for value in VENCIMENTO_FILTER_OPTIONS
                if not filter_vencimento(candidates, [value]).empty
            ]
        else:
            values = candidates[spec["field"]].fillna("").astype(str).str.strip()
            present = set(values[values.ne("")])
            if spec.get("include_blank") and (values.eq("").any()):
                present.add("(Em branco)")
            allowed = spec.get("options")
            options = [value for value in allowed if value in present] if allowed is not None else sorted(present, key=str.casefold)

        key = spec["key"]
        current = _selection_values(st.session_state.get(key, []))
        selected = [value for value in selections[key] if value in options]
        # Only write back when dynamic options invalidated a previous choice.
        # Avoid reassigning an already-cleared widget value before it renders,
        # which can make Streamlit restore the old selection on rerun.
        if selected != current:
            st.session_state[key] = selected
        selections[key] = selected
        if columns is None:
            selections[spec["key"]] = st.multiselect(
                spec["label"], options, key=spec["key"],
                placeholder=spec.get("placeholder", "Todos"),
                on_change=_remember_multiselect_change,
                args=(spec["key"],),
            )
        else:
            with columns[index]:
                selections[spec["key"]] = st.multiselect(
                    spec["label"], options, key=spec["key"],
                    placeholder=spec.get("placeholder", "Todos"),
                    on_change=_remember_multiselect_change,
                    args=(spec["key"],),
                )
    return selections


def synchronized_periods(
    frame: pd.DataFrame,
    period_specs: list[dict],
    filter_specs: list[dict] | None = None,
    columns: list | None = None,
) -> dict[str, tuple[date, date] | None]:
    """Make each date range respond to the other dates and active dimension filters."""
    def as_date(value: object) -> date | None:
        if value is None:
            return None
        try:
            parsed = pd.Timestamp(value)
            return None if pd.isna(parsed) else parsed.date()
        except (TypeError, ValueError, OverflowError):
            return None

    def complete_period(value: object) -> tuple[date, date] | None:
        if not isinstance(value, (tuple, list)) or len(value) != 2:
            return None
        start, end = as_date(value[0]), as_date(value[1])
        return (start, end) if start is not None and end is not None else None

    filter_specs = filter_specs or []
    for spec in filter_specs:
        _consume_multiselect_clear(spec["key"])
    selections = {
        spec["key"]: _selection_values(st.session_state.get(spec["key"], []))
        for spec in filter_specs
    }
    preferred_periods: dict[str, tuple[date, date] | None] = {}
    for spec in period_specs:
        key = spec["key"]
        preferred_key = _preferred_period_key(key)
        preferred = complete_period(st.session_state.get(preferred_key))
        if preferred is None:
            # Migrate older sessions from the source date range, not the old
            # widget value: earlier versions overwrote that value with the
            # temporary limits produced by cascading filters.
            source_dates = pd.to_datetime(frame[spec["field"]], errors="coerce").dropna()
            if not source_dates.empty:
                source_lo, source_hi = source_dates.min().date(), source_dates.max().date()
                default = spec.get("default")
                start = min(max(default[0], source_lo), source_hi) if default else source_lo
                preferred = (start, source_hi)
                st.session_state[preferred_key] = preferred
        preferred_periods[key] = preferred
    periods = dict(preferred_periods)

    for index, spec in enumerate(period_specs):
        candidates = frame
        for other in period_specs:
            if other["key"] != spec["key"] and periods[other["key"]] is not None:
                candidates = apply_period(
                    candidates, other["field"], periods[other["key"]],
                    include_missing_when_full=other.get("include_missing_when_full", False),
                )
        candidates = _apply_filter_specs(candidates, filter_specs, selections)

        values = pd.to_datetime(candidates[spec["field"]], errors="coerce").dropna()
        if values.empty:
            values = pd.to_datetime(frame[spec["field"]], errors="coerce").dropna()
        if values.empty:
            periods[spec["key"]] = None
            continue

        lo, hi = values.min().date(), values.max().date()
        prior = preferred_periods[spec["key"]]
        key = spec["key"]
        raw_value = st.session_state.get(key)
        is_partial_pair = (
            isinstance(raw_value, (tuple, list))
            and len(raw_value) == 2
            and complete_period(raw_value) is None
        )
        if is_partial_pair or (prior is None and key in st.session_state and raw_value is None):
            # Streamlit returns a partial pair while the user is picking the
            # second endpoint. Keep that in-progress choice instead of
            # replacing it with the saved range or comparing a date with None.
            if raw_value is None:
                widget_value = None
            else:
                partial = [as_date(value) for value in raw_value[:2]]
                widget_value = tuple(
                    min(max(value, lo), hi) if value is not None else None
                    for value in partial
                )
        elif prior is None:
            default = spec.get("default")
            start = min(max(default[0], lo), hi) if default else lo
            end = hi
            widget_value = (start, end)
            preferred_periods[key] = widget_value
            st.session_state[_preferred_period_key(key)] = widget_value
        else:
            # Apply other filters' limits to the visible widget only. Keep the
            # preferred range intact so clearing a filter restores it.
            start = max(prior[0], lo)
            end = min(prior[1], hi)
            if start > end:
                start, end = lo, hi
            widget_value = (start, end)
        if st.session_state.get(key) != widget_value:
            st.session_state[key] = widget_value
        if columns is None:
            result = st.date_input(
                spec["label"], min_value=lo, max_value=hi,
                format="DD/MM/YYYY", key=key,
                on_change=_remember_user_period_change, args=(key,),
            )
        else:
            with columns[index]:
                result = st.date_input(
                    spec["label"], min_value=lo, max_value=hi,
                    format="DD/MM/YYYY", key=key,
                    on_change=_remember_user_period_change, args=(key,),
                )
        periods[key] = complete_period(result)
    return periods


@st.cache_data(show_spinner=False, max_entries=256)
def summary_table(
    df: pd.DataFrame,
    field: str,
    *,
    limit: int | None = 12,
    options: list[str] | None = None,
    display_field: str | None = None,
) -> pd.DataFrame:
    values = df[field].fillna("(Em branco)").astype(str)
    grouped = values.value_counts().rename("QTDE")
    if options is not None:
        grouped = grouped.reindex(options, fill_value=0)
    grouped = grouped.loc[grouped.gt(0)]
    out = grouped.rename_axis(field).reset_index()
    out = out.sort_values(
        ["QTDE", field],
        ascending=[False, True],
        kind="mergesort",
    )
    if limit is not None:
        out = out.head(limit)
    total_count = len(df)
    base = max(total_count, 1)
    out["%"] = out["QTDE"].map(lambda val: 100 * val / max(base, 1))
    # The bar and printed percentage use the same whole-table denominator.
    out["BARRA"] = out["%"]
    out["%"] = out["%"].map(lambda val: f"{decimal_br(val, 2)}%")
    total_pct = 100.0 if total_count else 0.0
    label_field = display_field or field
    if display_field:
        out = out.rename(columns={field: display_field})
    total = pd.DataFrame([{label_field: "Total", "QTDE": total_count, "%": f"{decimal_br(total_pct, 2)}%", "BARRA": total_pct}])
    return pd.concat([out, total], ignore_index=True)


@st.cache_data(show_spinner=False, max_entries=256)
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
    dimension_labels = {"SHIPPER": "CLIENTES", "EMISSOR": "LOJAS E PA'S"}
    display_field = dimension_labels.get(row_field, row_field.replace("_", " "))
    if df.empty:
        return pd.DataFrame(columns=[display_field, *labels, "Total"])
    row_values = df[row_field].fillna("(Em branco)")
    pivot = pd.crosstab(row_values, df[bucket_field].fillna("(Em branco)"))
    columns = order or labels
    pivot = pivot.reindex(columns=columns, fill_value=0)
    # Keep records with a missing aging bucket in the row/grand totals even
    # though they do not appear in one of the visible aging columns.
    pivot["Total"] = row_values.value_counts(dropna=False).reindex(pivot.index, fill_value=0)
    all_rows = pivot.copy()
    pivot = pivot.sort_values("Total", ascending=False).head(limit)
    grand = {name: int(all_rows[name].sum()) for name in columns}
    grand["Total"] = int(len(df))
    grand[row_field] = "Total"
    pivot = pivot.reset_index()
    pivot = pd.concat([pivot, pd.DataFrame([grand])], ignore_index=True)
    pivot = pivot.rename(columns={row_field: display_field})
    return pivot


def aging_totals(df: pd.DataFrame, label: str, bucket_field: str) -> pd.DataFrame:
    row = {"STATUS ATUAL": label}
    counts = df[bucket_field].value_counts(dropna=False)
    row.update({bucket: int(counts.get(bucket, 0)) for bucket in BUCKET_LABELS})
    row["Total"] = len(df)
    return pd.DataFrame([row])


def footer_detail(
    df: pd.DataFrame,
    *,
    title: str = "DETALHES DOS CTEs",
    height: int = 280,
    export_name: str = "detalhes_ctes.xlsx",
    key_suffix: str = "detail_cte",
) -> None:
    """Show original imported detail columns with dashboard classifications."""
    chosen = detail_export_frame(df)
    detail_column_config: dict[str, object] = {
        "CTE": st.column_config.TextColumn("CTE", width="medium"),
        "PEDIDO": st.column_config.TextColumn("PEDIDO", width="medium"),
        "STATUS_ATUAL": st.column_config.TextColumn("STATUS_ATUAL", width="medium"),
        "DESCRICAO_ATUAL": st.column_config.TextColumn("DESCRICAO_ATUAL", width="medium"),
        "STATUS_TRATATIVA": st.column_config.TextColumn("STATUS_TRATATIVA", width="large"),
        "RESULTADO DA CONSULTA": st.column_config.TextColumn("RESULTADO DA CONSULTA", width="large"),
        "TIPO": st.column_config.TextColumn("TIPO", width="medium"),
        "AGING": st.column_config.TextColumn("AGING", width="small"),
    }
    for column in chosen.columns:
        if pd.api.types.is_datetime64_any_dtype(chosen[column]):
            detail_column_config[column] = st.column_config.DateColumn(
                column, format="DD/MM/YYYY", width="small"
            )

    with st.container(border=True):
        title_col, actions_col = st.columns([1, .24], gap="small", vertical_alignment="center")
        with title_col:
            st.markdown(f'<div class="panel-title">{esc(title)}</div>', unsafe_allow_html=True)
        with actions_col:
            buttons_col, _right_padding = st.columns([1, .08], gap=None)
            with buttons_col:
                with st.container(horizontal=True, horizontal_alignment="right", vertical_alignment="center", gap="small"):
                    st.download_button(
                        "⇩",
                        data=lambda: dataframe_to_xlsx(
                            chosen,
                            cache_key=dataframe_xlsx_cache_key(
                                chosen, stable_source_rows=True
                            ),
                        ),
                        file_name=export_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"export_{key_suffix}",
                        help="Exportar tabela em XLSX",
                        on_click="ignore",
                    )
                    focus_clicked = st.button("⛶", key=f"focus_{key_suffix}", help="Abrir tabela geral em modo foco")
        if focus_clicked:
            show_table_focus(title, chosen, native=True)

        st.dataframe(
            chosen,
            use_container_width=True,
            hide_index=True,
            height=height,
            column_config=detail_column_config,
        )
        st.caption(f"Total · {number_br(chosen['CTE'].notna().sum())}")


def backlog_page() -> None:
    df = BASE.copy()

    if "backlog_shipper" in st.session_state:
        previous_clients = st.session_state["backlog_shipper"]
        if not isinstance(previous_clients, list):
            previous_clients = [previous_clients] if previous_clients else []
        st.session_state["backlog_shipper"] = [
            client for client in previous_clients if client in CLIENT_FILTER_NAMES
        ]
    if "backlog_emissor" in st.session_state:
        previous_emitters = st.session_state["backlog_emissor"]
        if not isinstance(previous_emitters, list):
            previous_emitters = [previous_emitters] if previous_emitters else []
        st.session_state["backlog_emissor"] = [
            emitter for emitter in previous_emitters if emitter in EMISSOR_FILTER_NAMES
        ]

    # First row mirrors the small Power BI cards and slicers.
    top = st.columns([1, 1, .72, .75, 1.1, 1.1, 1.1, .78, .82], gap="small", vertical_alignment="center")
    with top[0]:
        kpi("Soma de VALOR_NF", money_short(float(df["VALOR_NF"].sum())), "＄", "")
    with top[1]:
        kpi("Quantidade", count_short(len(df)), "▤", "")
    backlog_filter_specs = [
        {"key": "backlog_emissor", "label": "LOJAS E PA'S", "field": "EMISSOR", "options": EMISSOR_FILTER_NAMES, "selection_group": "party_type", "classification_type": "Lojas e PAs"},
        {"key": "backlog_shipper", "label": "CLIENTES", "field": "SHIPPER", "options": CLIENT_FILTER_NAMES, "selection_group": "party_type", "classification_type": "Cliente"},
        {"key": "backlog_status", "label": "STATUS_ATUAL", "field": "STATUS_ATUAL", "options": BACKLOG_STATUS_NAMES},
        {"key": "backlog_vencimento", "label": "VENCIMENTO", "kind": "vencimento"},
    ]
    backlog_periods = synchronized_periods(
        df,
        [
            {"key": "backlog_sol", "label": "DATA_SOLICITAÇÃO", "field": "SOLICITACAO", "default": (date(2024, 7, 1), date(2026, 9, 3)), "include_missing_when_full": True},
            {"key": "backlog_evento", "label": "DATA_EVENTO", "field": "DATA_ATUAL", "default": (date(2024, 7, 10), date(2026, 9, 7)), "include_missing_when_full": True},
            {"key": "backlog_emissao", "label": "DT_EMISSAO", "field": "DT_EMISSAO", "include_missing_when_full": True},
        ],
        backlog_filter_specs,
        [top[4], top[5], top[6]],
    )
    request_period = backlog_periods["backlog_sol"]
    event_period = backlog_periods["backlog_evento"]
    emission_period = backlog_periods["backlog_emissao"]
    df = apply_period(df, "SOLICITACAO", request_period, include_missing_when_full=True)
    df = apply_period(df, "DATA_ATUAL", event_period, include_missing_when_full=True)
    df = apply_period(df, "DT_EMISSAO", emission_period, include_missing_when_full=True)
    synced = synchronized_multiselects(
        df,
        backlog_filter_specs,
        [top[2], top[3], top[7], top[8]],
    )
    df = _apply_filter_specs(df, backlog_filter_specs, synced)

    conflicts = df.loc[backlog_conflict_mask(df), [
        "CTE", "PEDIDO", "SHIPPER", "EMISSOR", "STATUS_ATUAL", "VALOR_NF"
    ]].copy()
    pair_assignments = st.session_state.setdefault("backlog_conflict_pair_assignments_v2", {})
    pending_conflicts = conflicts.copy()
    if not conflicts.empty:
        conflicts["_pair_key"] = (
            conflicts["SHIPPER"].map(normalize_dimension_name)
            + "\x1f"
            + conflicts["EMISSOR"].map(normalize_dimension_name)
        )
        pair_options = (
            conflicts.groupby(["_pair_key", "SHIPPER", "EMISSOR"], dropna=False)
            .agg(PEDIDOS=("CTE", "size"), VALOR_NF=("VALOR_NF", "sum"))
            .reset_index()
        )
        for pair_key in pair_options["_pair_key"]:
            if pair_assignments.get(pair_key) not in {"CLIENTES", "LOJAS E PA'S"}:
                pair_assignments[pair_key] = "CLIENTES"
        pending_keys = set(
            pair_options.loc[
                ~pair_options["_pair_key"].map(pair_assignments).isin(["CLIENTES", "LOJAS E PA'S"]),
                "_pair_key",
            ]
        )
        pending_conflicts = conflicts.loc[conflicts["_pair_key"].isin(pending_keys)].drop(columns="_pair_key")

    left, mid, right = st.columns([1.0, 1.0, .9], gap="small")
    client_rows = df.loc[df["TIPO"].eq("Cliente")]
    store_rows = df.loc[df["TIPO"].eq("Lojas e PAs")]
    categorized_rows = pd.concat([store_rows, client_rows]).sort_index()

    with left:
        store_summary = summary_table(store_rows, "EMISSOR", limit=None, options=EMISSOR_FILTER_NAMES, display_field="LOJAS E PA'S")
        table_panel("LOJAS E PA'S", store_summary, height=224, dark_header=True, bar_column="BARRA", numeric_columns=("QTDE",), total_note=number_br(store_rows.shape[0]), cell_export_context=(store_rows, "EMISSOR", None))
    with mid:
        client_summary = summary_table(client_rows, "SHIPPER", limit=None, options=CLIENT_FILTER_NAMES, display_field="CLIENTES")
        table_panel("CLIENTES", client_summary, height=224, dark_header=True, bar_column="BARRA", numeric_columns=("QTDE",), total_note=number_br(client_rows.shape[0]), cell_export_context=(client_rows, "SHIPPER", None))
    with right:
        status_summary = summary_table(categorized_rows, STATUS_DIMENSION, limit=None, options=STATUS_DIMENSION_NAMES)
        table_panel("STATUS", status_summary, height=224, dark_header=True, numeric_columns=("QTDE",), total_note=number_br(categorized_rows.shape[0]), cell_export_context=(categorized_rows, STATUS_DIMENSION, None))

    if not pending_conflicts.empty:
        pending_conflict_records = df.loc[df.index.isin(pending_conflicts.index)]
        with st.expander(f"Conflitos sem solução · {number_br(len(pending_conflicts))} pedidos"):
            table_panel(
                "CONFLITOS SEM SOLUÇÃO",
                pending_conflicts,
                height=280,
                key_suffix="pending_conflicts",
                export_name="conflitos_sem_solucao.xlsx",
                export_frame=pending_conflict_records,
            )

    assigned_indices = categorized_rows.index
    rua_rows = df.loc[~df.index.isin(assigned_indices)]
    with st.expander(f"RUA · {number_br(len(rua_rows))} pedidos"):
        rua_left, rua_mid, rua_right = st.columns([1, 1, .9], gap="small")
        with rua_left:
            rua_store_summary = summary_table(rua_rows, "EMISSOR", limit=None, display_field="LOJAS E PA'S")
            table_panel("LOJAS E PA'S", rua_store_summary, height=224, dark_header=True, bar_column="BARRA", numeric_columns=("QTDE",), total_note=number_br(len(rua_rows)), key_suffix="rua", cell_export_context=(rua_rows, "EMISSOR", None))
        with rua_mid:
            rua_client_summary = summary_table(rua_rows, "SHIPPER", limit=None, display_field="CLIENTES")
            table_panel("CLIENTES", rua_client_summary, height=224, dark_header=True, bar_column="BARRA", numeric_columns=("QTDE",), total_note=number_br(len(rua_rows)), key_suffix="rua", cell_export_context=(rua_rows, "SHIPPER", None))
        with rua_right:
            rua_status_summary = summary_table(rua_rows, STATUS_DIMENSION, limit=None, options=STATUS_DIMENSION_NAMES)
            table_panel("STATUS", rua_status_summary, height=224, dark_header=True, numeric_columns=("QTDE",), total_note=number_br(len(rua_rows)), key_suffix="rua", cell_export_context=(rua_rows, STATUS_DIMENSION, None))
        footer_detail(
            rua_rows,
            title="PEDIDOS EM RUA",
            height=280,
            export_name="pedidos_rua.xlsx",
            key_suffix="rua_pedidos",
        )

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
    cols = st.columns([1, 1, 1, 1] if include_aging else [1, 1, 1, .9], gap="small", vertical_alignment="center")
    selected = df.copy()
    # KPI widgets sit first, with date range slicers occupying the right half.
    metric_cards(cols[:2], selected)
    if include_aging:
        with cols[2]:
            period = period_picker(selected, "DATA_ATUAL", "DATA_EVENTO", f"{key}_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
        with cols[3]:
            sol = period_picker(selected, "SOLICITACAO", "DATA SOLICITAÇÃO", f"{key}_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    else:
        with cols[2]:
            period = period_picker(selected, "DATA_ATUAL", "DATA_EVENTO", f"{key}_data", default=(date(2024, 7, 10), date(2026, 9, 7)))
        with cols[3]:
            sol = period_picker(selected, "SOLICITACAO", "DATA SOLICITAÇÃO", f"{key}_sol", default=(date(2024, 7, 1), date(2026, 9, 3)))
    selected = apply_period(selected, "DATA_ATUAL", period)
    selected = apply_period(selected, "SOLICITACAO", sol)
    return selected


def entrada_page() -> None:
    analysis_base = CLASSIFIED_BASE
    if analysis_base.empty:
        st.info("Não há pedidos classificados como Cliente ou Lojas e PAs nesta base.")
        return
    top = st.columns([1, 1, 1.05, 1.05, .9, .85], gap="small", vertical_alignment="center")
    metric_cards(top[:2], analysis_base)
    entry_filter_specs = [
        {"key": "entrada_aging", "label": "AGING", "field": "AgingEntradaPBIX", "options": PBIX_ENTRY_AGE_BUCKETS},
        {"key": "entrada_vencimento", "label": "VENCIMENTO", "kind": "vencimento"},
    ]
    period_base = analysis_base.copy()
    period_base["AgingEntradaPBIX"] = powerbi_aging_bucket(
        period_base["DATA_EVENTO"], as_of=PBIX_ENTRY_AGE_AS_OF
    )
    entry_periods = synchronized_periods(
        period_base,
        [
            {"key": "entrada_data", "label": "DATA_EVENTO", "field": "DATA_ATUAL", "default": (date(2024, 7, 10), date(2026, 9, 7))},
            {"key": "entrada_sol", "label": "DATA SOLICITAÇÃO", "field": "SOLICITACAO", "default": (date(2024, 7, 1), date(2026, 9, 3))},
        ],
        entry_filter_specs,
        [top[2], top[3]],
    )
    at_period = entry_periods["entrada_data"]
    sol_period = entry_periods["entrada_sol"]
    date_base = apply_period(period_base, "DATA_ATUAL", at_period)
    date_base = apply_period(date_base, "SOLICITACAO", sol_period)
    synced = synchronized_multiselects(
        date_base,
        entry_filter_specs,
        [top[4], top[5]],
    )
    aging_filter = synced["entrada_aging"]
    due_filter = synced["entrada_vencimento"]
    selected = filtered(date_base, {"AgingEntradaPBIX": aging_filter})
    selected = filter_vencimento(selected, due_filter)

    a, b = st.columns([1.45, 1.0], gap="small")
    with a:
        manifest_blank = selected["MANIFESTO"].isna() | selected["MANIFESTO"].eq("")
        if selected["CIDADE_ORIGEM"].notna().any():
            city_not_closed = (
                selected["CIDADE_ORIGEM"].fillna("").str.strip().str.upper().ne("ENCERRADO")
            )
        else:
            city_not_closed = selected["FINALIZADO?"].ne("FINALIZADO")
        cli = selected[
            selected["TIPO"].eq("Cliente")
            & selected["STATUS_ATUAL"].isin(
                ["EM DEVOLUCAO - ENTRADA", "ENTRADA"]
            )
            & manifest_blank
            & city_not_closed
        ].copy()
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
            cell_export_context=(cli, "SHIPPER", "AgingEntradaPBIX"),
        )
    with b:
        cust = selected[
            selected["STATUS_ATUAL"].isin(["CUSTODIA", "DEVOLUCAO RECUSADA"])
            & manifest_blank
            & city_not_closed
        ].copy()
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
            cell_export_context=(cust, "DESCRICAO_ATUAL", "AgingEntradaPBIX"),
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
            cell_export_context=(stores, "EMISSOR", "AgingEntradaPBIX"),
        )
    with d:
        stuck_records = selected[
            selected["STATUS_ATUAL"].astype("string").str.upper().str.startswith("TRAVADO", na=False)
        ].copy()
        closed_mask = (
            selected["CIDADE_ORIGEM"].eq("ENCERRADO")
            if selected["CIDADE_ORIGEM"].notna().any()
            else selected["FINALIZADO?"].eq("FINALIZADO")
        )
        closed_records = selected[closed_mask].copy()
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
            cell_export_context=(stuck_records, None, "AgingEntradaPBIX"),
        )
        table_panel(
            "AGING - ENCERRADO",
            status_closed,
            height=66,
            numeric_columns=status_numeric_columns,
            cell_export_context=(closed_records, None, "AgingEntradaPBIX"),
        )
    footer_detail(selected if not selected.empty else analysis_base)


def control_clients_page() -> None:
    # Client control status and detail rows follow the same backlog source so
    # their totals stay consistent with the imported base.
    control_base = CONTROL_CLIENTS.iloc[:0].copy()
    # Keep this page's filters isolated from Backlog; only the source rows are shared.
    status_base = BASE.loc[BASE["TIPO"].eq("Cliente")].copy()
    top = st.columns([.9, .9, 1.35, 1.35, 1.05, 1.05, 1.05], gap="small", vertical_alignment="center")
    status_base["AgingControle"] = powerbi_aging_bucket(
        status_base["DATA_EVENTO"], as_of=PBIX_ENTRY_AGE_AS_OF
    )
    aging_key = "cc_aging_filter_v3"
    status_key = "cc_status_v2"
    shipper_key = "cc_shipper_v2"
    control_filter_specs = [
        {"key": aging_key, "label": "AGING", "field": "AgingControle", "options": PBIX_ENTRY_AGE_BUCKETS},
        {"key": status_key, "label": "STATUS", "field": "STATUS_ATUAL", "include_blank": True},
        {"key": shipper_key, "label": "CLIENTES", "field": "SHIPPER", "include_blank": True},
    ]
    control_periods = synchronized_periods(
        status_base,
        [
            {"key": "cc_data_v2", "label": "DATA_EVENTO", "field": "DATA_ATUAL"},
            {"key": "cc_sol_v2", "label": "DATA SOLICITAÇÃO", "field": "SOLICITACAO"},
        ],
        control_filter_specs,
        [top[2], top[3]],
    )
    data_period = control_periods["cc_data_v2"]
    sol_period = control_periods["cc_sol_v2"]
    df = apply_period(control_base, "DATA_ATUAL", data_period)
    df = apply_period(df, "SOLICITACAO", sol_period)
    df["AgingControle"] = powerbi_aging_bucket(
        df["DATA_EVENTO"], as_of=PBIX_ENTRY_AGE_AS_OF
    )
    status_rows = apply_period(
        status_base, "DATA_ATUAL", data_period, include_missing_when_full=True
    )
    status_rows = apply_period(
        status_rows, "SOLICITACAO", sol_period, include_missing_when_full=True
    )

    status_rows = status_rows.copy()
    status_rows["AgingControle"] = powerbi_aging_bucket(
        status_rows["DATA_EVENTO"], as_of=PBIX_ENTRY_AGE_AS_OF
    )

    synced = synchronized_multiselects(
        status_rows,
        control_filter_specs,
        [top[4], top[5], top[6]],
    )
    aging_filter = synced[aging_key]
    status_filter = synced[status_key]
    shipper_filter = synced[shipper_key]
    selected_dimensions = {"STATUS_ATUAL": status_filter, "SHIPPER": shipper_filter}
    df = filtered(df, selected_dimensions)
    status_rows = filtered(status_rows, selected_dimensions)
    metric_cards(top[:2], status_rows)

    def control_matrix(
        records: pd.DataFrame, row_field: str, limit: int
    ) -> tuple[pd.DataFrame, tuple[str, ...]]:
        matrix = matrix_table(
            records,
            row_field,
            "AgingControle",
            limit=limit,
            bucket_labels=PBIX_ENTRY_AGE_BUCKETS,
        )
        empty_buckets = [
            bucket
            for bucket in PBIX_ENTRY_AGE_BUCKETS
            if bucket in matrix.columns and not matrix[bucket].astype(int).sum()
        ]
        matrix = matrix.drop(columns=empty_buckets)
        numeric = tuple(
            [bucket for bucket in PBIX_ENTRY_AGE_BUCKETS if bucket in matrix.columns]
            + (["Total"] if "Total" in matrix.columns else [])
        )
        return matrix, numeric

    left, right = st.columns([1.15, 1], gap="small")
    with left:
        wait = df[df["STATUS_TRATATIVA"].eq("AGUARDANDO AGENDAMENTO")]
        wait_matrix, wait_numeric = control_matrix(wait, "SHIPPER", 11)
        table_panel("AGUARDANDO AGENDAMENTO", wait_matrix, height=206, numeric_columns=wait_numeric)
        shipped = df[df["STATUS_TRATATIVA"].eq("EXPEDIDO")]
        shipped_matrix, shipped_numeric = control_matrix(shipped, "SHIPPER", 11)
        table_panel("EXPEDIDO", shipped_matrix, height=198, numeric_columns=shipped_numeric)
    with right:
        status_matrix, status_numeric = control_matrix(status_rows, "STATUS_ATUAL", 8)
        table_panel(
            "STATUS",
            status_matrix,
            height=170,
            numeric_columns=status_numeric,
            cell_export_context=(status_rows, "STATUS_ATUAL", "AgingControle"),
        )
        scheduled = df[df["STATUS_TRATATIVA"].eq("AGENDADO")]
        scheduled_matrix, scheduled_numeric = control_matrix(scheduled, "SHIPPER", 5)
        table_panel("AGENDADO", scheduled_matrix, height=134, numeric_columns=scheduled_numeric)
        dimension = "CIDADE_ORIGEM" if df["CIDADE_ORIGEM"].notna().any() else "STATUS_ATUAL"
        consolidated_matrix, consolidated_numeric = control_matrix(df, dimension, 5)
        table_panel("CONSOLIDADO", consolidated_matrix, height=134, numeric_columns=consolidated_numeric)
    footer_detail(status_rows)


def control_stores_page() -> None:
    control_base = CONTROL_STORES.iloc[:0].copy()
    status_base = CONTROL_STORES.copy()
    status_base["AgingControle"] = powerbi_aging_bucket(
        status_base["DATA_EVENTO"], as_of=PBIX_ENTRY_AGE_AS_OF
    )
    top = st.columns([1, 1, 1.1, 1.1, 1, 1, 1], gap="small", vertical_alignment="center")
    store_filter_specs = [
        {"key": "cl_aging", "label": "AGING", "field": "AgingControle", "options": PBIX_ENTRY_AGE_BUCKETS},
        {"key": "cl_status", "label": "STATUS", "field": "STATUS_ATUAL"},
        {"key": "cl_emissor", "label": "LOJAS E PA'S", "field": "EMISSOR", "options": EMISSOR_FILTER_NAMES},
    ]
    store_periods = synchronized_periods(
        status_base,
        [
            {"key": "cl_data", "label": "DATA_EVENTO", "field": "DATA_ATUAL", "default": (date(2024, 7, 10), date(2026, 9, 7)), "include_missing_when_full": True},
            {"key": "cl_sol", "label": "DATA SOLICITAÇÃO", "field": "SOLICITACAO", "default": (date(2024, 7, 1), date(2026, 9, 3)), "include_missing_when_full": True},
        ],
        store_filter_specs,
        columns=[top[2], top[3]],
    )
    data_period = store_periods["cl_data"]
    sol_period = store_periods["cl_sol"]
    df = apply_period(control_base, "DATA_ATUAL", data_period)
    df = apply_period(df, "SOLICITACAO", sol_period)
    df["AgingControle"] = powerbi_aging_bucket(
        df["DATA_EVENTO"], as_of=PBIX_ENTRY_AGE_AS_OF
    )
    date_rows = apply_period(
        status_base, "DATA_ATUAL", data_period, include_missing_when_full=True
    )
    date_rows = apply_period(
        date_rows, "SOLICITACAO", sol_period, include_missing_when_full=True
    )
    synced = synchronized_multiselects(
        date_rows,
        store_filter_specs,
        [top[4], top[5], top[6]],
    )
    aging_filter = synced["cl_aging"]
    status_filter = synced["cl_status"]
    emitter_filter = synced["cl_emissor"]
    selected_dimensions = {
        "AgingControle": aging_filter,
        "STATUS_ATUAL": status_filter,
        "EMISSOR": emitter_filter,
    }
    detail_rows = date_rows
    df = filtered(df, selected_dimensions)
    metric_cards(top[:2], detail_rows)
    a, b = st.columns([1.25, 1.0], gap="small")
    with a:
        waiting = df[df["STATUS_TRATATIVA"].eq("LOJAS -TRANSF")]
        table_panel("AGUARDANDO RETIRADA", matrix_table(waiting, "EMISSOR", "AgingEntrada", limit=27), height=477, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    with b:
        shipped = df[df["STATUS_TRATATIVA"].eq("EXPEDIDO")]
        table_panel("EXPEDIDO", matrix_table(shipped, "EMISSOR", "AgingExpedição", limit=25), height=477, numeric_columns=tuple(BUCKET_LABELS + ["Total"]))
    footer_detail(detail_rows)


def faixa_valor_page() -> None:
    analysis_base = CLASSIFIED_BASE
    if analysis_base.empty:
        st.info("Não há pedidos classificados como Cliente ou Lojas e PAs nesta base.")
        return
    top = st.columns([1, 1, 1.05, .9, .9, .9, .75, .75], gap="small", vertical_alignment="center")
    metric_cards(top[:2], analysis_base)
    faixa_filter_specs = [
        {"key": "fv_faixa_valor", "label": "FAIXA DE VALOR", "field": "Faixa_Valor", "options": list(reversed(VALUE_BANDS)), "placeholder": "Todas"},
        {"key": "fv_emissor", "label": "LOJAS E PA'S", "field": "EMISSOR", "options": EMISSOR_FILTER_NAMES, "selection_group": "party_type", "classification_type": "Lojas e PAs"},
        {"key": "fv_shipper", "label": "CLIENTES", "field": "SHIPPER", "options": SHIPPER_NAMES, "selection_group": "party_type", "classification_type": "Cliente"},
        {"key": "fv_tipo", "label": "TIPO", "field": "TIPO", "options": category_options("TIPO", analysis_base)},
        {"key": "fv_vencimento", "label": "VENCIMENTO", "kind": "vencimento"},
    ]
    faixa_periods = synchronized_periods(
        analysis_base,
        [{"key": "fv_data", "label": "DATA_EVENTO", "field": "DATA_ATUAL"}],
        faixa_filter_specs,
        [top[2]],
    )
    date_filter = faixa_periods["fv_data"]
    date_base = apply_period(analysis_base.copy(), "DATA_ATUAL", date_filter)
    present_value_bands = set(category_options("Faixa_Valor", date_base))
    value_band_options = [band for band in reversed(VALUE_BANDS) if band in present_value_bands]
    faixa_filter_specs[3]["options"] = category_options("TIPO", date_base)
    synced = synchronized_multiselects(
        date_base,
        faixa_filter_specs,
        [top[3], top[4], top[5], top[6], top[7]],
    )
    df = date_base
    df = _apply_filter_specs(df, faixa_filter_specs, synced)
    a, b, c = st.columns([1.35, .95, .95], gap="small")
    with a:
        value_summary = summary_table(df, "Faixa_Valor", limit=8)
        value_order = {band: index for index, band in enumerate(value_band_options)}
        value_rows = value_summary.loc[value_summary["Faixa_Valor"].ne("Total")].copy()
        value_rows["_sort_order"] = value_rows["Faixa_Valor"].map(value_order)
        value_rows = value_rows.sort_values("_sort_order").drop(columns="_sort_order")
        total_row = value_summary.loc[value_summary["Faixa_Valor"].eq("Total")]
        value_summary = pd.concat([value_rows, total_row], ignore_index=True)
        table_panel("FAIXA DE VALOR", value_summary, height=225, dark_header=True, numeric_columns=("QTDE",), cell_export_context=(df, "Faixa_Valor", None))
    with b:
        client_records = df[df["TIPO"].eq("Cliente")]
        client_summary = summary_table(client_records, "SHIPPER", limit=10, display_field="CLIENTES")
        table_panel("CLIENTES", client_summary, height=225, dark_header=True, numeric_columns=("QTDE",), cell_export_context=(client_records, "SHIPPER", None))
    with c:
        store_records = df[df["TIPO"].eq("Lojas e PAs")]
        store_summary = summary_table(store_records, "EMISSOR", limit=10, display_field="LOJAS E PA'S")
        table_panel("LOJAS E PA'S", store_summary, height=225, dark_header=True, numeric_columns=("QTDE",), cell_export_context=(store_records, "EMISSOR", None))
    d, e = st.columns([1.15, 1], gap="small")
    with d:
        status_matrix = matrix_table(df, "STATUS_ATUAL", "AgingEntrada", limit=9)
        table_panel("STATUS", status_matrix, height=235, numeric_columns=tuple(BUCKET_LABELS + ["Total"]), cell_export_context=(df, "STATUS_ATUAL", "AgingEntrada"))
    with e:
        occurrence_matrix = matrix_table(df, "DESCRICAO_ATUAL", "AgingEntrada", limit=9)
        table_panel("OCORRÊNCIA", occurrence_matrix, height=235, numeric_columns=tuple(BUCKET_LABELS + ["Total"]), cell_export_context=(df, "DESCRICAO_ATUAL", "AgingEntrada"))
    footer_detail(df if not df.empty else analysis_base)


def lista_cte_page() -> None:
    sidebar, main = st.columns([.78, 2.2], gap="small")
    with sidebar:
        with st.container(border=True):
            st.markdown('<div class="panel-title" style="font-size:17px;text-align:center">FILTROS</div>', unsafe_allow_html=True)
            # Lista CTE exibe a base integral; TIPO permite alternar entre as
            # três classificações exclusivas, incluindo os pedidos em Rua.
            df = BASE.copy()
            lista_filter_specs = [
                {"key": "lista_vencimento", "label": "VENCIMENTO", "kind": "vencimento"},
                {"key": "lista_SHIPPER", "label": "CLIENTES", "field": "SHIPPER", "options": category_options("SHIPPER", BASE), "selection_group": "party_type", "classification_type": "Cliente"},
                {"key": "lista_EMISSOR", "label": "LOJAS E PA'S", "field": "EMISSOR", "options": category_options("EMISSOR", BASE), "selection_group": "party_type", "classification_type": "Lojas e PAs"},
                {"key": "lista_STATUS_ATUAL", "label": "STATUS ATUAL", "field": "STATUS_ATUAL", "options": category_options("STATUS_ATUAL", BASE)},
                {"key": "lista_TIPO", "label": "TIPO", "field": "TIPO", "options": ["Cliente", "Lojas e PAs", "Rua"]},
                {"key": "lista_MANIFESTO", "label": "MANIFESTO", "field": "MANIFESTO", "options": ["(Em branco)", *sorted(BASE["MANIFESTO"].dropna().astype(str).unique())[:30]], "include_blank": True, "placeholder": "Pesquisar"},
                {"key": "lista_PROXIMA_UNIDADE", "label": "PROXIMA_UNIDADE", "field": "PROXIMA_UNIDADE", "options": category_options("PROXIMA_UNIDADE", BASE)},
                {"key": "lista_Faixa_Valor", "label": "Faixa_Valor", "field": "Faixa_Valor", "options": VALUE_BANDS},
            ]
            lista_periods = synchronized_periods(
                df,
                [
                    {"key": "lista_sol", "label": "DATA SOLICITAÇÃO", "field": "SOLICITACAO", "include_missing_when_full": True},
                    {"key": "lista_emissao", "label": "DATA EMISSÃO", "field": "DT_EMISSAO", "include_missing_when_full": True},
                ],
                lista_filter_specs,
            )
            df = apply_period(df, "SOLICITACAO", lista_periods["lista_sol"], include_missing_when_full=True)
            df = apply_period(df, "DT_EMISSAO", lista_periods["lista_emissao"], include_missing_when_full=True)
            synced = synchronized_multiselects(
                df,
                lista_filter_specs,
            )
            df = _apply_filter_specs(df, lista_filter_specs, synced)
            search = st.text_input("PESQUISAR CTE / PEDIDO", key="lista_busca", placeholder="Pesquisar")
            if search.strip():
                mask = df["CTE"].astype(str).str.contains(search, case=False) | df["PEDIDO"].astype(str).str.contains(search, case=False)
                df = df[mask]
    with main:
        with st.container(border=True):
            st.markdown(f'<div class="panel-title">RESULTADO DA CONSULTA · {number_br(len(df))} REGISTROS</div>', unsafe_allow_html=True)
            view = detail_export_frame(df)
            detail_column_config: dict[str, object] = {
                "CTE": st.column_config.TextColumn("CTE", width="medium"),
                "PEDIDO": st.column_config.TextColumn("PEDIDO", width="medium"),
                "TIPO": st.column_config.TextColumn("TIPO", width="medium"),
                "AGING": st.column_config.TextColumn("AGING", width="small"),
            }
            for column in view.columns:
                if pd.api.types.is_datetime64_any_dtype(view[column]):
                    detail_column_config[column] = st.column_config.DateColumn(
                        column, format="DD/MM/YYYY", width="small"
                    )
            st.dataframe(
                view,
                use_container_width=True,
                hide_index=True,
                height=535,
                column_config=detail_column_config,
            )
with st.sidebar:
    st.markdown("### Páginas")
    selected_page = st.radio("Navegação do relatório", PAGES, index=0, label_visibility="collapsed", key="report_page")
    st.caption(f"Base ativa: {active_source} · {number_br(len(BASE))} CTEs")

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
