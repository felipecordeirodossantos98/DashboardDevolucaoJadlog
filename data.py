"""Leitura e preparação da base operacional fornecida em data.xlsx."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd


AGE_BUCKETS = ["D+1", "D+3", "D+4", "D+5", "D+6-10", "D+11-15", "D+16-30", "D+30+", "Sem dados"]
PBIX_ENTRY_AGE_BUCKETS = [
    "D+0", "D+1", "D+2", "D+3", "D+4", "D+5",
    "D+6-10", "D+11-15", "D+16-30", "D+30+",
]
PBIX_ENTRY_AGE_AS_OF: pd.Timestamp | None = None
VALUE_BANDS = ["<300", "300~500", "500~700", "700~1000", "1000~2000", "2000~5000", "5000~10000", ">10000"]

SOURCE_COLUMNS = [
    "CTE", "PEDIDO", "STATUS_ATUAL", "SOLICITACAO", "DESCRICAO_ATUAL",
    "STATUS_TRATATIVA", "EMISSOR", "SOLICITANTE", "PROXIMA_UNIDADE",
    "SHIPPER", "VALOR_NF", "TIPO", "DT_EMISSAO", "MANIFESTO",
    "DT_ENTRADA_REVERSAS",
]


def _identifier_text(series: pd.Series) -> pd.Series:
    """Preserve identifiers as text and remove Excel's trailing .0 when present."""
    def normalize(value: object) -> object:
        if pd.isna(value):
            return pd.NA
        if isinstance(value, (int, np.integer)):
            return str(int(value))
        if isinstance(value, (float, np.floating)) and np.isfinite(value) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    return series.map(normalize).astype("string")


def _excel_date(series: pd.Series) -> pd.Series:
    text_dates = pd.to_datetime(series, dayfirst=True, errors="coerce")
    text_values = series.astype("string").str.strip()
    is_serial = text_values.str.fullmatch(r"\d+(?:\.\d+)?", na=False)
    if not is_serial.any():
        return text_dates
    serial = pd.to_numeric(series.where(is_serial), errors="coerce")
    excel_dates = pd.to_datetime(
        serial, unit="D", origin="1899-12-30", errors="coerce"
    )
    return text_dates.where(~is_serial, excel_dates)


def _age_bucket(days: pd.Series) -> pd.Series:
    numeric_days = pd.to_numeric(days, errors="coerce")
    numeric_days = numeric_days.where(numeric_days.ge(0))
    buckets = pd.cut(
        numeric_days,
        bins=[-np.inf, 1, 3, 4, 5, 10, 15, 30, np.inf],
        labels=AGE_BUCKETS[:-1],
        right=True,
    )
    return buckets.astype("string").fillna("Sem dados")


def powerbi_aging_bucket(
    data_actual: pd.Series, *, as_of: pd.Timestamp | None = None
) -> pd.Series:
    """Apply the PBIX aging rule: DATEDIFF(source_date, TODAY(), DAY)."""
    reference_date = (
        pd.Timestamp.today().normalize()
        if as_of is None
        else pd.Timestamp(as_of).normalize()
    )
    current_dates = pd.to_datetime(data_actual, errors="coerce").dt.normalize()
    days = (reference_date - current_dates).dt.days
    conditions = [
        days.eq(0),
        days.eq(1),
        days.eq(2),
        days.eq(3),
        days.eq(4),
        days.eq(5),
        days.between(6, 10),
        days.between(11, 15),
        days.between(16, 30),
    ]
    return pd.Series(
        np.select(conditions, PBIX_ENTRY_AGE_BUCKETS[:-1], default="D+30+"),
        index=data_actual.index,
        dtype="string",
    )


def load_source_data(
    source: str | Path | bytes | bytearray | BytesIO,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load either supported Jadlog workbook and normalize it for the dashboard."""
    if isinstance(source, (bytes, bytearray)):
        workbook: str | Path | BytesIO = BytesIO(source)
    elif hasattr(source, "read"):
        workbook = source  # type: ignore[assignment]
    else:
        source_path = Path(source)
        if not source_path.is_file():
            raise FileNotFoundError(f"Base de dados não encontrada: {source_path}")
        workbook = source_path

    try:
        excel_file = pd.ExcelFile(workbook, engine="openpyxl")
        sheet_name = "Export" if "Export" in excel_file.sheet_names else excel_file.sheet_names[0]
        frame = pd.read_excel(excel_file, sheet_name=sheet_name)
    except (ValueError, IndexError) as exc:
        raise ValueError("O arquivo não contém uma planilha com dados para importar.") from exc

    frame.columns = [str(column).strip() for column in frame.columns]
    # Map the BI export names to the dashboard's canonical dimensions.
    # Prefer CLIENTE(S) over the older SHIPPER field when both are present.
    client_column = next(
        (name for name in ("CLIENTE", "CLIENTES") if name in frame.columns),
        None,
    )
    if client_column is not None:
        frame["SHIPPER"] = frame[client_column]
    aliases = {
        "SOLICITACAO": ("DATA_SOLICITACAO", "DATA_SOLICITAÇÃO"),
        "DESCRICAO_ATUAL": ("OCORRENCIA",),
        "SHIPPER": ("CLIENTE", "CLIENTES"),
        "EMISSOR": ("LOJAS E PA'S", "LOJAS E PAS"),
        "DT_ENTRADA_REVERSAS": ("DATA_ULTIMA_DATA", "DATA_EVENTO"),
    }
    for target, alternatives in aliases.items():
        if target not in frame.columns:
            alternate = next((name for name in alternatives if name in frame.columns), None)
            if alternate is not None:
                frame[target] = frame[alternate]

    missing = [column for column in ("CTE", "STATUS_ATUAL") if column not in frame.columns]
    if missing:
        raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(missing))

    # Exports from Power BI may contain a totals row, a blank row, and a filter
    # description after the detail records. Keep only rows with a CTE value.
    cte_values = _identifier_text(frame["CTE"])
    cte_labels = cte_values.str.strip().str.casefold()
    keep_rows = (
        cte_values.notna()
        & cte_values.str.strip().ne("")
        & cte_labels.ne("total")
        & ~cte_labels.str.startswith("filtros aplicados", na=False)
    )
    frame = frame.loc[keep_rows].copy()
    frame["CTE"] = cte_values.loc[keep_rows]
    if frame.empty:
        raise ValueError("Nenhum registro com CTE foi encontrado no arquivo.")

    actual_column = next(
        (name for name in ("DATA_ATUAL", "DATA_ULTIMA_DATA", "DT_ENTRADA_REVERSAS") if name in frame),
        None,
    )
    model_data_actual = _excel_date(frame[actual_column]) if actual_column else None
    model_city_origin = frame["CIDADE_ORIGEM"] if "CIDADE_ORIGEM" in frame else None
    model_treatment = frame["STATUS_TRATATIVA"] if "STATUS_TRATATIVA" in frame else None
    source_type = frame["TIPO"] if "TIPO" in frame else None
    finalized_source = (
        frame["FINALIZADO?"]
        if "FINALIZADO?" in frame
        else frame["FINALIZADO"] if "FINALIZADO" in frame else None
    )

    for column in SOURCE_COLUMNS:
        if column not in frame.columns:
            frame[column] = pd.NA
    frame = frame.loc[:, SOURCE_COLUMNS].copy()
    for column in [
        "STATUS_ATUAL", "DESCRICAO_ATUAL", "STATUS_TRATATIVA", "EMISSOR",
        "SOLICITANTE", "PROXIMA_UNIDADE", "SHIPPER", "TIPO",
    ]:
        frame[column] = frame[column].astype("string").str.strip().replace("", pd.NA)

    for column in ["PEDIDO", "MANIFESTO"]:
        frame[column] = _identifier_text(frame[column])

    frame["SOLICITACAO"] = _excel_date(frame["SOLICITACAO"])
    frame["DT_EMISSAO"] = _excel_date(frame["DT_EMISSAO"])
    frame["DT_ENTRADA_REVERSAS"] = _excel_date(frame["DT_ENTRADA_REVERSAS"])
    frame["VALOR_NF"] = pd.to_numeric(frame["VALOR_NF"], errors="coerce")

    # Retain PBIX fields where the export has them. Newer exports call the
    # event date DATA_ULTIMA_DATA and the client/occurrence fields CLIENTE and
    # OCORRENCIA, so the aliases above preserve those values.
    frame["DATA_ATUAL"] = (
        model_data_actual
        if model_data_actual is not None
        else frame["DT_ENTRADA_REVERSAS"]
    )
    frame["DATA_EVENTO"] = frame["DT_ENTRADA_REVERSAS"].where(
        frame["DT_ENTRADA_REVERSAS"].notna(), frame["DATA_ATUAL"]
    )
    frame["CIDADE_ORIGEM"] = (
        model_city_origin.astype("string").str.strip().replace("", pd.NA)
        if model_city_origin is not None
        else pd.Series(pd.NA, index=frame.index, dtype="string")
    )
    frame["STATUS_TRATATIVA"] = (
        model_treatment.astype("string").str.strip().replace("", pd.NA)
        if model_treatment is not None
        else pd.Series(pd.NA, index=frame.index, dtype="string")
    )
    frame["TIPO"] = (
        source_type.astype("string").str.strip().replace("", pd.NA).fillna("Não classificado")
        if source_type is not None
        else pd.Series("Não classificado", index=frame.index, dtype="string")
    )

    # Normalize both the old boolean-like completion field and the newer
    # FINALIZADO value into the same labels used by the dashboard.
    treatment = frame["STATUS_TRATATIVA"].str.upper()
    terminal_by_treatment = treatment.isin(["ENCERRADO", "BAIXADO"])
    if finalized_source is None:
        frame["FINALIZADO?"] = np.where(terminal_by_treatment, "FINALIZADO", "EM ABERTO")
    else:
        finalized_values = finalized_source.astype("string").str.strip().str.upper()
        finalized_yes = finalized_values.isin(
            ["SIM", "S", "YES", "TRUE", "1", "FINALIZADO", "ENCERRADO", "BAIXADO"]
        )
        finalized_no = finalized_values.isin(
            ["NAO", "NÃO", "N", "NO", "FALSE", "0", "EM ABERTO", "ABERTO"]
        )
        frame["FINALIZADO?"] = np.select(
            [finalized_yes, finalized_no, terminal_by_treatment],
            ["FINALIZADO", "EM ABERTO", "FINALIZADO"],
            default="EM ABERTO",
        )

    # Aging is calculated from dates present in the export, using the elapsed
    # time between the source milestones rather than generated demo values.
    entry_date = frame["DT_ENTRADA_REVERSAS"].dt.normalize()
    request_date = frame["SOLICITACAO"].dt.normalize()
    emission_date = frame["DT_EMISSAO"].dt.normalize()
    frame["AgingEntrada"] = _age_bucket((entry_date - request_date).dt.days)
    frame["AgingExpedição"] = _age_bucket((entry_date - emission_date).dt.days)
    frame["AgingSaida"] = frame["AgingExpedição"]
    frame["AgingOrder"] = (entry_date - request_date).dt.days
    frame["AgingExpediçãoOrder"] = (entry_date - emission_date).dt.days

    frame["Faixa_Valor"] = pd.cut(
        frame["VALOR_NF"],
        bins=[-np.inf, 300, 500, 700, 1000, 2000, 5000, 10000, np.inf],
        labels=VALUE_BANDS,
        right=False,
    ).astype("string")

    clients = frame.loc[frame["TIPO"].eq("Cliente")].copy()
    stores = frame.loc[frame["TIPO"].eq("Lojas e PAs")].copy()
    return frame, clients, stores
