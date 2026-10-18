"""Leitura e preparação da base operacional fornecida em data.xlsx."""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd


AGE_BUCKETS = ["D+1", "D+3", "D+4", "D+5", "D+6-10", "D+11-15", "D+16-30", "D+30+", "Sem dados"]
PBIX_ENTRY_AGE_BUCKETS = [
    "D+0", "D+1", "D+2", "D+3", "D+4", "D+5",
    "D+6-10", "D+11-15", "D+16-30", "D+30+",
]
PBIX_ENTRY_AGE_AS_OF = pd.Timestamp("2026-09-08")
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
    serial = pd.to_numeric(series, errors="coerce")
    excel_dates = pd.to_datetime(serial, unit="D", origin="1899-12-30", errors="coerce")
    text_dates = pd.to_datetime(series, dayfirst=True, errors="coerce")
    return excel_dates.where(serial.notna(), text_dates)


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
    """Apply the PBIX AgingEntrada DAX: DATEDIFF(DATA_ATUAL, TODAY(), DAY)."""
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


def load_source_data(path: str | Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Read the Excel export and return the full base plus its client/store slices."""
    source_path = Path(path)
    if not source_path.is_file():
        raise FileNotFoundError(f"Base de dados não encontrada: {source_path}")

    frame = pd.read_excel(source_path, sheet_name="Export", engine="openpyxl")
    missing = [column for column in SOURCE_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError("Colunas ausentes na aba Export: " + ", ".join(missing))

    # These model fields are not present in the original export. The app
    # workbook adds the PBIX values, matched one-to-one by CTE.
    model_data_actual = (
        pd.to_datetime(frame["DATA_ATUAL"], dayfirst=True, errors="coerce")
        if "DATA_ATUAL" in frame.columns
        else None
    )
    model_city_origin = (
        frame["CIDADE_ORIGEM"].astype("string").str.strip().replace("", pd.NA)
        if "CIDADE_ORIGEM" in frame.columns
        else None
    )
    frame = frame.loc[:, SOURCE_COLUMNS].copy()
    for column in [
        "STATUS_ATUAL", "DESCRICAO_ATUAL", "STATUS_TRATATIVA", "EMISSOR",
        "SOLICITANTE", "PROXIMA_UNIDADE", "SHIPPER", "TIPO",
    ]:
        frame[column] = frame[column].astype("string").str.strip().replace("", pd.NA)

    for column in ["CTE", "PEDIDO", "MANIFESTO"]:
        frame[column] = _identifier_text(frame[column])

    frame["SOLICITACAO"] = _excel_date(frame["SOLICITACAO"])
    frame["DT_EMISSAO"] = _excel_date(frame["DT_EMISSAO"])
    frame["DT_ENTRADA_REVERSAS"] = pd.to_datetime(
        frame["DT_ENTRADA_REVERSAS"], dayfirst=True, errors="coerce"
    )
    frame["VALOR_NF"] = pd.to_numeric(frame["VALOR_NF"], errors="coerce")

    # Retain PBIX model fields when present; otherwise use the closest fields
    # available in the source export.
    frame["DATA_ATUAL"] = (
        model_data_actual
        if model_data_actual is not None
        else frame["DT_ENTRADA_REVERSAS"]
    )
    frame["DATA_EVENTO"] = frame["DT_ENTRADA_REVERSAS"]
    frame["CIDADE_ORIGEM"] = (
        model_city_origin
        if model_city_origin is not None
        else frame["STATUS_TRATATIVA"]
    )

    # The export has no explicit completion flag; terminal treatment statuses
    # are treated as finalized, while all other treatment statuses stay open.
    treatment = frame["STATUS_TRATATIVA"].str.upper()
    frame["FINALIZADO?"] = np.where(
        treatment.isin(["ENCERRADO", "BAIXADO"]), "FINALIZADO", "EM ABERTO"
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
