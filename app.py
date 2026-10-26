from __future__ import annotations

import csv
import hmac
import os
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


REQUIRED_COLUMNS = {
    "Date",
    "Quantity",
    "Sales_Amount",
    "Transaction_Count",
    "Customer_Count",
}
LOCAL_CSV = Path(__file__).resolve().parent / "ventas_diarias.csv"
DEFAULT_CSV = LOCAL_CSV
MONTH_NAMES = (
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
)
WEEKDAY_NAMES = (
    "lunes",
    "martes",
    "miércoles",
    "jueves",
    "viernes",
    "sábado",
    "domingo",
)
VIEWS = (
    "Resumen",
    "Ventas por día",
    "Ventas por mes",
    "Ventas por año",
)

st.set_page_config(
    page_title="Saturno | Inteligencia de ventas",
    page_icon="🪐",
    layout="wide",
)

st.markdown(
    """
    <style>
    :root {
        --navy: #111827;
        --sidebar: #151d2d;
        --surface: #ffffff;
        --canvas: #f3f6fb;
        --ink: #14213d;
        --muted: #697586;
        --accent: #5468e8;
        --border: #e5eaf2;
    }
    .stApp, [data-testid="stAppViewContainer"] {
        background: var(--canvas);
        color: var(--ink);
    }
    [data-testid="stHeader"] { background: rgba(243, 246, 251, .95); }
    [data-testid="stSidebar"] {
        background: var(--sidebar);
        border-right: 1px solid rgba(255,255,255,.08);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] p { color: #e7ecf5; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: #ffffff; }
    [data-testid="stMetric"] {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 18px 20px;
        box-shadow: 0 5px 18px rgba(20, 33, 61, .05);
        min-height: 115px;
    }
    [data-testid="stMetricLabel"] p {
        color: var(--muted);
        font-size: .76rem;
        font-weight: 600;
    }
    [data-testid="stMetricValue"] p {
        color: #000000 !important;
        font-weight: 800 !important;
        font-size: clamp(.9rem, 1.4vw, 1.25rem);
        letter-spacing: -.04em;
        white-space: nowrap;
    }
    [data-testid="stDataFrame"] {
        border: 1px solid var(--border);
        border-radius: 12px;
        overflow: hidden;
    }
    [data-testid="stRadio"] label { font-weight: 600; }
    h1, h2, h3 { color: var(--ink); letter-spacing: -.025em; }
    .brand-lockup {
        display: flex;
        align-items: center;
        gap: 14px;
        margin: 0 0 22px;
    }
    .brand-lockup h1 { margin: 0; font-size: 1.8rem; }
    .brand-lockup p { margin: 2px 0 0; color: #697586; }
    .brand-logo {
        display: block;
        width: 58px;
        height: 58px;
        flex: 0 0 58px;
    }
    .section-kicker {
        color: var(--accent);
        font-size: .75rem;
        font-weight: 800;
        letter-spacing: .12em;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def excel_sheet_names(file_bytes: bytes) -> list[str]:
    with pd.ExcelFile(BytesIO(file_bytes)) as workbook:
        return workbook.sheet_names


@st.cache_data(show_spinner=False)
def load_data(
    file_bytes: bytes, filename: str, sheet_name: str | None = None
) -> tuple[pd.DataFrame, str]:
    extension = Path(filename).suffix.lower()
    if extension == ".csv":
        data = read_csv_file(file_bytes)
    elif extension in {".xlsx", ".xls"}:
        data = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name)
    else:
        raise ValueError("Formato no compatible. Carga un archivo CSV o Excel.")
    if data.empty:
        raise ValueError("La hoja seleccionada no contiene registros.")
    data.attrs["filename"] = filename
    return clean_missing_values(data), filename


def read_csv_file(file_bytes: bytes) -> pd.DataFrame:
    text = None
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            text = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("No se pudo reconocer la codificación del archivo CSV.")

    sample = text[:8192]
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        first_line = sample.splitlines()[0] if sample.splitlines() else ""
        delimiter = max(",;\t|", key=first_line.count)

    return pd.read_csv(
        BytesIO(text.encode("utf-8")),
        sep=delimiter,
        skipinitialspace=True,
    )


def clean_missing_values(data: pd.DataFrame) -> pd.DataFrame:
    data = data.copy()
    data.columns = [str(column).strip() for column in data.columns]
    if len(set(data.columns)) != len(data.columns):
        raise ValueError(
            "El archivo tiene nombres de columna duplicados o iguales "
            "después de quitar espacios."
        )

    data = data.replace(r"^\s*$", pd.NA, regex=True)
    empty_columns = int(data.isna().all(axis=0).sum())
    data = data.dropna(axis=1, how="all")
    incomplete_rows = int(data.isna().any(axis=1).sum())
    data = data.dropna(axis=0, how="any").copy()
    data.attrs["removed_rows"] = incomplete_rows
    data.attrs["removed_columns"] = empty_columns
    if data.empty:
        raise ValueError(
            "No quedan filas completas después de quitar los valores vacíos."
        )
    return data


def prepare_sales_data(data: pd.DataFrame) -> pd.DataFrame:
    raw_dates = data["Date"].astype("string")
    has_time_data = raw_dates.str.contains(
        r"\d{1,2}:\d{2}", regex=True, na=False
    ).any()
    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    for column in (
        "Quantity",
        "Sales_Amount",
        "Transaction_Count",
        "Customer_Count",
    ):
        data[column] = pd.to_numeric(data[column], errors="coerce")

    valid = data[list(REQUIRED_COLUMNS)].notna().all(axis=1)
    invalid_rows = int((~valid).sum())
    data = data.loc[valid].copy()
    if data.empty:
        raise ValueError("No hay filas de ventas con fechas y datos válidos.")
    data.attrs["invalid_rows"] = data.attrs.get("removed_rows", 0) + invalid_rows
    data.attrs["removed_columns"] = data.attrs.get("removed_columns", 0)
    data.attrs["has_time_data"] = bool(has_time_data)
    return data


def render_generic_data(data: pd.DataFrame, filename: str) -> None:
    st.markdown(
        """
        <div class="brand-lockup">
          <div>
            <div class="section-kicker">Explorador de datos</div>
            <h1>Archivo listo para explorar</h1>
            <p>La app leyó el archivo y adaptó el análisis a sus columnas.</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"Archivo: {filename}")

    metric_columns = st.columns(3)
    metric_columns[0].metric("Registros", format_number(len(data)))
    metric_columns[1].metric("Columnas", format_number(len(data.columns)))
    metric_columns[2].metric(
        "Filas incompletas quitadas",
        format_number(data.attrs.get("removed_rows", 0)),
    )
    if data.attrs.get("removed_columns", 0):
        st.info(
            f"Se quitaron {data.attrs['removed_columns']:,} columnas totalmente vacías."
        )

    preview_tab, columns_tab, explore_tab = st.tabs(
        ["Vista previa", "Columnas", "Explorar campo"]
    )
    with preview_tab:
        st.subheader("Datos leídos")
        st.dataframe(data.head(1000), width="stretch", hide_index=True)
        st.download_button(
            "Descargar datos como CSV",
            data=data.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"{Path(filename).stem}_datos.csv",
            mime="text/csv",
        )

    with columns_tab:
        st.subheader("Estructura del archivo")
        column_summary = pd.DataFrame(
            {
                "Columna": data.columns.astype(str),
                "Tipo de dato": [str(dtype) for dtype in data.dtypes],
                "Valores presentes": [
                    int(data[column].notna().sum()) for column in data.columns
                ],
                "Valores vacíos": [
                    int(data[column].isna().sum()) for column in data.columns
                ],
                "Valores distintos": [
                    int(data[column].nunique(dropna=True))
                    for column in data.columns
                ],
            }
        )
        st.dataframe(column_summary, width="stretch", hide_index=True)
        numeric_columns = data.select_dtypes(include="number").columns.tolist()
        if numeric_columns:
            st.subheader("Resumen de campos numéricos")
            st.dataframe(
                data[numeric_columns].describe().transpose(),
                width="stretch",
            )

    with explore_tab:
        st.subheader("Explora una columna")
        selected_column = st.selectbox(
            "Columna",
            data.columns.tolist(),
            format_func=str,
        )
        values = data[selected_column].dropna()
        if values.empty:
            st.info("Esta columna no tiene valores para visualizar.")
        elif pd.api.types.is_numeric_dtype(values):
            st.metric(
                f"Promedio · {selected_column}",
                format_number(float(values.mean()), 2),
            )
            st.plotly_chart(
                style_chart(
                    px.histogram(
                        data,
                        x=selected_column,
                        labels={str(selected_column): str(selected_column)},
                        color_discrete_sequence=["#5468e8"],
                    )
                ),
                width="stretch",
            )
        elif pd.api.types.is_datetime64_any_dtype(values):
            date_counts = (
                values.dt.to_period("M")
                .value_counts()
                .sort_index()
                .rename_axis("Periodo")
                .rename("Registros")
                .reset_index()
            )
            date_counts["Periodo"] = date_counts["Periodo"].astype(str)
            st.plotly_chart(
                style_chart(
                    px.line(
                        date_counts,
                        x="Periodo",
                        y="Registros",
                        markers=True,
                        color_discrete_sequence=["#5468e8"],
                    )
                ),
                width="stretch",
            )
        else:
            value_counts = (
                values.astype("string")
                .value_counts()
                .head(15)
                .sort_values()
                .rename_axis("Valor")
                .rename("Registros")
                .reset_index()
            )
            st.plotly_chart(
                style_chart(
                    px.bar(
                        value_counts,
                        x="Registros",
                        y="Valor",
                        orientation="h",
                        color_discrete_sequence=["#18a88b"],
                    )
                ),
                width="stretch",
            )


def format_number(value: float, decimals: int = 0) -> str:
    return f"{value:,.{decimals}f}"


def find_default_csv() -> Path | None:
    for candidate in (LOCAL_CSV, DEFAULT_CSV):
        if candidate.is_file():
            return candidate
    return None


def configured_app_password() -> str | None:
    password = os.environ.get("SATURNO_APP_PASSWORD")
    if password and password.strip():
        return password
    try:
        password = st.secrets.get("app", {}).get("password")
    except FileNotFoundError:
        return None
    return (
        password
        if isinstance(password, str) and password.strip()
        else None
    )


def require_login() -> None:
    expected_password = configured_app_password()
    if expected_password is None:
        st.title("Acceso protegido")
        st.error("La contraseña de acceso todavía no está configurada.")
        st.markdown(
            "Antes de iniciar la app, configura la variable de entorno "
            "`SATURNO_APP_PASSWORD`. Para Streamlit Community Cloud, agrega "
            "el secreto `[app]` con la clave `password` en la configuración "
            "de Secrets. No guardes esa contraseña en el código ni en Git."
        )
        st.stop()

    if not st.session_state.get("authenticated", False):
        st.markdown(
            """
            <div class="brand-lockup">
              <div>
                <div class="section-kicker">Acceso privado</div>
                <h1>Bienvenido a Saturno</h1>
                <p>Ingresa la contraseña para abrir el panel de ventas.</p>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form", clear_on_submit=True):
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button(
                "Entrar",
                type="primary",
                width="stretch",
            )
        if submitted:
            matches = hmac.compare_digest(
                password.encode("utf-8"),
                expected_password.encode("utf-8"),
            )
            if matches:
                st.session_state["authenticated"] = True
                st.rerun()
            st.error("La contraseña no es correcta.")
        st.stop()


def month_label(value: pd.Timestamp) -> str:
    return f"{MONTH_NAMES[value.month - 1].capitalize()} {value.year}"


def style_chart(chart: object) -> object:
    chart.update_layout(
        margin=dict(l=12, r=18, t=28, b=12),
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=dict(color="#111827", family="Arial, sans-serif"),
        hoverlabel=dict(
            bgcolor="#111827",
            font=dict(color="#ffffff", family="Arial, sans-serif"),
        ),
        xaxis=dict(
            tickfont=dict(color="#111827"),
            title_font=dict(color="#111827"),
            showgrid=False,
            linecolor="#dce3ed",
        ),
        yaxis=dict(
            tickfont=dict(color="#111827"),
            title_font=dict(color="#111827"),
            gridcolor="#edf0f5",
            linecolor="#dce3ed",
        ),
    )
    return chart


def render_metrics(data: pd.DataFrame) -> None:
    revenue = data["Sales_Amount"].sum()
    transactions = data["Transaction_Count"].sum()
    customers = data["Customer_Count"].sum()
    average_ticket = revenue / transactions if transactions else 0
    columns = st.columns(4)
    columns[0].metric("Ventas totales", format_number(revenue, 2))
    columns[1].metric("Transacciones", format_number(transactions))
    columns[2].metric("Cliente-días", format_number(customers))
    columns[3].metric("Ticket promedio", format_number(average_ticket, 2))


def daily_summary(data: pd.DataFrame) -> pd.DataFrame:
    summary = (
        data.assign(Fecha=data["Date"].dt.normalize())
        .groupby("Fecha", as_index=False)
        .agg(
            Ventas=("Sales_Amount", "sum"),
            Transacciones=("Transaction_Count", "sum"),
        )
        .sort_values("Fecha")
    )
    summary.insert(
        1,
        "Día",
        summary["Fecha"].dt.day_of_week.map(
            lambda day: WEEKDAY_NAMES[day].capitalize()
        ),
    )
    return summary


def monthly_summary(data: pd.DataFrame) -> pd.DataFrame:
    summary = (
        data.assign(Mes=data["Date"].dt.to_period("M").dt.to_timestamp())
        .groupby("Mes", as_index=False)
        .agg(
            Ventas=("Sales_Amount", "sum"),
            Transacciones=("Transaction_Count", "sum"),
            **{"Cliente-días": ("Customer_Count", "sum")},
        )
        .sort_values("Mes")
    )
    summary.insert(1, "Periodo", summary["Mes"].map(month_label))
    return summary


def yearly_summary(data: pd.DataFrame) -> pd.DataFrame:
    return (
        data.assign(Año=data["Date"].dt.year)
        .groupby("Año", as_index=False)
        .agg(
            Ventas=("Sales_Amount", "sum"),
            Transacciones=("Transaction_Count", "sum"),
            **{"Cliente-días": ("Customer_Count", "sum")},
        )
        .sort_values("Año")
    )


def show_aggregated_table(
    data: pd.DataFrame, date_columns: tuple[str, ...] = ()
) -> None:
    column_config = {
        "Ventas": st.column_config.NumberColumn("Ventas", format="%,.2f"),
        "Transacciones": st.column_config.NumberColumn(
            "Transacciones", format="%d"
        ),
        "Cliente-días": st.column_config.NumberColumn("Cliente-días", format="%d"),
    }
    for column in date_columns:
        column_config[column] = st.column_config.DateColumn(
            column, format="DD/MM/YYYY"
        )
    st.dataframe(
        data,
        width="stretch",
        hide_index=True,
        column_config=column_config,
    )
    st.download_button(
        "Descargar resumen CSV",
        data=data.to_csv(index=False).encode("utf-8-sig"),
        file_name="resumen_ventas.csv",
        mime="text/csv",
        key=f"download_{'_'.join(data.columns)}",
    )


require_login()

source_path = find_default_csv()
if source_path is None:
    st.error("No se encontró el archivo CSV de ventas configurado.")
    st.stop()
source_bytes = source_path.read_bytes()
source_name = source_path.name

with st.sidebar:
    st.markdown(
        """
        <div style="display:flex;align-items:center;gap:10px;margin:4px 0 24px;">
          <svg xmlns="http://www.w3.org/2000/svg" width="42" height="42"
               viewBox="0 0 124 124" aria-hidden="true">
            <circle cx="62" cy="62" r="60" fill="#ffffff"/>
            <ellipse cx="62" cy="62" rx="53" ry="19"
                     transform="rotate(-22 62 62)" fill="none"
                     stroke="#000000" stroke-width="6"/>
            <circle cx="62" cy="62" r="29" fill="#000000"/>
            <path d="M48 51l2 5 6 .5-4 4 1 5-5-3-5 3 1-5-4-4 6-.5z
                     M73 43l2 4 4 .3-3 3 1 4-4-2-4 2 1-4-3-3 4-.3z
                     M72 70l2 4 5 .3-4 3 1 5-4-3-4 3 1-5-4-3 5-.3z"
                  fill="#ffffff"/>
          </svg>
          <div><strong style="color:#fff;font-size:1.12rem;">SATURNO</strong>
          <div style="color:#9ca9bd;font-size:.72rem;letter-spacing:.08em;">
          VENTAS · ANALÍTICA</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("### Archivo activo")
    st.caption(source_name)

try:
    data, _ = load_data(source_bytes, source_name)
except (
    ValueError,
    pd.errors.ParserError,
    pd.errors.EmptyDataError,
    UnicodeDecodeError,
    ImportError,
    OSError,
) as error:
    st.error(f"No se pudo leer el archivo: {error}")
    st.stop()

if not REQUIRED_COLUMNS.issubset(data.columns):
    render_generic_data(data, source_name)
    st.stop()

try:
    sales = prepare_sales_data(data)
except ValueError as error:
    st.error(f"No se pudo interpretar el archivo como ventas: {error}")
    st.stop()

with st.sidebar:
    st.caption(f"{len(sales):,} registros válidos")
    st.divider()
    st.markdown("### Espacio de trabajo")
    view = st.radio("Análisis", VIEWS, label_visibility="collapsed")
    if st.button("Cerrar sesión", width="stretch"):
        st.session_state["authenticated"] = False
        st.rerun()
    st.divider()
    st.caption("La información se calcula a partir del archivo cargado.")

if sales.attrs["invalid_rows"] or sales.attrs["removed_columns"]:
    cleanup_details = []
    if sales.attrs["invalid_rows"]:
        cleanup_details.append(
            f"{sales.attrs['invalid_rows']:,} filas incompletas o inválidas"
        )
    if sales.attrs["removed_columns"]:
        cleanup_details.append(
            f"{sales.attrs['removed_columns']:,} columnas totalmente vacías"
        )
    st.warning(
        "Limpieza aplicada: " + " y ".join(cleanup_details) + "."
    )

st.markdown(
    """
    <div class="brand-lockup">
      <svg class="brand-logo" xmlns="http://www.w3.org/2000/svg"
           viewBox="0 0 124 124" role="img" aria-label="Saturno con estrellas">
        <circle cx="62" cy="62" r="60" fill="#ffffff"
                stroke="#e5eaf2" stroke-width="2"/>
        <ellipse cx="62" cy="62" rx="53" ry="19"
                 transform="rotate(-22 62 62)" fill="none"
                 stroke="#000000" stroke-width="6"/>
        <circle cx="62" cy="62" r="29" fill="#000000"/>
        <path d="M48 51l2 5 6 .5-4 4 1 5-5-3-5 3 1-5-4-4 6-.5z
                 M73 43l2 4 4 .3-3 3 1 4-4-2-4 2 1-4-3-3 4-.3z
                 M72 70l2 4 5 .3-4 3 1 5-4-3-4 3 1-5-4-3 5-.3z"
              fill="#ffffff"/>
      </svg>
      <div>
        <h1>Panel de ventas</h1>
        <p>Indicadores y tendencias para tomar mejores decisiones.</p>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption(
    f"Periodo analizado: {sales['Date'].min():%d/%m/%Y} — "
    f"{sales['Date'].max():%d/%m/%Y} · {len(sales):,} registros válidos"
)

st.markdown(
    f'<div class="section-kicker">{view}</div>',
    unsafe_allow_html=True,
)

if view == "Resumen":
    render_metrics(sales)
    st.subheader("Evolución mensual")
    months = monthly_summary(sales)
    chart = px.bar(
        months,
        x="Periodo",
        y="Ventas",
        labels={"Periodo": "Mes", "Ventas": "Ventas"},
        color_discrete_sequence=["#5468e8"],
        hover_data={"Transacciones": True, "Cliente-días": True},
    )
    chart.update_traces(
        marker_line_width=0,
        hovertemplate="%{x}<br>Ventas: %{y:,.2f}<extra></extra>",
    )
    chart.update_layout(showlegend=False)
    st.plotly_chart(style_chart(chart), width="stretch")

    daily = daily_summary(sales)
    best_day = daily.loc[daily["Ventas"].idxmax()]
    best_month = months.loc[months["Ventas"].idxmax()]
    insight_columns = st.columns(2)
    insight_columns[0].metric(
        "Día con más ventas",
        f"{best_day['Fecha']:%d/%m/%Y}",
        f"{format_number(best_day['Ventas'], 2)} en ventas",
    )
    insight_columns[1].metric(
        "Mes con más ventas",
        best_month["Periodo"],
        f"{format_number(best_month['Ventas'], 2)} en ventas",
    )

elif view == "Ventas por día":
    daily = daily_summary(sales)
    available_years = sorted(sales["Date"].dt.year.unique().tolist())
    selected_year = st.selectbox(
        "Año",
        available_years,
        key="daily_year",
    )
    selected_data = daily.loc[daily["Fecha"].dt.year == selected_year].copy()
    render_metrics(sales.loc[sales["Date"].dt.year == selected_year])
    st.subheader(f"Ventas diarias · {selected_year}")
    chart = px.line(
        selected_data,
        x="Fecha",
        y="Ventas",
        markers=True,
        labels={"Fecha": "Día", "Ventas": "Ventas"},
        color_discrete_sequence=["#5468e8"],
        hover_data={"Día": True, "Transacciones": True, "Ventas": ":,.2f"},
    )
    chart.update_traces(
        line=dict(width=3),
        marker=dict(size=6),
        hovertemplate="%{x|%d/%m/%Y}<br>Ventas: %{y:,.2f}<extra></extra>",
    )
    st.plotly_chart(style_chart(chart), width="stretch")

    if sales.attrs["has_time_data"]:
        st.subheader("Detalle por hora")
        selected_day = st.selectbox(
            "Día para consultar",
            selected_data["Fecha"].tolist(),
            format_func=lambda date: (
                f"{WEEKDAY_NAMES[date.dayofweek].capitalize()} "
                f"{date:%d/%m/%Y}"
            ),
            key="hourly_day",
        )
        same_day = sales.loc[
            (sales["Date"].dt.date == selected_day.date())
            & (sales["Date"].dt.year == selected_year)
        ].copy()
        hourly = (
            same_day.assign(Hora=same_day["Date"].dt.hour)
            .groupby("Hora", as_index=False)["Sales_Amount"]
            .sum()
            .rename(columns={"Sales_Amount": "Ventas"})
        )
        hourly["Hora"] = hourly["Hora"].map(lambda hour: f"{hour:02d}:00")
        hour_chart = px.bar(
            hourly,
            x="Hora",
            y="Ventas",
            labels={"Hora": "Hora", "Ventas": "Ventas"},
            color_discrete_sequence=["#18a88b"],
        )
        st.plotly_chart(style_chart(hour_chart), width="stretch")
    else:
        st.info(
            "El archivo contiene fechas, pero no horas. Para analizar ventas "
            "por hora, carga un CSV cuya columna Date incluya la hora "
            "(por ejemplo, 02/01/2016 14:30)."
        )

    st.subheader("Totales diarios")
    show_aggregated_table(selected_data, ("Fecha",))

elif view == "Ventas por mes":
    months = monthly_summary(sales)
    available_years = sorted(months["Mes"].dt.year.unique().tolist())
    selected_year = st.selectbox(
        "Año",
        available_years,
        key="monthly_year",
    )
    selected_months = months.loc[months["Mes"].dt.year == selected_year].copy()
    render_metrics(sales.loc[sales["Date"].dt.year == selected_year])
    st.subheader(f"Ventas mensuales · {selected_year}")
    chart = px.bar(
        selected_months,
        x="Periodo",
        y="Ventas",
        labels={"Periodo": "Mes", "Ventas": "Ventas"},
        color_discrete_sequence=["#5468e8"],
        hover_data={"Transacciones": True, "Cliente-días": True},
    )
    chart.update_traces(
        hovertemplate="%{x}<br>Ventas: %{y:,.2f}<extra></extra>"
    )
    st.plotly_chart(style_chart(chart), width="stretch")
    st.subheader("Totales mensuales")
    show_aggregated_table(
        selected_months.drop(columns="Mes"),
    )

elif view == "Ventas por año":
    years = yearly_summary(sales)
    render_metrics(sales)
    st.subheader("Ventas anuales")
    chart = px.bar(
        years,
        x="Año",
        y="Ventas",
        labels={"Año": "Año", "Ventas": "Ventas"},
        color_discrete_sequence=["#5468e8"],
        hover_data={"Transacciones": True, "Cliente-días": True},
    )
    chart.update_traces(
        hovertemplate="%{x}<br>Ventas: %{y:,.2f}<extra></extra>"
    )
    st.plotly_chart(style_chart(chart), width="stretch")
    st.subheader("Totales anuales")
    show_aggregated_table(years)
    if len(years) > 1:
        years["Cambio interanual (%)"] = years["Ventas"].pct_change() * 100
        st.subheader("Cambio respecto al año anterior")
        st.dataframe(
            years[["Año", "Cambio interanual (%)"]],
            width="stretch",
            hide_index=True,
            column_config={
                "Cambio interanual (%)": st.column_config.NumberColumn(
                    "Cambio interanual", format="%.2f%%"
                )
            },
        )

else:
    st.info(
        "El archivo de despliegue contiene únicamente totales diarios "
        "anonimizados; el detalle por categoría y producto no está incluido."
    )
