import streamlit as st
import pandas as pd
import numpy as np
import io
import re

st.set_page_config(
    page_title="Prime Day · Verificador Cluster A",
    page_icon="🔍",
    layout="wide",
)

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────

TABS_EUR   = {"ES-FR", "FR-FR", "ES-IT", "IT-IT", "ES-DE", "DE-DE", "PT", "NL", "BE"}
TABS_LOCAL = {"PL", "SE"}
ALL_TABS   = TABS_EUR | TABS_LOCAL

TAB_COUNTRY = {
    "ES-FR": "FR", "FR-FR": "FR",
    "ES-IT": "IT", "IT-IT": "IT",
    "ES-DE": "DE", "DE-DE": "DE",
    "PT": "PT", "NL": "NL", "BE": "BE",
    "PL": "PL", "SE": "SE",
}

DRAFT_COL_IDX = {
    "ES": 7,   # H
    "IT": 8,   # I
    "FR": 9,   # J
    "DE": 10,  # K
    "PL": 11,  # L
    "SE": 12,  # M
    "NL": 13,  # N
    "BE": 14,  # O
}

COUNTRIES_ORDER = ["ES", "FR", "IT", "DE", "PL", "SE", "NL", "BE"]

CURRENCY = {"PL": "PLN", "SE": "SEK"}

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def parse_price(val) -> float | None:
    """Parsea valores de precio: float, int, string con €/PLN/SEK/comas."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    s = re.sub(r"[€PLNSEKszł\s]", "", s)
    s = s.replace(",", ".")
    # eliminar puntos de miles si hay dos decimales al final
    if re.match(r"^\d{1,3}(\.\d{3})+\.\d{2}$", s):
        s = s.replace(".", "", s.count(".") - 1)
    try:
        return float(s)
    except ValueError:
        return None


def find_header_row(df_raw: pd.DataFrame, markers: list[str]) -> int:
    """Encuentra la fila de cabecera buscando columnas clave."""
    for i in range(min(10, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.upper().str.strip()
        if all(any(m.upper() in cell for cell in row) for m in markers):
            return i
    return 0


def read_excel_sheet(file_bytes, sheet_name=0) -> pd.DataFrame:
    return pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None,
        dtype=str,
        engine="openpyxl",
    )


def col_by_pattern(columns: list[str], pattern: str) -> int | None:
    """Devuelve el índice de la primera columna que cumple el patrón regex."""
    for i, c in enumerate(columns):
        if re.search(pattern, str(c), re.IGNORECASE):
            return i
    return None


# ─────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_draft(file_bytes: bytes) -> pd.DataFrame:
    raw = read_excel_sheet(file_bytes)
    hi  = find_header_row(raw, ["SKU", "CLUSTER"])
    raw.columns = raw.iloc[hi].astype(str).str.strip()
    df = raw.iloc[hi + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def load_nacional(file_bytes: bytes) -> dict[str, float]:
    raw = read_excel_sheet(file_bytes)
    hi  = find_header_row(raw, ["REFERENCIA", "PVP"])
    raw.columns = raw.iloc[hi].astype(str).str.strip()
    df = raw.iloc[hi + 1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    ref_col  = col_by_pattern(list(df.columns), r"^REFERENCIA$")
    pvp_col  = col_by_pattern(list(df.columns), r"PVP\s*PUB")

    if ref_col is None or pvp_col is None:
        st.error("Tarifa nacional: no se encontró columna REFERENCIA o PVP PUB.")
        return {}

    mapping = {}
    for _, row in df.iterrows():
        ref = str(row.iloc[ref_col]).strip()
        pvp = parse_price(row.iloc[pvp_col])
        if ref and ref not in ("nan", "") and pvp is not None:
            mapping[ref] = pvp
    return mapping


@st.cache_data(show_spinner=False)
def load_internacional(file_bytes: bytes) -> dict[str, dict]:
    """
    Devuelve {TAB_NAME_UPPER: {'map': {ref: price}, 'is_local': bool}}
    """
    xls    = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    result = {}

    for sn in xls.sheet_names:
        key = sn.strip().upper()
        if key not in ALL_TABS:
            continue
        is_local = key in TABS_LOCAL

        raw = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sn,
            header=None,
            dtype=str,
            engine="openpyxl",
        )
        hi = find_header_row(raw, ["REFERENCIA"])
        raw.columns = raw.iloc[hi].astype(str).str.strip()
        df = raw.iloc[hi + 1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        cols = list(df.columns)

        ref_col = col_by_pattern(cols, r"^REFERENCIA$")
        if ref_col is None:
            continue

        if is_local:
            # Col N: PVP (MONEDA LOCAL) — buscar patrón sin PUB
            price_col = col_by_pattern(cols, r"PVP\s*\(")
            if price_col is None:
                # fallback: cualquier PVP que no sea PUB ni MIN
                price_col = next(
                    (i for i, c in enumerate(cols)
                     if re.search(r"PVP", c, re.I)
                     and not re.search(r"PUB|MIN", c, re.I)),
                    None,
                )
        else:
            price_col = col_by_pattern(cols, r"PVP\s*PUB")

        if price_col is None:
            st.warning(f"Pestaña {sn}: no se encontró columna de precio. Se omite.")
            continue

        mapping = {}
        for _, row in df.iterrows():
            ref = str(row.iloc[ref_col]).strip()
            pvp = parse_price(row.iloc[price_col])
            if ref and ref not in ("nan", "") and pvp is not None:
                mapping[ref] = pvp

        result[key] = {"map": mapping, "is_local": is_local}

    return result


# ─────────────────────────────────────────────
# VERIFICACIÓN
# ─────────────────────────────────────────────

def run_check(draft_df: pd.DataFrame, nac_map: dict, inter_maps: dict) -> pd.DataFrame:
    cols = list(draft_df.columns)

    # Identificar columnas clave en el draft
    sku_col     = col_by_pattern(cols, r"^SKU$")
    cluster_col = col_by_pattern(cols, r"^CLUSTER$")
    subfam_col  = col_by_pattern(cols, r"SUBFAMILIA")
    asin_col    = col_by_pattern(cols, r"^ASIN$")

    if sku_col is None or cluster_col is None:
        st.error("El draft no tiene columnas SKU o CLUSTER reconocibles.")
        return pd.DataFrame()

    records = []

    for _, row in draft_df.iterrows():
        cluster = str(row.iloc[cluster_col]).strip().upper()
        if cluster != "A":
            continue

        sku    = str(row.iloc[sku_col]).strip()
        subfam = str(row.iloc[subfam_col]).strip() if subfam_col is not None else ""
        asin   = str(row.iloc[asin_col]).strip()  if asin_col  is not None else ""

        for country in COUNTRIES_ORDER:
            draft_price = parse_price(row.iloc[DRAFT_COL_IDX[country]])

            if country == "ES":
                tarifa_price = nac_map.get(sku)
                source       = "Nacional"
                currency     = "€"

                estado, msg, diff = _evaluar(draft_price, tarifa_price)
                records.append({
                    "SKU": sku, "Subfamilia": subfam, "ASIN": asin,
                    "País": country, "Fuente": source, "Moneda": currency,
                    "Precio draft": draft_price, "PVP PUB tarifa": tarifa_price,
                    "Diferencia": diff, "Estado": estado, "Detalle": msg,
                })
            else:
                tabs_for_country = [t for t, c in TAB_COUNTRY.items() if c == country]
                found_any = False

                for tab in tabs_for_country:
                    td = inter_maps.get(tab.upper())
                    if td is None:
                        continue
                    tarifa_price = td["map"].get(sku)
                    if tarifa_price is None:
                        continue
                    currency     = CURRENCY.get(country, "€")
                    source       = f"{tab}{' (local)' if td['is_local'] else ' (€)'}"
                    estado, msg, diff = _evaluar(draft_price, tarifa_price)
                    records.append({
                        "SKU": sku, "Subfamilia": subfam, "ASIN": asin,
                        "País": country, "Fuente": source, "Moneda": currency,
                        "Precio draft": draft_price, "PVP PUB tarifa": tarifa_price,
                        "Diferencia": diff, "Estado": estado, "Detalle": msg,
                    })
                    found_any = True

                if not found_any:
                    tabs_label = "/".join(tabs_for_country)
                    currency   = CURRENCY.get(country, "€")
                    records.append({
                        "SKU": sku, "Subfamilia": subfam, "ASIN": asin,
                        "País": country, "Fuente": tabs_label, "Moneda": currency,
                        "Precio draft": draft_price, "PVP PUB tarifa": None,
                        "Diferencia": None, "Estado": "⚠️ Sin tarifa", "Detalle": "SKU no encontrado en tarifa",
                    })

    return pd.DataFrame(records)


def _evaluar(draft_price, tarifa_price):
    if draft_price is None:
        return "⚠️ Sin precio", "Sin precio en draft", None
    if tarifa_price is None:
        return "⚠️ Sin tarifa", "SKU no en tarifa", None
    diff = round(draft_price - tarifa_price, 4)
    if diff < -0.01:
        return "❌ Inferior", f"Por debajo {diff:.2f}", diff
    return "✅ OK", "Igual o superior", diff


# ─────────────────────────────────────────────
# EXPORT
# ─────────────────────────────────────────────

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        # Hoja de errores
        errores = df[df["Estado"] != "✅ OK"].copy()
        errores.to_excel(writer, sheet_name="Errores", index=False)
        # Hoja completa
        df.to_excel(writer, sheet_name="Completo", index=False)
    return buf.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig").encode("utf-8-sig")


# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────

st.title("🔍 Verificador de precios · Prime Day Cluster A")
st.caption("Comprueba que los SKUs con Cluster A en el draft tienen precio ≥ PVP PUB. en tarifas nacional e internacional.")

# ── Sidebar: carga de ficheros ──────────────
with st.sidebar:
    st.header("📂 Ficheros de entrada")

    f_draft = st.file_uploader(
        "1 · Draft V3_PrimeDay_Expansión",
        type=["xlsx", "xls"],
        help="SKU col D · Cluster col P · Precios cols H–O",
    )
    f_nac = st.file_uploader(
        "2 · Tarifa nacional",
        type=["xlsx", "xls"],
        help="REFERENCIA col A · PVP PUB. col P",
    )
    f_inter = st.file_uploader(
        "3 · Tarifa internacional (TURACO)",
        type=["xlsx", "xls"],
        help="Pestañas: ES-FR FR-FR ES-IT IT-IT ES-DE DE-DE PT NL BE PL SE",
    )

    run_btn = st.button("▶️ Verificar", type="primary", use_container_width=True,
                        disabled=not (f_draft and f_nac and f_inter))

    st.divider()
    st.markdown("**Lógica aplicada**")
    st.markdown("""
- Solo filas con **Cluster = A**
- España → tarifa nacional (col A → col P)
- FR/IT/DE/NL/BE/PT → tarifas internacionales col C → **col M** (€)
- PL / SE → col C → **col N** (moneda local)
- El precio del draft debe ser **≥** PVP PUB. de tarifa
    """)

# ── Estado de carga ──────────────────────────
c1, c2, c3 = st.columns(3)
c1.success("✓ Draft cargado"   ) if f_draft else c1.info("⬆ Pendiente draft")
c2.success("✓ Nacional cargada") if f_nac   else c2.info("⬆ Pendiente nacional")
c3.success("✓ Inter cargada"  ) if f_inter  else c3.info("⬆ Pendiente internacional")

# ── Ejecución ────────────────────────────────
if run_btn:
    with st.spinner("Cargando y procesando ficheros..."):
        draft_df  = load_draft(f_draft.read())
        nac_map   = load_nacional(f_nac.read())
        inter_maps= load_internacional(f_inter.read())

    st.success(f"Tarifas internacionales: {len(inter_maps)} pestañas reconocidas → {', '.join(sorted(inter_maps.keys()))}")

    with st.spinner("Verificando precios..."):
        result_df = run_check(draft_df, nac_map, inter_maps)

    if result_df.empty:
        st.warning("No se encontraron filas con Cluster A o hubo un error de lectura.")
        st.stop()

    # ── Métricas ─────────────────────────────
    total    = len(result_df)
    errors   = (result_df["Estado"] == "❌ Inferior").sum()
    warnings = result_df["Estado"].str.startswith("⚠️").sum()
    ok       = (result_df["Estado"] == "✅ OK").sum()
    skus_a   = result_df["SKU"].nunique()

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("SKUs Cluster A",    skus_a)
    m2.metric("Comprobaciones",    total)
    m3.metric("✅ OK",              ok)
    m4.metric("❌ Precio inferior", errors,   delta=f"-{errors}" if errors else None, delta_color="inverse")
    m5.metric("⚠️ Sin dato",       warnings, delta=f"-{warnings}" if warnings else None, delta_color="inverse")

    st.divider()

    # ── Filtros ──────────────────────────────
    col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 3])
    with col_f1:
        pais_opts = ["Todos"] + COUNTRIES_ORDER
        fil_pais  = st.selectbox("País", pais_opts)
    with col_f2:
        estado_opts = ["Todos", "❌ Inferior", "⚠️ Sin tarifa", "⚠️ Sin precio", "✅ OK"]
        fil_estado  = st.selectbox("Estado", estado_opts)
    with col_f3:
        fil_sku = st.text_input("Buscar SKU")
    with col_f4:
        show_only_errors = st.checkbox("Solo errores y avisos", value=True)

    filtered = result_df.copy()
    if fil_pais != "Todos":
        filtered = filtered[filtered["País"] == fil_pais]
    if fil_estado != "Todos":
        filtered = filtered[filtered["Estado"] == fil_estado]
    if fil_sku:
        filtered = filtered[filtered["SKU"].str.contains(fil_sku, case=False, na=False)]
    if show_only_errors:
        filtered = filtered[filtered["Estado"] != "✅ OK"]

    st.caption(f"Mostrando {len(filtered)} de {total} comprobaciones")

    # ── Tabla ────────────────────────────────
    def color_estado(val):
        if val == "❌ Inferior":   return "background-color:#FCEBEB; color:#A32D2D; font-weight:500"
        if val.startswith("⚠️"):  return "background-color:#FAEEDA; color:#854F0B"
        if val == "✅ OK":         return "background-color:#EAF3DE; color:#3B6D11"
        return ""

    def color_diff(val):
        try:
            v = float(val)
            if v < -0.01: return "color:#A32D2D; font-weight:500"
            return "color:#3B6D11"
        except:
            return ""

    display_cols = ["SKU", "Subfamilia", "País", "Fuente", "Moneda",
                    "Precio draft", "PVP PUB tarifa", "Diferencia", "Estado", "Detalle"]
    display_df = filtered[display_cols].copy()

    styled = (
        display_df.style
        .applymap(color_estado, subset=["Estado"])
        .applymap(color_diff,   subset=["Diferencia"])
        .format({
            "Precio draft":    lambda x: f"{x:.2f}" if pd.notna(x) else "—",
            "PVP PUB tarifa":  lambda x: f"{x:.2f}" if pd.notna(x) else "—",
            "Diferencia":      lambda x: f"{x:+.2f}" if pd.notna(x) else "—",
        })
    )

    st.dataframe(styled, use_container_width=True, height=480, hide_index=True)

    # ── Descargas ─────────────────────────────
    st.divider()
    d1, d2, d3 = st.columns(3)

    only_errors = result_df[result_df["Estado"] != "✅ OK"]

    with d1:
        st.download_button(
            "⬇️ Exportar errores (.xlsx)",
            data=to_excel_bytes(only_errors),
            file_name="errores_primeday_clusterA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "⬇️ Exportar errores (.csv)",
            data=to_csv_bytes(only_errors),
            file_name="errores_primeday_clusterA.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with d3:
        st.download_button(
            "⬇️ Resultado completo (.xlsx)",
            data=to_excel_bytes(result_df),
            file_name="resultado_completo_clusterA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ── Resumen por país ──────────────────────
    st.divider()
    st.subheader("Resumen por país")
    summary = (
        result_df.groupby("País")["Estado"]
        .value_counts()
        .unstack(fill_value=0)
        .reindex(COUNTRIES_ORDER)
    )
    st.dataframe(summary, use_container_width=True)
