
import streamlit as st

import re
import time
import requests
from datetime import datetime
import tempfile
from pathlib import Path
import os
from copy import copy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# =========================
# CONFIG STREAMLIT
# =========================
st.set_page_config(page_title="DevengosCuentas2026", layout="wide")
st.title("DevengosCuentas2026 — Programa 1 + Programa 2")

# --- API KEY SOLO POR SECRETS (NO input en página)
try:
    API_KEY = (st.secrets.get("API_KEY", "") or "").strip()
except Exception:
    API_KEY = ""

if not API_KEY:
    st.error("Falta API_KEY en Secrets (Streamlit Cloud → Manage app → Settings → Secrets).")
    st.stop()

# =========================
# (PROGRAMA 1) — TAL CUAL
# =========================

# COLUMNAS API (PREPARADAS) - NUEVO ESQUEMA
MAX_ITEMS = 20

API_COLUMNAS = [
    "Proveedor_Nombre",
    "Proveedor_Rut",
    "Codigo_OC",
    "N_Licitacion",
    "Descripcion_OC",
    "FechaCreacion_OC",
    "TotalNeto_OC",
    "NombreContacto_OC",
    "CantidadItems_OC",
]

for i in range(1, MAX_ITEMS + 1):
    API_COLUMNAS.extend([
        f"Cantidad_{i}",
        f"EspecificacionComprador_{i}",
        f"PrecioNeto_{i}",
    ])

COLUMNAS_ORDEN = ["Fecha", "Folio", "Título", "Monto"] + API_COLUMNAS


def copiar_estilo(origen, destino):
    destino.font = copy(origen.font)
    destino.border = copy(origen.border)
    destino.fill = copy(origen.fill)
    destino.number_format = origen.number_format
    destino.alignment = copy(origen.alignment)


def limpiar_folio(serie):
    return serie.astype(float).astype(int).astype(str)


def normalizar_monto(serie):
    return (
        serie.astype(str)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
        .str.strip()
        .astype(float)
    )


def leer_sigfe_nuevo(ruta_excel):
    df = pd.read_excel(ruta_excel, skiprows=5)
    df = df[df["Tipo Vista"] == "Tipo Flujo"]

    df = df.rename(columns={
        df.columns[1]: "Monto_Total_Cuenta",
        df.columns[-1]: "Monto"
    })

    df["Cuenta"] = df["Concepto"].astype(str).str.extract(r"(\d+)")
    df["Nombre_Cuenta"] = (
        df["Concepto"]
        .astype(str)
        .str.replace(r"\d+", "", regex=True)
        .str.replace("-", "", regex=False)
        .str.strip()
    )

    df["Fecha"] = pd.to_datetime(df["Fecha"]).dt.date

    df = df[[
        "Cuenta",
        "Nombre_Cuenta",
        "Fecha",
        "Folio",
        "Título",
        "Monto"
    ]].copy()

    df["Monto"] = normalizar_monto(df["Monto"])
    df["Folio"] = limpiar_folio(df["Folio"])

    for col in API_COLUMNAS:
        df[col] = ""

    return df


def leer_maestro(ruta_maestro):
    base_cols = ["Cuenta", "Fecha", "Folio", "Título", "Monto"] + API_COLUMNAS

    if not os.path.exists(ruta_maestro):
        return pd.DataFrame(columns=base_cols)

    xls = pd.ExcelFile(ruta_maestro)
    dfs = []

    for sheet in xls.sheet_names:
        temp = pd.read_excel(xls, sheet_name=sheet, skiprows=0)

        columnas_min = {"Fecha", "Folio", "Monto"}
        if not columnas_min.issubset(temp.columns):
            continue

        cols = list(temp.columns)
        if len(cols) >= 3:
            cols[2] = "Título"
            temp.columns = cols

        temp = temp.copy()
        temp["Cuenta"] = str(sheet)

        temp = temp[temp["Folio"].notna()]
        temp["Folio"] = limpiar_folio(temp["Folio"])

        temp["Fecha"] = pd.to_datetime(temp["Fecha"], errors="coerce").dt.date
        temp = temp[temp["Fecha"].notna()]

        for col in API_COLUMNAS:
            if col not in temp.columns:
                temp[col] = ""

        for col in ["Título"]:
            if col not in temp.columns:
                temp[col] = ""

        dfs.append(temp)

    if not dfs:
        return pd.DataFrame(columns=base_cols)

    return pd.concat(dfs, ignore_index=True)


def obtener_devengos_nuevos(df_nuevo, df_maestro):
    if df_maestro.empty:
        return df_nuevo.copy()

    df_maestro = df_maestro.copy()
    df_nuevo = df_nuevo.copy()

    df_maestro["Cuenta"] = df_maestro["Cuenta"].astype(str)
    df_nuevo["Cuenta"] = df_nuevo["Cuenta"].astype(str)

    df_maestro["Folio"] = limpiar_folio(df_maestro["Folio"])
    df_nuevo["Folio"] = limpiar_folio(df_nuevo["Folio"])

    claves_maestro = set(
        df_maestro[["Cuenta", "Folio"]]
        .drop_duplicates()
        .itertuples(index=False, name=None)
    )

    claves_nuevo = df_nuevo[["Cuenta", "Folio"]].itertuples(index=False, name=None)
    mask = [clave not in claves_maestro for clave in claves_nuevo]

    return df_nuevo.loc[mask].copy()


def aplicar_formato_corporativo(wb, df_total):
    header_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=9)
    header_fill = PatternFill("solid", fgColor="1F4E78")

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for ws in wb.worksheets:
        cuenta = ws.title

        nombre_cuenta = ""
        if "Nombre_Cuenta" in df_total.columns:
            aux = df_total[
                (df_total["Cuenta"] == cuenta) &
                (df_total["Nombre_Cuenta"].notna())
            ]["Nombre_Cuenta"]
            if not aux.empty:
                nombre_cuenta = aux.iloc[0]

        header_titulo = f"DEVENGOS 2026 - CUENTA {cuenta} - {nombre_cuenta}"

        for j, colname in enumerate(COLUMNAS_ORDEN, start=1):
            if j == 3:
                valor_header = header_titulo
            else:
                valor_header = colname

            c = ws.cell(row=1, column=j)
            c.value = valor_header
            c.font = header_font
            c.fill = header_fill
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")

        ws.freeze_panes = "A2"

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(COLUMNAS_ORDEN)):
            for cell in row:
                cell.font = body_font
                cell.border = thin_border

        for cell in ws["A"][1:]:
            cell.number_format = "DD-MM-YYYY"
        for cell in ws["D"][1:]:
            cell.number_format = "#,##0"

        for col_cells in ws.iter_cols(min_row=1, max_row=ws.max_row, max_col=len(COLUMNAS_ORDEN)):
            max_length = 0
            col_letter = get_column_letter(col_cells[0].column)
            for cell in col_cells:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2


def ejecutar_programa_1(ruta_nuevo, ruta_maestro):
    df_nuevo = leer_sigfe_nuevo(ruta_nuevo)
    df_maestro = leer_maestro(ruta_maestro)

    archivo_nuevo = not os.path.exists(ruta_maestro)
    df_nuevos = obtener_devengos_nuevos(df_nuevo, df_maestro)

    if not os.path.exists(ruta_maestro):
        with pd.ExcelWriter(ruta_maestro, engine="openpyxl") as writer:
            for cuenta in df_nuevos["Cuenta"].unique():
                datos = df_nuevos[df_nuevos["Cuenta"] == cuenta].drop(columns=["Cuenta", "Nombre_Cuenta"])
                datos = datos.reindex(columns=COLUMNAS_ORDEN)
                datos.to_excel(writer, sheet_name=str(cuenta), index=False)

    wb = load_workbook(ruta_maestro)

    if not archivo_nuevo:
        for cuenta in df_nuevos["Cuenta"].unique():
            ws = wb[cuenta] if cuenta in wb.sheetnames else wb.create_sheet(cuenta)
            nuevos = df_nuevos[df_nuevos["Cuenta"] == cuenta]

            if ws.max_row < 1:
                for j, colname in enumerate(COLUMNAS_ORDEN, start=1):
                    ws.cell(row=1, column=j).value = colname

            fila_modelo = 2 if ws.max_row >= 2 else 1

            for _, row in nuevos.iterrows():
                fila_nueva = ws.max_row + 1
                valores = [row.get(col, "") for col in COLUMNAS_ORDEN]

                for col_idx, valor in enumerate(valores, start=1):
                    celda_nueva = ws.cell(row=fila_nueva, column=col_idx, value=valor)
                    celda_modelo = ws.cell(row=fila_modelo, column=col_idx)
                    copiar_estilo(celda_modelo, celda_nueva)

    df_total = pd.concat([df_maestro, df_nuevos], ignore_index=True)
    aplicar_formato_corporativo(wb, df_total)
    wb.save(ruta_maestro)


# =========================
# (PROGRAMA 2) — TAL CUAL
# =========================
TIMEOUT = 25
MAX_REINTENTOS = 4
PAUSA_ENTRE_LLAMADAS = 0.25  # seg


def normalizar_guiones(s: str) -> str:
    return (s.replace("\u2010", "-")
             .replace("\u2011", "-")
             .replace("\u2012", "-")
             .replace("\u2013", "-")
             .replace("\u2212", "-"))


def extraer_codigo_oc(texto):
    if texto is None:
        return None
    s = normalizar_guiones(str(texto))
    m = re.search(r"(\d+)-(\d+)-([A-Za-z0-9]+)", s)
    if not m:
        return None
    return f"{m.group(1)}-{m.group(2)}-{m.group(3).upper()}"


def mapear_headers_fila1(ws):
    headers = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v is None:
            continue
        k = str(v).strip()
        if k:
            headers[k] = col
    return headers


def buscar_columna_texto(ws, texto_objetivo, fila=1):
    texto_objetivo = str(texto_objetivo).strip().lower()
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=fila, column=col).value
        if v is None:
            continue
        if texto_objetivo in str(v).strip().lower():
            return col
    return None

_ILLEGAL_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

def limpiar_excel(valor):
    
    if valor is None:
        return None
    if isinstance(valor, str):
        v = valor.replace("\r", " ").replace("\n", " ")
        return _ILLEGAL_RE.sub("", v)
    return valor

def obtener_datos_oc(session, codigo_oc):
    url = "https://api.mercadopublico.cl/servicios/v1/publico/ordenesdecompra.json"
    params = {"codigo": codigo_oc, "ticket": API_KEY}

    last_err = None

    for intento in range(1, MAX_REINTENTOS + 1):
        try:
            r = session.get(url, params=params, timeout=TIMEOUT)
            status = r.status_code

            try:
                resp = r.json()
            except Exception:
                last_err = f"respuesta no-JSON (status {status})"
                resp = None

            if status != 200:
                last_err = f"status {status}"
            elif not resp:
                pass
            else:
                if resp.get("Listado"):
                    o = resp["Listado"][0]
                    datos = {}

                    datos["Codigo_OC"] = o.get("Codigo")
                    datos["N_Licitacion"] = o.get("CodigoLicitacion")

                    datos["Descripcion_OC"] = (o.get("Descripcion", "") or "") \
                        .replace("\n", " ").replace("\r", " ")

                    fecha = o.get("Fechas", {}).get("FechaCreacion")
                    datos["FechaCreacion_OC"] = (
                        datetime.fromisoformat(fecha.split(".")[0]).strftime("%d/%m/%Y")
                        if fecha else ""
                    )

                    datos["TotalNeto_OC"] = o.get("TotalNeto")
                    datos["NombreContacto_OC"] = (o.get("Comprador", {}) or {}).get("NombreContacto")
                    datos["CantidadItems_OC"] = (o.get("Items", {}) or {}).get("Cantidad") or 0

                    prov = o.get("Proveedor", {}) or {}
                    datos["Proveedor_Nombre"] = prov.get("Nombre", "")
                    datos["Proveedor_Rut"] = prov.get("RutSucursal") or prov.get("Rut") or ""

                    items = (o.get("Items", {}) or {}).get("Listado", []) or []

                    for i, item in enumerate(items[:MAX_ITEMS], start=1):
                        datos[f"Cantidad_{i}"] = item.get("Cantidad")
                        datos[f"EspecificacionComprador_{i}"] = (item.get("EspecificacionComprador", "") or "") \
                            .replace("\n", " ").replace("\r", " ")
                        datos[f"PrecioNeto_{i}"] = item.get("PrecioNeto")

                    return datos, None

                msg = resp.get("Mensaje") or resp.get("message") or resp.get("Error") or ""
                last_err = f"sin Listado (status {status}) {msg}".strip()

        except requests.RequestException as e:
            last_err = f"request error: {type(e).__name__}"

        time.sleep(0.8 * intento)

    return None, last_err or "sin datos"


def ejecutar_programa_2(ruta_maestro):
    wb = load_workbook(ruta_maestro)

    # ✅ PROBAR SOLO HOJA 1 y 2
    HOJAS_OMITIDAS = {"220400400101","220400400102"}

    cache_ok = {}
    cache_fail = {}

    session = requests.Session()

    for ws in wb.worksheets:
        if ws.title in HOJAS_OMITIDAS:
            continue

        headers = mapear_headers_fila1(ws)

        col_titulo = buscar_columna_texto(ws, "DEVENGOS 2026 - CUENTA", fila=1)
        if not col_titulo:
            continue

        if "Codigo_OC" not in headers:
            continue

        col_codigo = headers["Codigo_OC"]
        fila = 2

        while fila <= ws.max_row:
            titulo = ws.cell(row=fila, column=col_titulo).value

            codigo_oc = extraer_codigo_oc(titulo)
            if not codigo_oc:
                fila += 1
                continue

            if ws.cell(row=fila, column=col_codigo).value:
                fila += 1
                continue

            if codigo_oc in cache_ok:
                datos = cache_ok[codigo_oc]
            else:
                time.sleep(PAUSA_ENTRE_LLAMADAS)
                datos, err = obtener_datos_oc(session, codigo_oc)

                if datos:
                    cache_ok[codigo_oc] = datos
                else:
                    cache_fail[codigo_oc] = err
                    fila += 1
                    continue

            for campo, valor in datos.items():
                if campo in headers:
                    ws.cell(row=fila, column=headers[campo]).value = limpiar_excel(valor)

            fila += 1

    wb.save(ruta_maestro)


# =========================
# UI + DESCARGA (NO se rompe)
# =========================
sigfe_file = st.file_uploader("1) Subir SA_MayorPresupuestario.xls", type=["xls", "xlsx"])
maestro_file = st.file_uploader("2) (Opcional) Subir DevengosCuentas2026.xlsx existente", type=["xlsx"])

if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

c1, c2 = st.columns(2)
with c1:
    run_p1 = st.checkbox("Ejecutar Programa 1", value=True)
with c2:
    run_p2 = st.checkbox("Ejecutar Programa 2 (API)", value=True)

if st.button("Procesar"):
    if sigfe_file is None:
        st.error("Falta SA_MayorPresupuestario.xls")
        st.stop()

    with st.spinner("Procesando..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            ruta_sigfe = tmp / "SA_MayorPresupuestario.xls"
            ruta_sigfe.write_bytes(sigfe_file.getbuffer())

            ruta_maestro = tmp / "DevengosCuentas2026.xlsx"
            if maestro_file is not None:
                ruta_maestro.write_bytes(maestro_file.getbuffer())

            if run_p1:
                ejecutar_programa_1(str(ruta_sigfe), str(ruta_maestro))

            if run_p2:
                ejecutar_programa_2(str(ruta_maestro))

            st.session_state.excel_bytes = ruta_maestro.read_bytes()
            st.success(f"✅ Listo. Archivo generado: {len(st.session_state.excel_bytes):,} bytes")

if st.session_state.excel_bytes:
    st.download_button(
        "⬇️ Descargar DevengosCuentas2026.xlsx",
        data=st.session_state.excel_bytes,
        file_name="DevengosCuentas2026.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

