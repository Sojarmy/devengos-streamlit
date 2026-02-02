
import streamlit as st
import tempfile
from pathlib import Path
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from copy import copy

# =========================
# CONFIG
# =========================
MAX_ITEMS = 20

# =========================
# UTILIDADES
# =========================
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

# =========================
# COLUMNAS (SIN Item_#)
# =========================
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

# =========================
# LECTURA SIGFE
# =========================
def leer_sigfe_nuevo(ruta_excel):
    df = None
    last_err = None
    for eng in ["xlrd", "openpyxl", None]:
        try:
            df = pd.read_excel(ruta_excel, skiprows=5, engine=eng)
            break
        except Exception as e:
            last_err = e
            df = None
    if df is None:
        raise RuntimeError(f"No pude leer SIGFE: {last_err}")

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

# =========================
# LECTURA MAESTRO
# =========================
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

        if "Título" not in temp.columns:
            temp["Título"] = ""

        dfs.append(temp)

    if not dfs:
        return pd.DataFrame(columns=base_cols)

    return pd.concat(dfs, ignore_index=True)

# =========================
# NUEVOS DEVENGOS
# =========================
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

# =========================
# FORMATO CORPORATIVO
# =========================
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
                nombre_cuenta = str(aux.iloc[0])

        header_titulo = f"DEVENGOS 2026 - CUENTA {cuenta} - {nombre_cuenta}"

        for j, colname in enumerate(COLUMNAS_ORDEN, start=1):
            valor = header_titulo if j == 3 else colname
            c = ws.cell(row=1, column=j)
            c.value = valor
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
            max_len = 0
            letter = get_column_letter(col_cells[0].column)
            for cell in col_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 60)

# =========================
# PROGRAMA 1 (FUNCION)
# =========================
def generar_devengos_desde_sigfe(ruta_sigfe, ruta_maestro):
    df_nuevo = leer_sigfe_nuevo(ruta_sigfe)
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

    if (not archivo_nuevo) and (not df_nuevos.empty):
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

    df_total = pd.concat([df_maestro, df_nuevos], ignore_index=True) if not df_maestro.empty else df_nuevo.copy()
    aplicar_formato_corporativo(wb, df_total)
    wb.save(ruta_maestro)

# =========================
# STREAMLIT UI (SOLO P1)
# =========================
st.set_page_config(page_title="DevengosCuentas2026", layout="wide")
st.title("DevengosCuentas2026 - Programa 1")

sigfe_file = st.file_uploader("Subir SA_MayorPresupuestario.xls", type=["xls", "xlsx"])
maestro_file = st.file_uploader("(Opcional) Subir DevengosCuentas2026.xlsx existente", type=["xlsx"])

if st.button("Procesar Programa 1"):
    if sigfe_file is None:
        st.error("Falta SA_MayorPresupuestario.xls")
        st.stop()

    with st.spinner("Procesando Programa 1..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)

            ruta_sigfe = tmp / "SA_MayorPresupuestario.xls"
            ruta_sigfe.write_bytes(sigfe_file.getbuffer())

            ruta_maestro = tmp / "DevengosCuentas2026.xlsx"
            if maestro_file is not None:
                ruta_maestro.write_bytes(maestro_file.getbuffer())

            generar_devengos_desde_sigfe(str(ruta_sigfe), str(ruta_maestro))

            st.success("✅ Programa 1 listo")
            st.download_button(
                "⬇️ Descargar DevengosCuentas2026.xlsx",
                data=ruta_maestro.read_bytes(),
                file_name="DevengosCuentas2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
