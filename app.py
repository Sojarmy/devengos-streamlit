import streamlit as st
import tempfile
from pathlib import Path

st.set_page_config(page_title="Devengos HBV", layout="wide")

st.title("DevengosCuentas2026")
st.write("Sube los archivos SIGFE y descarga el Excel final.")

sigfe_file = st.file_uploader(
    "1️⃣ Subir SA_MayorPresupuestario.xls",
    type=["xls", "xlsx"]
)

maestro_file = st.file_uploader(
    "2️⃣ (Opcional) Subir DevengosCuentas2026.xlsx",
    type=["xlsx"]
)

procesar = st.button("Procesar")

def guardar(uploaded, ruta):
    with open(ruta, "wb") as f:
        f.write(uploaded.getbuffer())

if procesar:
    if not sigfe_file:
        st.error("Falta SA_MayorPresupuestario.xls")
        st.stop()

    with st.spinner("Procesando..."):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)

            ruta_sigfe = tmp / "SA_MayorPresupuestario.xls"
            guardar(sigfe_file, ruta_sigfe)

            if maestro_file:
                ruta_maestro = tmp / "DevengosCuentas2026.xlsx"
                guardar(maestro_file, ruta_maestro)
            else:
                ruta_maestro = tmp / "DevengosCuentas2026.xlsx"

            # 🔴 AQUÍ DESPUÉS PEGAREMOS TUS FUNCIONES
            # generar_maestro_desde_sigfe(ruta_sigfe, ruta_maestro)
            # completar_desde_api(ruta_maestro)

            if not ruta_maestro.exists():
                st.error("No se generó el archivo final")
                st.stop()

            st.success("Proceso completado")

            with open(ruta_maestro, "rb") as f:
                st.download_button(
                    "Descargar DevengosCuentas2026.xlsx",
                    f,
                    file_name="DevengosCuentas2026.xlsx"
                )
