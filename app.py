"""
app.py — Interfaz principal ONCH · Automatizador SEDAPAL
Ejecutar con:  streamlit run app.py
"""
import sys
from pathlib import Path

# Asegurar que el directorio raíz esté en el path de Python
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import streamlit as st
from datetime import datetime

from config import OUTPUT_DIR, MESES
from modules.ingesta import (
    listar_archivos, guardar_archivo, eliminar_archivo,
    get_plantilla_path, get_estilos_path,
)
from modules.generador import generar_plantilla
from modules.conciliacion import leer_totales_output, comparar_con_esperado

# ─── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="ONCH · Automatizador SEDAPAL",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS mínimo ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .stTabs [data-baseweb="tab"] { font-size: 1rem; font-weight: 600; }
    .metric-label { font-size: 0.85rem !important; }
</style>
""", unsafe_allow_html=True)

# ─── SIDEBAR — Parámetros globales ───────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🔧 ONCH · SEDAPAL")
    st.markdown("---")
    st.subheader("⚙️ Parámetros")

    zona = st.selectbox("Zona / Región", ["NORTE", "CENTRO"], index=0, key="zona")
    tipo = st.selectbox("Tipo de Entidad", ["CR", "POZO"], index=0, key="tipo")

    st.markdown("---")
    st.subheader("📅 Período")

    mes_idx = st.selectbox(
        "Mes", options=list(range(12)),
        format_func=lambda i: MESES[i],
        index=datetime.now().month - 1,
    )
    anio = st.number_input("Año", min_value=2024, max_value=2030,
                           value=datetime.now().year, step=1)

    mes_label   = f"{MESES[mes_idx].upper()}{anio}"
    output_fname = f"{zona}_{tipo}_{mes_label}.xlsm"
    output_path  = OUTPUT_DIR / output_fname
    plantilla_path = get_plantilla_path(zona, tipo)
    estilos_path   = get_estilos_path(zona, tipo)

    st.markdown("---")
    st.info(
        f"**Período:** {MESES[mes_idx]} {anio}\n\n"
        f"**Zona:** {zona}  |  **Tipo:** {tipo}\n\n"
        f"**Salida:** `{output_fname}`"
    )

# ─── TABS ─────────────────────────────────────────────────────────────────────
st.title("ONCH · Automatizador SEDAPAL")
tab1, tab2, tab3 = st.tabs([
    "Gestión de Archivos",
    "Generación",
    "Conciliación",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Gestión de Archivos
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header(f"📁 Archivos — {zona} / {tipo}")

    # ── Archivos fuente ────────────────────────────────────────────────────────
    st.subheader("📤 Archivos Fuente (.xlsx)")
    st.caption(
        "Sube los archivos de los CR trabajados en el período. "
        "Se guardan de forma permanente y estarán disponibles en futuras sesiones."
    )

    col_up, col_list = st.columns([1, 2])

    with col_up:
        uploaded = st.file_uploader(
            "Arrastra o selecciona archivos .xlsx",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"fuentes_{zona}_{tipo}",
        )
        if uploaded:
            for f in uploaded:
                guardar_archivo(f, zona, tipo, "fuentes")
            st.success(f"✅ {len(uploaded)} archivo(s) guardado(s)")
            st.rerun()

    with col_list:
        archivos = listar_archivos(zona, tipo, "fuentes")
        if archivos:
            df_files = pd.DataFrame([
                {
                    "Archivo": a["nombre"],
                    "CR":      a["cr"],
                    "Fecha":   a["fecha_archivo"],
                    "Tamaño":  a["tamano"],
                }
                for a in archivos
            ])
            st.dataframe(df_files, use_container_width=True, hide_index=True)
            st.caption(f"Total: **{len(archivos)}** archivo(s)")

            # Eliminar archivos
            with st.expander("🗑️ Eliminar archivos"):
                to_delete = st.multiselect(
                    "Selecciona archivos a eliminar",
                    options=[a["nombre"] for a in archivos],
                )
                if st.button("Eliminar seleccionados", disabled=not to_delete):
                    for nombre in to_delete:
                        eliminar_archivo(zona, tipo, "fuentes", nombre)
                    st.success(f"Eliminados: {len(to_delete)}")
                    st.rerun()
        else:
            st.info(f"No hay archivos fuente para {zona}/{tipo}. Sube tus .xlsx arriba.")

    st.markdown("---")

    # ── Plantillas ─────────────────────────────────────────────────────────────
    st.subheader("📋 Plantillas de Referencia")
    col_p1, col_p2 = st.columns(2)

    with col_p1:
        status_plantilla = f"✅ {plantilla_path.name}" if plantilla_path.exists() else "❌ No cargada"
        st.markdown(f"**Plantilla vacía**\n\n{status_plantilla}")
        p_up = st.file_uploader(
            "Sube la plantilla vacía",
            type=["xlsm"],
            key=f"plantilla_{zona}_{tipo}",
        )
        if p_up:
            guardar_archivo(p_up, zona, tipo, "plantillas",
                            nombre_fijo=f"plantilla_{zona.lower()}_vacia.xlsm")
            st.success("✅ Plantilla guardada")
            st.rerun()

    with col_p2:
        status_estilos = f"✅ {estilos_path.name}" if estilos_path.exists() else "❌ No cargado"
        st.markdown(f"**Referencia/Estilos**\n\n{status_estilos}")
        e_up = st.file_uploader(
            "Sube el archivo de referencia",
            type=["xlsm"],
            key=f"estilos_{zona}_{tipo}",
        )
        if e_up:
            guardar_archivo(e_up, zona, tipo, "plantillas",
                            nombre_fijo=f"estilos_{zona.lower()}_referencia.xlsm")
            st.success("✅ Estilos guardados")
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Generación
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header(f"⚙️ Generación — {zona} / {tipo} / {MESES[mes_idx]} {anio}")

    archivos_fuente = listar_archivos(zona, tipo, "fuentes")
    can_run = (
        len(archivos_fuente) > 0
        and plantilla_path.exists()
        and estilos_path.exists()
    )

    # Estado
    col_st, col_acc = st.columns([3, 1])
    with col_st:
        st.subheader("Estado actual")
        def check(cond, label):
            return f"{'✅' if cond else '❌'} {label}"

        st.write(check(len(archivos_fuente) > 0,
                       f"Archivos fuente: **{len(archivos_fuente)}** archivo(s)"))
        st.write(check(plantilla_path.exists(),
                       f"Plantilla vacía: `{plantilla_path.name}`"))
        st.write(check(estilos_path.exists(),
                       f"Archivo de estilos: `{estilos_path.name}`"))

        if output_path.exists():
            size_mb = output_path.stat().st_size / 1024 / 1024
            st.write(f"📄 Archivo previo: `{output_fname}` ({size_mb:.2f} MB) — se sobreescribirá")

    with col_acc:
        st.subheader("Acción")
        if not can_run:
            st.warning("Completa los requisitos en 'Gestión de Archivos'.")

        run_btn = st.button(
            "▶ Generar",
            disabled=not can_run,
            type="primary",
            use_container_width=True,
        )

    # Ejecución
    if run_btn:
        source_paths = [a["path"] for a in archivos_fuente]

        log_area    = st.empty()
        progress_bar = st.progress(0.0)
        all_logs: list[str] = []

        def on_log(msg: str, pct: float):
            all_logs.append(msg)
            log_area.code("\n".join(all_logs[-30:]))
            progress_bar.progress(min(pct, 1.0))

        try:
            logs, total_rows = generar_plantilla(
                zona=zona,
                source_paths=source_paths,
                plantilla_path=str(plantilla_path),
                estilos_path=str(estilos_path),
                output_path=str(output_path),
                log_callback=on_log,
            )
            st.success(
                f"✅ **{output_fname}** generado correctamente — {total_rows} filas"
            )
        except Exception as exc:
            st.error(f"❌ Error durante la generación: {exc}")
            st.exception(exc)

    # Descarga si el archivo ya existe
    if output_path.exists():
        st.markdown("---")
        with open(output_path, "rb") as fout:
            data_bytes = fout.read()
        st.download_button(
            label=f"⬇️ Descargar {output_fname}",
            data=data_bytes,
            file_name=output_fname,
            mime="application/vnd.ms-excel.sheet.macroenabled.12",
            use_container_width=False,
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Conciliación
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header(f"📊 Conciliación — {zona} / {tipo} / {MESES[mes_idx]} {anio}")

    if not output_path.exists():
        st.warning(
            f"No hay archivo generado para **{zona}/{tipo}/{MESES[mes_idx]} {anio}**. "
            "Ve a la pestaña 'Generación' primero."
        )
        st.stop()

    # Leer totales reales del archivo generado
    try:
        totales_reales = leer_totales_output(str(output_path))
    except Exception as exc:
        st.error(f"No se pudo leer el archivo generado: {exc}")
        st.stop()

    crs_detectados = sorted(totales_reales.keys())

    if not crs_detectados:
        st.warning("El archivo generado no contiene filas PM03 legibles.")
        st.stop()

    st.caption(f"CRs detectados en el archivo: **{', '.join(crs_detectados)}**")
    st.markdown("---")

    # ── Ingresar montos esperados ──────────────────────────────────────────────
    st.subheader("📥 Montos Esperados / Teóricos")
    st.caption(
        "Ingresa el monto total esperado por CR. "
        "Puede ser el valor del contrato, la pre-valorización o el presupuesto aprobado."
    )

    session_key = f"esperados_{zona}_{tipo}_{mes_label}"
    if session_key not in st.session_state:
        st.session_state[session_key] = pd.DataFrame({
            "CR":                  crs_detectados,
            "Monto Esperado (S/)": [0.0] * len(crs_detectados),
        })
    else:
        # Mantener sincronizado si cambiaron los CRs
        existing = st.session_state[session_key]
        existing_crs = set(existing["CR"].tolist())
        new_crs = set(crs_detectados) - existing_crs
        if new_crs:
            new_rows = pd.DataFrame({
                "CR": sorted(new_crs),
                "Monto Esperado (S/)": [0.0] * len(new_crs),
            })
            st.session_state[session_key] = pd.concat(
                [existing, new_rows], ignore_index=True
            ).sort_values("CR").reset_index(drop=True)

    col_edit, col_import = st.columns([2, 1])

    with col_edit:
        edited_df = st.data_editor(
            st.session_state[session_key],
            use_container_width=True,
            hide_index=True,
            column_config={
                "CR": st.column_config.TextColumn("CR", disabled=True, width="small"),
                "Monto Esperado (S/)": st.column_config.NumberColumn(
                    "Monto Esperado (S/)",
                    format="S/ %.2f",
                    step=0.01,
                    min_value=0.0,
                    width="medium",
                ),
            },
            key=f"editor_{session_key}",
        )
        st.session_state[session_key] = edited_df

    with col_import:
        st.markdown("**O importa desde CSV/Excel**")
        esp_file = st.file_uploader(
            "Cargar montos esperados",
            type=["csv", "xlsx"],
            key=f"esp_upload_{session_key}",
        )
        if esp_file:
            try:
                if esp_file.name.endswith(".csv"):
                    df_imp = pd.read_csv(esp_file)
                else:
                    df_imp = pd.read_excel(esp_file)
                # Buscar columnas CR y Monto
                col_cr  = next((c for c in df_imp.columns if "cr" in c.lower()), None)
                col_mon = next((c for c in df_imp.columns if "monto" in c.lower() or
                                "esperado" in c.lower() or "teorico" in c.lower()), None)
                if col_cr and col_mon:
                    imp_dict = dict(zip(df_imp[col_cr].astype(str), df_imp[col_mon].astype(float)))
                    df_base = st.session_state[session_key].copy()
                    df_base["Monto Esperado (S/)"] = df_base["CR"].map(imp_dict).fillna(0.0)
                    st.session_state[session_key] = df_base
                    st.success("✅ Importado correctamente")
                    st.rerun()
                else:
                    st.error("El archivo debe tener columnas 'CR' y 'Monto Esperado'.")
            except Exception as exc:
                st.error(f"Error al importar: {exc}")

    st.markdown("---")

    # ── Calcular y mostrar conciliación ───────────────────────────────────────
    if st.button("📊 Calcular Conciliación", type="primary"):
        esperados_dict = dict(
            zip(
                st.session_state[session_key]["CR"],
                st.session_state[session_key]["Monto Esperado (S/)"],
            )
        )
        resultado = comparar_con_esperado(totales_reales, esperados_dict)
        df_res    = pd.DataFrame(resultado)

        st.subheader("📋 Resultado de Conciliación")

        # Métricas resumen
        total_real     = df_res["Real (S/)"].sum()
        total_esperado = df_res["Esperado (S/)"].sum()
        total_dif      = total_real - total_esperado

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Real (S/)",     f"{total_real:,.2f}")
        m2.metric("Total Esperado (S/)", f"{total_esperado:,.2f}")
        m3.metric(
            "Diferencia (S/)",
            f"{total_dif:+,.2f}",
            delta_color="inverse" if total_dif < 0 else "normal",
        )
        m4.metric("CRs conciliados", len(df_res))

        st.markdown("")

        # Tabla con colores
        def color_dif(val):
            if abs(val) < 0.02:
                return "background-color: #d4edda; color: #155724"
            elif val > 0:
                return "background-color: #fff3cd; color: #856404"
            else:
                return "background-color: #f8d7da; color: #721c24"

        styled = (
            df_res.style
            .map(color_dif, subset=["Diferencia"])
            .format({
                "E (S/)":        "S/ {:,.2f}",
                "TA (S/)":       "S/ {:,.2f}",
                "AH (S/)":       "S/ {:,.2f}",
                "Real (S/)":     "S/ {:,.2f}",
                "Esperado (S/)": "S/ {:,.2f}",
                "Diferencia":    "S/ {:+,.2f}",
                "Var %":         "{:+.1f}%",
            })
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)

        # Exportar
        csv_data = df_res.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="⬇️ Exportar Conciliación (.csv)",
            data=csv_data,
            file_name=f"conciliacion_{zona}_{tipo}_{mes_label}.csv",
            mime="text/csv",
        )
