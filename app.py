"""
app.py — Interfaz principal ONCH · Automatizador SEDAPAL
Ejecutar con:  streamlit run app.py
"""
import re
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
from modules.conciliacion import leer_totales_output, comparar_con_esperado, normalizar_codigo
from modules import asistente

# ─── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="ONCH · Automatizador SEDAPAL",
    page_icon="🔧",
    layout="wide",
)

# ─── CSS mínimo ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    header[data-testid="stHeader"] { display: none !important; }
    footer { display: none !important; }
    #MainMenu { display: none !important; }
    .block-container { padding-top: 0.5rem !important; }
    .stTabs [data-baseweb="tab"] { font-size: 1rem; font-weight: 600; }
    .metric-label { font-size: 0.85rem !important; }
    /* Barra de parámetros */
    .param-bar {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 8px;
        padding: 0.5rem 1rem 0.2rem 1rem;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ─── Barra de parámetros (reemplaza la sidebar) ───────────────────────────────
st.markdown('<div class="param-bar">', unsafe_allow_html=True)

col_titulo, col_zona, col_tipo, col_mes, col_anio, col_info = st.columns(
    [1.4, 1, 0.9, 1.4, 0.75, 2.5]
)

with col_titulo:
    st.markdown("### 🔧 ONCH · SEDAPAL")

with col_zona:
    zona = st.selectbox("Zona", ["NORTE", "CENTRO"], index=0)

with col_tipo:
    tipo = st.selectbox("Tipo", ["CR", "POZO"], index=0)

with col_mes:
    mes_idx = st.selectbox(
        "Mes", options=list(range(12)),
        format_func=lambda i: MESES[i],
        index=datetime.now().month - 1,
    )

with col_anio:
    anio = st.number_input("Año", min_value=2024, max_value=2030,
                           value=datetime.now().year, step=1)

# Variables derivadas (deben calcularse después de los widgets)
mes_label    = f"{MESES[mes_idx].upper()}{anio}"
output_fname = f"{zona}_{tipo}_{mes_label}.xlsm"
output_path  = OUTPUT_DIR / output_fname
plantilla_path = get_plantilla_path(zona, tipo)
estilos_path   = get_estilos_path(zona, tipo)

with col_info:
    st.markdown(
        f"📅 **{MESES[mes_idx]} {anio}** &nbsp;·&nbsp; 📁 `{output_fname}`",
        help=f"Zona: {zona} | Tipo: {tipo} | Período: {MESES[mes_idx]} {anio}"
    )

st.markdown('</div>', unsafe_allow_html=True)

# ─── Session state — archivos fuente (persisten al cambiar de pestaña) ────────
fuentes_key = f"ss_fuentes_{zona}_{tipo}"
if fuentes_key not in st.session_state:
    st.session_state[fuentes_key] = {}   # { nombre_archivo: bytes }



# ─── TABS ─────────────────────────────────────────────────────────────────────
# Log de la última generación (lo usa la pestaña Asistente)
logs_key = f"ss_logs_{output_fname}"

tab1, tab2, tab3, tab4 = st.tabs([
    "📁  Gestión de Archivos",
    "⚙️  Generación",
    "📊  Conciliación",
    "💬  Asistente",
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
        "Se guardan en esta sesión y estarán disponibles en todas las pestañas."
    )

    # Versión del uploader — se incrementa al eliminar archivos para resetear el widget
    _uploader_v_key = f"_uploader_v_{zona}_{tipo}"
    if _uploader_v_key not in st.session_state:
        st.session_state[_uploader_v_key] = 0

    col_up, col_list = st.columns([1, 2])

    with col_up:
        uploaded = st.file_uploader(
            "Arrastra o selecciona archivos .xlsx",
            type=["xlsx"],
            accept_multiple_files=True,
            key=f"fuentes_{zona}_{tipo}_{st.session_state[_uploader_v_key]}",
        )
        if uploaded:
            for f in uploaded:
                content = f.read()
                if content:
                    st.session_state[fuentes_key][f.name] = content
            st.success(f"✅ {len(uploaded)} archivo(s) cargado(s)")

    with col_list:
        ss_files = st.session_state[fuentes_key]
        if ss_files:
            df_files = pd.DataFrame([
                {"Archivo": nombre, "Tamaño": f"{len(data)/1024:.1f} KB"}
                for nombre, data in ss_files.items()
            ])
            st.dataframe(df_files, use_container_width=True, hide_index=True)
            st.caption(f"Total: **{len(ss_files)}** archivo(s) en esta sesión")

            with st.expander("🗑️ Eliminar archivos"):
                to_delete = st.multiselect(
                    "Selecciona archivos a eliminar",
                    options=list(ss_files.keys()),
                )
                if st.button("Eliminar seleccionados", disabled=not to_delete):
                    for nombre in to_delete:
                        st.session_state[fuentes_key].pop(nombre, None)
                    # Incrementar versión para resetear el widget y no re-cargar archivos eliminados
                    st.session_state[_uploader_v_key] += 1
                    st.rerun()
        else:
            st.info(f"No hay archivos fuente para {zona}/{tipo}. Sube tus .xlsx arriba.")

    st.markdown("---")

    # ── Plantillas ─────────────────────────────────────────────────────────────
    st.subheader("📋 Plantillas de Referencia")
    col_p1, col_p2 = st.columns(2)

    _plantilla_orig_key = f"plantilla_orig_{zona}_{tipo}"
    _estilos_orig_key   = f"estilos_orig_{zona}_{tipo}"

    with col_p1:
        st.markdown("**Plantilla vacía (.xlsm)**")
        if plantilla_path.exists():
            orig = st.session_state.get(_plantilla_orig_key, plantilla_path.name)
            col_nombre, col_btn = st.columns([3, 1])
            col_nombre.success(f"✅ `{orig}`")
            if col_btn.button("🗑️ Quitar", key=f"del_plantilla_{zona}_{tipo}",
                              use_container_width=True):
                plantilla_path.unlink(missing_ok=True)
                st.session_state.pop(_plantilla_orig_key, None)
                st.rerun()
        else:
            st.warning("❌ No cargada")
            p_up = st.file_uploader(
                "Sube la plantilla vacía",
                type=["xlsm"],
                key=f"plantilla_{zona}_{tipo}",
            )
            if p_up:
                guardar_archivo(p_up, zona, tipo, "plantillas",
                                nombre_fijo=f"plantilla_{zona.lower()}_vacia.xlsm")
                st.session_state[_plantilla_orig_key] = p_up.name
                st.success("✅ Plantilla guardada")
                st.rerun()

    with col_p2:
        st.markdown("**Archivo de estilos/referencia (.xlsm)**")
        if estilos_path.exists():
            orig = st.session_state.get(_estilos_orig_key, estilos_path.name)
            col_nombre, col_btn = st.columns([3, 1])
            col_nombre.success(f"✅ `{orig}`")
            if col_btn.button("🗑️ Quitar", key=f"del_estilos_{zona}_{tipo}",
                              use_container_width=True):
                estilos_path.unlink(missing_ok=True)
                st.session_state.pop(_estilos_orig_key, None)
                st.rerun()
        else:
            st.warning("❌ No cargado")
            e_up = st.file_uploader(
                "Sube el archivo de referencia",
                type=["xlsm"],
                key=f"estilos_{zona}_{tipo}",
            )
            if e_up:
                guardar_archivo(e_up, zona, tipo, "plantillas",
                                nombre_fijo=f"estilos_{zona.lower()}_referencia.xlsm")
                st.session_state[_estilos_orig_key] = e_up.name
                st.success("✅ Estilos guardados")
                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Generación
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header(f"⚙️ Generación — {zona} / {tipo} / {MESES[mes_idx]} {anio}")

    ss_files = st.session_state.get(fuentes_key, {})
    n_fuentes = len(ss_files)
    can_run = (
        n_fuentes > 0
        and plantilla_path.exists()
        and estilos_path.exists()
    )

    # Estado
    col_st, col_acc = st.columns([3, 1])
    with col_st:
        st.subheader("Estado actual")
        def check(cond, label):
            return f"{'✅' if cond else '❌'} {label}"

        st.write(check(n_fuentes > 0,
                       f"Archivos fuente: **{n_fuentes}** archivo(s)"))
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
        import tempfile as _tempfile

        # Escribir bytes de session_state a /tmp/ para que generar_plantilla pueda leerlos
        _tmp_dir = Path(_tempfile.mkdtemp())
        source_paths = []
        for _nombre, _data in ss_files.items():
            _p = _tmp_dir / _nombre
            _p.write_bytes(_data)
            source_paths.append(str(_p))

        log_area    = st.empty()
        progress_bar = st.progress(0.0)
        all_logs = []

        st.session_state[logs_key] = all_logs

        def on_log(msg, pct):
            all_logs.append(msg)
            log_area.code("\n".join(all_logs[-30:]))
            progress_bar.progress(min(pct, 1.0))

        try:
            logs, total_rows = generar_plantilla(
                zona=zona,
                tipo=tipo,
                source_paths=source_paths,
                plantilla_path=str(plantilla_path),
                estilos_path=str(estilos_path),
                output_path=str(output_path),
                log_callback=on_log,
            )
            st.success(
                f"✅ **{output_fname}** generado correctamente — {total_rows} filas"
            )
            if any("Sin datos" in l or "SIN MAPEAR" in l or "IGNORADO" in l for l in all_logs):
                st.info("Hubo observaciones en esta generación. Revisa la pestaña "
                        "**💬 Asistente** para ver qué pasó y cómo corregirlo.")
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
# TAB 4 — Asistente (sin IA: reglas sobre el log y los datos generados)
# ══════════════════════════════════════════════════════════════════════════════
# (Se ejecuta antes que Conciliación porque esa pestaña usa st.stop(), que corta
#  el resto de la página cuando no hay archivo generado.)
@st.cache_data(show_spinner=False)
def _cargar_datos_asistente(path: str, mtime: float):
    return asistente.cargar_datos(path)


def _mostrar_hallazgos(hallazgos):
    if not hallazgos:
        st.info("Todavía no hay un log en esta sesión. Genera el archivo en la pestaña "
                "**⚙️ Generación** y vuelve aquí.")
        return
    for h in hallazgos:
        caja = {"error": st.error, "aviso": st.warning, "ok": st.success}[h.nivel]
        caja(f"**{h.titulo}**\n\n{h.detalle}\n\n**Qué hacer:** {h.que_hacer}")
        if h.items:
            with st.expander(f"Ver detalle ({len(h.items)})"):
                for it in h.items:
                    st.markdown(f"- {it}")


def _soles(x: float) -> str:
    return f"S/ {x:,.2f}"


with tab4:
    st.header(f"💬 Asistente — {zona} / {tipo} / {MESES[mes_idx]} {anio}")
    st.caption("Responde con el log de la última generación y el archivo generado de "
               "esta zona, tipo y mes. Funciona sin IA y sin costo.")

    logs_sesion = st.session_state.get(logs_key, [])
    hallazgos = asistente.analizar_log(logs_sesion)
    df_out = None
    if output_path.exists():
        try:
            df_out = _cargar_datos_asistente(str(output_path), output_path.stat().st_mtime)
        except Exception as exc:
            st.warning(f"No pude leer `{output_fname}`: {exc}")

    if df_out is None or df_out.empty:
        st.info(f"Aún no existe `{output_fname}` con datos. Las preguntas sobre montos "
                "y materiales se activan después de generar.")

    # ── Pregunta libre ────────────────────────────────────────────────────────
    duda = st.text_input(
        "Escribe tu duda",
        placeholder="Ej.: ¿cuánto suma P-642?  ·  ¿qué salió con precio 0?  ·  escobilla  ·  ¿por qué sin datos?",
    )

    if duda.strip():
        q = asistente.normalize(duda)
        respondido = False
        hay_datos = df_out is not None and not df_out.empty

        # 1) Un CR o Pozo concreto
        m = re.search(r"\b(cr|p|pozo)[\s\-]*(\d+[a-z]?)\b", q)
        if m and hay_datos:
            pref = "CR" if m.group(1) == "cr" else "P"
            cod = f"{pref}-{m.group(2).upper()}"
            tot = asistente.totales_por_ubicacion(df_out)
            if cod in tot.index:
                r = tot.loc[cod]
                st.success(
                    f"**{cod}** suma **{_soles(r['Total'])}** con {int(r['Equipos'])} equipo(s): "
                    f"Bomba {_soles(r['Bomba'])} · Tablero {_soles(r['Tablero'])} · "
                    f"Accesorios {_soles(r['Accesorios'])}."
                )
                st.dataframe(asistente.detalle_ubicacion(df_out, cod),
                             use_container_width=True, hide_index=True)
            else:
                motivo = next((h for h in hallazgos if cod in h.titulo), None)
                if motivo:
                    _mostrar_hallazgos([motivo])
                else:
                    st.warning(f"{cod} no está en `{output_fname}`. Revisa que su archivo "
                               "fuente esté cargado y que la Zona y el Tipo sean los correctos.")
            respondido = True

        # 2) Temas generales
        if not respondido and any(k in q for k in ("sin datos", "error", "problema", "fallo",
                                                    "que paso", "ignorad", "observacion")):
            _mostrar_hallazgos(hallazgos)
            respondido = True
        if not respondido and hay_datos and any(k in q for k in ("precio 0", "precio cero", "sin precio",
                                                                  "sin sap", "sin codigo", "en cero")):
            pc = asistente.precio_cero(df_out)
            if pc.empty:
                st.success("Ninguna fila de material o ACT salió con precio 0 ni sin código SAP.")
            else:
                st.warning(f"{len(pc)} fila(s) con precio 0 o sin código SAP:")
                st.dataframe(pc, use_container_width=True, hide_index=True)
            respondido = True
        if not respondido and hay_datos and any(k in q for k in ("mas usad", "ranking", "top", "mas caro",
                                                                  "mayor gasto")):
            st.dataframe(asistente.ranking_materiales(df_out).head(15),
                         use_container_width=True, hide_index=True)
            respondido = True
        if not respondido and hay_datos and any(k in q for k in ("total", "suma", "monto", "cuanto")):
            rs = asistente.resumen(df_out)
            st.success(
                f"El mes suma **{_soles(rs['total'])}** en {rs['ubicaciones']} ubicación(es) y "
                f"{rs['equipos']} equipo(s): Bomba {_soles(rs['bomba'])} · "
                f"Tablero {_soles(rs['tablero'])} · Accesorios {_soles(rs['accesorios'])}."
            )
            respondido = True

        # 3) Un material
        if not respondido and hay_datos:
            res = asistente.buscar_material(df_out, duda)
            if not res.empty:
                st.success(f"'{duda.strip()}' aparece en {res['Ubicación'].nunique()} ubicación(es), "
                           f"cantidad total {res['Cantidad'].sum():,.2f}, "
                           f"monto {_soles(res['Subtotal'].sum())}.")
                st.dataframe(res, use_container_width=True, hide_index=True)
                respondido = True

        # 4) Preguntas frecuentes
        faq = asistente.buscar_faq(duda)
        if faq:
            if respondido:
                st.markdown("**También puede servirte:**")
            for preg, resp in faq:
                with st.expander(preg, expanded=not respondido):
                    st.write(resp)
            respondido = True

        if not respondido:
            st.info("No encontré una respuesta para eso. Prueba con el código de un CR o Pozo "
                    "(ej. P-642), el nombre de un material, o elige una pregunta de abajo.")

    st.markdown("---")

    # ── Preguntas rápidas ─────────────────────────────────────────────────────
    opcion = st.radio(
        "O elige una pregunta",
        [
            "¿Qué pasó en la última generación?",
            "¿Cuánto suma cada CR / Pozo?",
            "¿Qué salió con precio 0 o sin código SAP?",
            "¿Qué materiales pesan más en el monto?",
            "Ver detalle de un CR / Pozo",
            "Preguntas frecuentes",
        ],
        horizontal=True,
        label_visibility="visible",
    )

    sin_datos_msg = "Esta respuesta necesita el archivo generado. Genera primero en **⚙️ Generación**."

    if opcion == "¿Qué pasó en la última generación?":
        _mostrar_hallazgos(hallazgos)
        if logs_sesion:
            with st.expander("Ver log completo"):
                st.code("\n".join(logs_sesion))

    elif opcion == "¿Cuánto suma cada CR / Pozo?":
        if df_out is None or df_out.empty:
            st.info(sin_datos_msg)
        else:
            rs = asistente.resumen(df_out)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total del mes", _soles(rs["total"]))
            c2.metric("Ubicaciones", rs["ubicaciones"])
            c3.metric("Equipos", rs["equipos"])
            c4.metric("Filas", rs["filas"])
            st.dataframe(
                asistente.totales_por_ubicacion(df_out).style.format({
                    "Bomba": "S/ {:,.2f}", "Tablero": "S/ {:,.2f}",
                    "Accesorios": "S/ {:,.2f}", "Total": "S/ {:,.2f}", "Equipos": "{:.0f}",
                }),
                use_container_width=True,
            )

    elif opcion == "¿Qué salió con precio 0 o sin código SAP?":
        if df_out is None or df_out.empty:
            st.info(sin_datos_msg)
        else:
            pc = asistente.precio_cero(df_out)
            if pc.empty:
                st.success("Ninguna fila de material o ACT salió con precio 0 ni sin código SAP.")
            else:
                st.warning(f"{len(pc)} fila(s) con precio 0 o sin código SAP. "
                           "Revisa las preguntas frecuentes para corregirlas.")
                st.dataframe(pc, use_container_width=True, hide_index=True)

    elif opcion == "¿Qué materiales pesan más en el monto?":
        if df_out is None or df_out.empty:
            st.info(sin_datos_msg)
        else:
            por = st.radio("Ordenar por", ["Subtotal", "Cantidad", "Usos"], horizontal=True)
            st.dataframe(asistente.ranking_materiales(df_out, por=por).head(20),
                         use_container_width=True, hide_index=True)

    elif opcion == "Ver detalle de un CR / Pozo":
        if df_out is None or df_out.empty:
            st.info(sin_datos_msg)
        else:
            tot = asistente.totales_por_ubicacion(df_out)
            cod = st.selectbox("Ubicación", sorted(tot.index))
            r = tot.loc[cod]
            st.success(f"**{cod}** suma **{_soles(r['Total'])}** con {int(r['Equipos'])} equipo(s).")
            st.dataframe(asistente.detalle_ubicacion(df_out, cod),
                         use_container_width=True, hide_index=True)

    else:
        for preg, resp in asistente.FAQ:
            with st.expander(preg):
                st.write(resp)


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

    # Etiqueta visible según el tipo (la columna interna sigue llamándose "CR")
    etiqueta  = "Pozo" if tipo == "POZO" else "CR"
    etiquetas = "Pozos" if tipo == "POZO" else "CRs"
    prefijo   = "P" if tipo == "POZO" else "CR"

    crs_detectados = sorted(totales_reales.keys())

    if not crs_detectados:
        st.warning(f"El archivo generado no contiene filas PM03 de {etiquetas} legibles.")
        st.stop()

    st.caption(f"{etiquetas} detectados en el archivo: **{', '.join(crs_detectados)}**")
    st.markdown("---")

    # ── Ingresar montos esperados ──────────────────────────────────────────────
    st.subheader("📥 Montos Esperados / Teóricos")
    st.caption(
        f"Ingresa el monto total esperado por {etiqueta}. "
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
                "CR": st.column_config.TextColumn(etiqueta, disabled=True, width="small"),
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
                col_cr  = next((c for c in df_imp.columns
                                if any(k in str(c).lower() for k in ("cr", "pozo", "ubic", "codigo", "código"))), None)
                col_mon = next((c for c in df_imp.columns if "monto" in c.lower() or
                                "esperado" in c.lower() or "teorico" in c.lower()), None)
                if col_cr and col_mon:
                    imp_dict = {
                        normalizar_codigo(k, prefijo): float(v)
                        for k, v in zip(df_imp[col_cr], df_imp[col_mon])
                        if pd.notna(k) and pd.notna(v)
                    }
                    df_base = st.session_state[session_key].copy()
                    df_base["Monto Esperado (S/)"] = df_base["CR"].map(imp_dict).fillna(0.0)
                    st.session_state[session_key] = df_base
                    st.success("✅ Importado correctamente")
                    st.rerun()
                else:
                    st.error(f"El archivo debe tener columnas '{etiqueta}' y 'Monto Esperado'.")
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
        m4.metric(f"{etiquetas} conciliados", len(df_res))
        df_res = df_res.rename(columns={"CR": etiqueta})

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
