import streamlit as st
import pandas as pd
import unicodedata
import re
import fitz  # PyMuPDF
from difflib import SequenceMatcher
import zipfile
import io

# Configuración de página
st.set_page_config(
    page_title="Ayuda vagos PRO MAX 🚀", 
    page_icon="🚀", 
    layout="wide"
)

st.title("Ayuda vagos PRO MAX 🚀")
st.subheader("Extractor de comprobantes de Recuperadores independientes")

SPREADSHEET_ID = "1LSYGE6PcvjaaYlgHK3iw6_WneP-jTcBG8S9fdfDMxFA"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv"
SHEET_PUBLIC_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit?usp=sharing"

if "reset_count" not in st.session_state:
    st.session_state["reset_count"] = 0

rc = st.session_state["reset_count"]

def parsear_monto(val):
    """Parsea montos en formato argentino a float puro evitando truncamientos."""
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace("$", "").replace("\xa0", "").strip()
    if not s or "total" in s.lower() or "suma" in s.lower():
        return 0.0
    
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        partes = s.split(".")
        if len(partes[-1]) == 3 or len(partes) > 2:
            s = s.replace(".", "")
            
    try:
        return float(s)
    except:
        return 0.0

def formatear_monto_ar(val):
    """Formatea float a formato moneda argentina: 61200.0 -> $ 61.200,00"""
    return f"$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def normalizar_texto(texto):
    if not texto:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto).upper())
    sin_tildes = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^A-Z0-9\s]', '', sin_tildes).strip()

def similar_palabra_banco(p_planilla, p_pdf):
    """Verifica si una palabra del nombre coincide o es muy similar (ej: HERNAN vs HERNN)"""
    if p_planilla in p_pdf or p_pdf in p_planilla:
        return True
    return SequenceMatcher(None, p_planilla, p_pdf).ratio() >= 0.80

@st.cache_data(ttl=300)
def cargar_datos_google_sheets(url):
    try:
        df_raw = pd.read_csv(url, header=None, dtype=str)
        if df_raw.empty:
            return pd.DataFrame()

        fila_encabezado = None
        for idx, row in df_raw.head(10).iterrows():
            texto_fila = " ".join([str(val).lower() for val in row.values if pd.notna(val)])
            if "cuit" in texto_fila and ("fecha" in texto_fila or "unidad" in texto_fila):
                fila_encabezado = idx
                break
        
        if fila_encabezado is None:
            return pd.DataFrame()

        df_clean = pd.read_csv(url, header=fila_encabezado, dtype=str)
        df_clean.columns = [str(c).strip() for c in df_clean.columns]
        
        col_fecha = next((c for c in df_clean.columns if "fecha" in c.lower()), None)
        col_up = next((c for c in df_clean.columns if "unidad" in c.lower() or "up" in c.lower()), None)
        col_cuit = next((c for c in df_clean.columns if "cuit" in c.lower()), None)
        col_nombre = next((c for c in df_clean.columns if "apellido" in c.lower() or "nombre" in c.lower()), None)
        col_monto = next((c for c in df_clean.columns if "total" in c.lower() or "monto" in c.lower()), None)
        
        if not all([col_fecha, col_up, col_cuit, col_nombre, col_monto]):
            return pd.DataFrame()
            
        df_final = df_clean[[col_fecha, col_up, col_cuit, col_nombre, col_monto]].copy()
        df_final.columns = ["Fecha", "Unidad_Productiva", "CUIT", "Nombre", "Monto"]
        
        df_final = df_final.dropna(subset=["CUIT", "Fecha"]).copy()
        df_final["CUIT"] = df_final["CUIT"].astype(str).str.replace(".0", "", regex=False).str.strip()
        df_final["Unidad_Productiva"] = df_final["Unidad_Productiva"].astype(str).str.strip()
        df_final["Nombre"] = df_final["Nombre"].astype(str).str.strip()
        df_final["Monto"] = df_final["Monto"].apply(parsear_monto)
        df_final["Fecha"] = df_final["Fecha"].astype(str).str.strip()
        
        df_final = df_final[~df_final["CUIT"].str.upper().str.contains("TOTAL|SUMA|NAN|UNIDAD", na=False)]
        df_final = df_final[df_final["Monto"] > 0]
        
        return df_final
    except Exception as e:
        return pd.DataFrame()

def extraer_fecha_pago_pdf(doc):
    for i in range(min(5, len(doc))):
        texto = doc[i].get_text("text")
        match = re.search(r"Fecha de Pago:\s*([\d{2}/-]+)", texto, re.IGNORECASE)
        if match:
            return match.group(1).replace("-", "/").strip()
        match_fecha = re.search(r"(\d{2}[/-]\d{2}[/-]\d{4})", texto)
        if match_fecha:
            return match_fecha.group(1).replace("-", "/").strip()
    return None

def normalizar_fecha_busqueda(fecha_str):
    if not fecha_str:
        return ""
    partes = fecha_str.replace("-", "/").split("/")
    if len(partes) == 3:
        try:
            d = str(int(partes[0]))
            m = str(int(partes[1]))
            a = partes[2]
            return f"{d}/{m}/{a}"
        except:
            pass
    return fecha_str

def sanitizar_texto_archivo(texto):
    return str(texto).replace("/", "-").replace("\\", "-").replace(":", "-").strip()

def generar_pdf_individual_bytes(doc, pag_idx):
    single_doc = fitz.open()
    single_doc.insert_pdf(doc, from_page=pag_idx, to_page=pag_idx)
    pdf_bytes = single_doc.write()
    single_doc.close()
    return pdf_bytes

def generar_pdf_unificado_bytes(doc, lista_pag_indices):
    uni_doc = fitz.open()
    for p_idx in lista_pag_indices:
        uni_doc.insert_pdf(doc, from_page=p_idx, to_page=p_idx)
    pdf_bytes = uni_doc.write()
    uni_doc.close()
    return pdf_bytes

# --- MÓDULO 1 ---
c_link, c_ref, _ = st.columns([2, 1, 3])
with c_link:
    st.link_button("📊 Abrir Planilla Maestra en Google Sheets", SHEET_PUBLIC_URL)
with c_ref:
    if st.button("🔄 Sincronizar"):
        st.cache_data.clear()
        st.rerun()

with st.spinner("Sincronizando con Google Sheets..."):
    df_master = cargar_datos_google_sheets(CSV_URL)

if df_master.empty:
    st.error("⚠️ No se pudieron sincronizar los datos de la planilla maestra.")
    st.stop()
else:
    st.success(f"✅ Conexión exitosa. Se sincronizaron **{len(df_master)} registros** desde la nube.")

st.divider()

# --- MÓDULO 2 ---
st.markdown("### Carga de PDF unificado")
uploaded_pdf = st.file_uploader("Cargá el pdf con los comprobantes de transferencia", type=["pdf"], key=f"pdf_up_{rc}")

ups_seleccionadas = []
fecha_busqueda = ""
fecha_display = ""
doc = None
df_fecha = pd.DataFrame()

if uploaded_pdf is not None:
    try:
        pdf_bytes = uploaded_pdf.read()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        fecha_detectada = extraer_fecha_pago_pdf(doc)
        
        if not fecha_detectada:
            st.warning("⚠️ No se pudo detectar automáticamente la Fecha de Pago en el PDF.")
        else:
            fecha_display = fecha_detectada.replace("-", "/")
            fecha_busqueda = normalizar_fecha_busqueda(fecha_display)
            
            st.info(f"📅 **Fecha detectada:** {fecha_display}")
            
            df_fecha = df_master[df_master["Fecha"].apply(lambda x: normalizar_fecha_busqueda(x) == fecha_busqueda)].copy()
            
            if df_fecha.empty:
                st.error(f"❌ No se encontraron registros en Google Sheets para la fecha `{fecha_busqueda}`.")
            else:
                ups_disponibles = sorted(df_fecha["Unidad_Productiva"].unique())
                st.success(f"🎯 Se detectaron **{len(ups_disponibles)} Unidades Productivas** disponibles para esta fecha.")
                
                st.markdown("#### 🏢 Seleccioná las Unidades Productivas a procesar:")
                with st.container(border=True):
                    def toggle_all_ups():
                        estado_maestro = st.session_state.master_select_all
                        for up_item in ups_disponibles:
                            st.session_state[f"chk_up_{up_item}"] = estado_maestro

                    st.checkbox("📌 **Seleccionar / Desseleccionar todas**", value=False, key="master_select_all", on_change=toggle_all_ups)
                    st.divider()
                    
                    for up in ups_disponibles:
                        df_up_temp = df_fecha[df_fecha["Unidad_Productiva"] == up]
                        cant_personas = len(df_up_temp)
                        monto_total_up = df_up_temp["Monto"].sum()
                        monto_fmt = formatear_monto_ar(monto_total_up)
                        
                        if st.checkbox(f"**{up}** — *({cant_personas} pagos | Total: {monto_fmt})*", value=False, key=f"chk_up_{up}"):
                            ups_seleccionadas.append(up)
    except Exception as e:
        st.error(f"Error al analizar el PDF: {e}")

# Botón Procesar con Contador Dinámico
cant_sel = len(ups_seleccionadas)
sufijo_sel = "seleccionado" if cant_sel == 1 else "seleccionados"
texto_btn_procesar = f"🚀 Procesar {cant_sel} {sufijo_sel}"

procesar_cliqueado = False
if doc and ups_seleccionadas:
    if st.button(texto_btn_procesar, type="primary"):
        procesar_cliqueado = True

st.divider()

# --- MOTOR DE BÚSQUEDA ESTRICTO BASADO EN MONTO (Ayuda Vagos 3002) ---
if procesar_cliqueado or "resultados_lotes" in st.session_state:
    if procesar_cliqueado:
        with st.spinner("Ejecutando motor estricto por monto y validación de nombres..."):
            try:
                paginas_datos = []
                for i_pag in range(len(doc)):
                    texto_raw = doc[i_pag].get_text("text")
                    texto_norm = normalizar_texto(texto_raw)
                    palabras_pagina = set(texto_norm.split())
                    
                    monto_pag = 0.0
                    match_m = re.search(r"IMPORTE DE LA TRANSFERENCIA:\s*\$?([\d\.,]+)", texto_raw, re.IGNORECASE)
                    if match_m:
                        monto_pag = parsear_monto(match_m.group(1))
                    
                    paginas_datos.append({
                        "pag_idx": i_pag,
                        "texto_norm": texto_norm,
                        "palabras_pagina": palabras_pagina,
                        "monto": monto_pag
                    })
                
                resultados_lotes = {}
                
                for up_nombre in ups_seleccionadas:
                    df_listado = df_fecha[df_fecha["Unidad_Productiva"] == up_nombre].reset_index(drop=True)
                    registros_procesados = []
                    
                    for idx, row in df_listado.iterrows():
                        cuit_orig = str(row.get("CUIT", "")).strip()
                        nombre_orig = str(row.get("Nombre", "Sin Nombre")).strip()
                        nombre_norm = normalizar_texto(nombre_orig)
                        monto_val = float(row.get("Monto", 0))
                        
                        hallado = False
                        num_pagina_encontrada = None
                        pag_idx_encontrado = None
                        
                        palabras_nombre = [p for p in nombre_norm.split() if len(p) >= 3]
                        
                        # REGLA DE ORO: EL MONTO ES CONDICIÓN OBLIGATORIA
                        if monto_val > 0:
                            for pdata in paginas_datos:
                                if abs(pdata["monto"] - monto_val) <= 0.5:
                                    coinc_tokens = sum(1 for pn in palabras_nombre if any(similar_palabra_banco(pn, pp) for pp in pdata["palabras_pagina"]))
                                    
                                    if not palabras_nombre or coinc_tokens >= min(1, len(palabras_nombre)):
                                        hallado = True
                                        num_pagina_encontrada = pdata["pag_idx"] + 1
                                        pag_idx_encontrado = pdata["pag_idx"]
                                        break
                        
                        monto_fmt_str = formatear_monto_ar(monto_val)
                        
                        registros_procesados.append({
                            "df_idx": idx,
                            "CUIT": cuit_orig,
                            "Nombre": nombre_orig,
                            "Monto": monto_fmt_str,
                            "Estado": "✅ Encontrado" if hallado else "❌ No encontrado",
                            "Página PDF": num_pagina_encontrada if hallado else "-",
                            "Pag_Index": pag_idx_encontrado,
                            "Encontrado": hallado
                        })
                    
                    resultados_lotes[up_nombre] = registros_procesados
                
                st.session_state["doc_pdf"] = doc
                st.session_state["resultados_lotes"] = resultados_lotes
                st.session_state["ups_seleccionadas_activas"] = ups_seleccionadas
                st.session_state["fecha_pago_str"] = fecha_display.replace("/", "-")
            except Exception as e:
                st.error(f"Error en el procesamiento: {e}")

    # --- DESCARGA DE COMPROBANTES ---
    if "resultados_lotes" in st.session_state:
        st.markdown("### 📥 Descarga de comprobantes")
        doc_actual = st.session_state["doc_pdf"]
        fecha_str = st.session_state["fecha_pago_str"]
        
        keys_lotes = list(st.session_state["resultados_lotes"].keys())
        titulos_tabs = []
        for k in keys_lotes:
            registros = st.session_state["resultados_lotes"][k]
            total_regs = len(registros)
            encontrados_regs = sum(1 for r in registros if r["Encontrado"])
            icono = "✅" if encontrados_regs == total_regs else "⚠️"
            titulos_tabs.append(f"{icono} {k} ({encontrados_regs}/{total_regs})")
            
        tabs_lotes = st.tabs(titulos_tabs)
        
        for idx_tab, up_nombre in enumerate(keys_lotes):
            with tabs_lotes[idx_tab]:
                registros = st.session_state["resultados_lotes"][up_nombre]
                up_sanitizada = sanitizar_texto_archivo(up_nombre)
                
                # Encabezado
                header_cols = st.columns([0.6, 1.8, 2.5, 1.3, 1.3, 1.2, 1.2])
                header_cols[0].markdown("**#**")
                header_cols[1].markdown("**CUIT**")
                header_cols[2].markdown("**Nombre**")
                header_cols[3].markdown("**Monto**")
                header_cols[4].markdown("**Estado**")
                header_cols[5].markdown("**Pág. PDF**")
                header_cols[6].markdown("**Acción**")
                st.divider()
                
                # Filas
                for i, reg in enumerate(registros):
                    row_cols = st.columns([0.6, 1.8, 2.5, 1.3, 1.3, 1.2, 1.2])
                    row_cols[0].write(str(i + 1))
                    row_cols[1].write(reg["CUIT"])
                    row_cols[2].write(reg["Nombre"])
                    row_cols[3].write(reg["Monto"])
                    row_cols[4].write(reg["Estado"])
                    row_cols[5].write(str(reg["Página PDF"]))
                    
                    with row_cols[6]:
                        if reg["Encontrado"] and reg["Pag_Index"] is not None:
                            ind_bytes = generar_pdf_individual_bytes(doc_actual, reg["Pag_Index"])
                            nombre_persona_san = sanitizar_texto_archivo(reg['Nombre'])
                            nombre_archivo_ind = f"Comprobante_RI_{up_sanitizada}_{nombre_persona_san}_{fecha_str}.pdf"
                            st.download_button(
                                label="⬇️ PDF",
                                data=ind_bytes,
                                file_name=nombre_archivo_ind,
                                mime="application/pdf",
                                key=f"dl_ind_{up_nombre}_{i}"
                            )
                        else:
                            st.write("—")
                
                st.divider()
                
                # Opciones de descarga de UP
                st.markdown("#### 📥 Opciones de Descarga para esta Unidad Productiva")
                
                registros_validos = [r for r in registros if r["Encontrado"] and r["Pag_Index"] is not None]
                
                if registros_validos:
                    col_d1, col_d2 = st.columns(2)
                    
                    with col_d1:
                        indices_up = [r["Pag_Index"] for r in registros_validos]
                        pdf_uni_bytes = generar_pdf_unificado_bytes(doc_actual, indices_up)
                        nombre_archivo_uni = f"Comprobantes_RI_{up_sanitizada}_{fecha_str}.pdf"
                        st.download_button(
                            label=f"📄 Descargar PDF Único ({up_nombre})",
                            data=pdf_uni_bytes,
                            file_name=nombre_archivo_uni,
                            mime="application/pdf",
                            key=f"dl_uni_{up_nombre}"
                        )
                    
                    with col_d2:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for r in registros_validos:
                                p_idx = r["Pag_Index"]
                                n_per = sanitizar_texto_archivo(r["Nombre"])
                                nombre_f = f"Comprobante_RI_{up_sanitizada}_{n_per}_{fecha_str}.pdf"
                                single_bytes = generar_pdf_individual_bytes(doc_actual, p_idx)
                                zip_file.writestr(nombre_f, single_bytes)
                        zip_buffer.seek(0)
                        
                        st.download_button(
                            label=f"🗂️ Descargar ZIP Separados ({up_nombre})",
                            data=zip_buffer,
                            file_name=f"ZIP_Comprobantes_{up_sanitizada}_{fecha_str}.zip",
                            mime="application/zip",
                            key=f"dl_zip_{up_nombre}"
                        )
                else:
                    st.info("No hay comprobantes encontrados para generar descargas en esta UP.")

        # Descarga Masiva Multi-UP
        st.markdown("---")
        ups_activas_count = len(st.session_state.get("ups_seleccionadas_activas", keys_lotes))
        
        zip_masivo_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_masivo_buffer, "w", zipfile.ZIP_DEFLATED) as zip_masivo:
            for up_n, regs in st.session_state["resultados_lotes"].items():
                up_san = sanitizar_texto_archivo(up_n)
                for r in regs:
                    if r["Encontrado"] and r["Pag_Index"] is not None:
                        p_idx = r["Pag_Index"]
                        n_per = sanitizar_texto_archivo(r["Nombre"])
                        nombre_f = f"Comprobante_RI_{up_san}_{n_per}_{fecha_str}.pdf"
                        single_bytes = generar_pdf_individual_bytes(doc_actual, p_idx)
                        path_en_zip = f"{up_n}/{nombre_f}"
                        zip_masivo.writestr(path_en_zip, single_bytes)
        zip_masivo_buffer.seek(0)
        
        st.download_button(
            label=f"📦 Descargar comprobantes de {ups_activas_count} UPs seleccionadas",
            data=zip_masivo_buffer,
            file_name=f"ZIP_Masivo_Todas_UPs_{fecha_str}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True
        )

# --- BOTÓN DE REINICIO ABAJO DE TODO ---
st.markdown("---")
col_reset1, col_reset2, col_reset3 = st.columns([1, 2, 1])
with col_reset2:
    if st.button("🔄 Reiniciar aplicación", type="secondary", use_container_width=True):
        st.session_state["reset_count"] += 1
        keys_to_clear = ["doc_pdf", "resultados_lotes", "ups_seleccionadas_activas", "fecha_pago_str"]
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()