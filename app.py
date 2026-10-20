import streamlit as st
import pandas as pd
import io
import fitz  # PyMuPDF
import unicodedata
import re
import zipfile
from difflib import SequenceMatcher

# Configuración de página con el nombre oficial V2
st.set_page_config(page_title="Ayuda vagos 3002 😎", page_icon="💸", layout="wide")

st.title("Ayuda vagos 3002 😎")
st.subheader("Extractor automático de comprobantes de transferencia (Procesamiento por Lotes y Manual)")

if "reset_count" not in st.session_state:
    st.session_state["reset_count"] = 0

def parsear_monto(val):
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace("$", "").replace("\xa0", "").strip()
    if not s or "total" in s.lower() or "suma" in s.lower():
        return 0.0
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except:
        return 0.0

def normalizar_texto(texto):
    if not texto:
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto).upper())
    sin_tildes = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return sin_tildes.replace("Ñ", "N")

def calcular_similitud(a, b):
    return SequenceMatcher(None, a, b).ratio()

def extraer_datos_pagina(texto_pag):
    monto_encontrado = "No detectado"
    titular_encontrado = "No detectado"
    
    match_monto = re.search(r"IMPORTE DE LA TRANSFERENCIA:\s*\$?([\d\.,]+)", texto_pag, re.IGNORECASE)
    if match_monto:
        monto_encontrado = match_monto.group(1).strip()
        
    lineas = [l.strip() for l in texto_pag.split("\n") if l.strip()]
    for i, linea in enumerate(lineas):
        if "BENEFICIARIO:" in linea.upper() or "TITULAR" in linea.upper():
            if ":" in linea:
                partes = linea.split(":")
                if len(partes) > 1 and len(partes[1].strip()) > 3:
                    titular_encontrado = partes[1].strip()
                    break
            if i + 1 < len(lineas):
                posible_nombre = lineas[i + 1]
                if posible_nombre and not posible_nombre.startswith("FAC") and len(posible_nombre) > 3:
                    titular_encontrado = posible_nombre
                    break
                    
    return titular_encontrado, monto_encontrado

def extraer_unidad_productiva(df_raw):
    for r in range(min(5, len(df_raw))):
        fila = df_raw.iloc[r].tolist()
        for c, val in enumerate(fila):
            val_str = str(val).strip().lower()
            if "unidad productiva" in val_str:
                if c + 1 < len(fila) and pd.notna(fila[c + 1]):
                    nombre_up = str(fila[c + 1]).strip()
                    if nombre_up and nombre_up.lower() != "nan":
                        return nombre_up
    return ""

def extraer_fecha_pago_pdf(doc):
    for i in range(min(5, len(doc))):
        texto = doc[i].get_text("text")
        match = re.search(r"Fecha de Pago:\s*([\d{2}/]+)", texto, re.IGNORECASE)
        if match:
            return match.group(1).replace("/", "-").strip()
        match_fecha = re.search(r"(\d{2}/\d{2}/\d{4})", texto)
        if match_fecha:
            return match_fecha.group(1).replace("/", "-").strip()
    return "Fecha_Desconocida"

def procesar_hoja_excel(xls, hoja_seleccionada):
    df_raw = pd.read_excel(xls, sheet_name=hoja_seleccionada, header=None)
    unidad_productiva = extraer_unidad_productiva(df_raw)
    
    fila_encabezado = None
    for idx, row in df_raw.iterrows():
        texto_fila = " ".join([str(val).lower() for val in row.values])
        if "cuit" in texto_fila and ("nombre" in texto_fila or "apellido" in texto_fila):
            fila_encabezado = idx
            break
            
    if fila_encabezado is None:
        return None, unidad_productiva

    raw_headers = df_raw.iloc[fila_encabezado].tolist()
    headers = []
    seen = {}
    for h in raw_headers:
        h_str = str(h).strip()
        if not h_str or h_str.lower() in ["nan", "none"]:
            h_str = "Col_Vacia"
        if h_str in seen:
            seen[h_str] += 1
            h_str = f"{h_str}_{seen[h_str]}"
        else:
            seen[h_str] = 0
        headers.append(h_str)

    df_datos = df_raw.iloc[fila_encabezado + 1:].copy()
    df_datos.columns = headers

    col_cuit = next((c for c in df_datos.columns if "cuit" in c.lower()), None)
    col_nombre = next((c for c in df_datos.columns if "apellido" in c.lower() or "nombre" in c.lower()), None)
    col_monto = next((c for c in df_datos.columns if any(m in c.lower() for m in ["monto", "valorizado", "total"])), None)

    cols_a_conservar = [c for c in [col_cuit, col_nombre, col_monto] if c is not None]
    if not cols_a_conservar:
        return None, unidad_productiva

    df_filtrado = df_datos.dropna(subset=[col_cuit]).copy()
    if col_cuit:
        df_filtrado[col_cuit] = df_filtrado[col_cuit].astype(str).str.replace(".0", "", regex=False).str.strip()
        df_filtrado = df_filtrado[~df_filtrado[col_cuit].str.upper().str.contains("TOTAL|SUMA|SUBTOTAL|NAN", na=False)]
    if col_nombre:
        df_filtrado = df_filtrado[~df_filtrado[col_nombre].astype(str).str.upper().str.contains("TOTAL|SUMA|SUBTOTAL", na=False)]
    if col_monto:
        df_filtrado[col_monto] = df_filtrado[col_monto].apply(parsear_monto)

    df_filtrado = df_filtrado[df_filtrado[col_monto] > 0]

    if col_cuit and col_monto and df_filtrado[col_cuit].duplicated().any():
        dic_agrupacion = {col_nombre: "first"} if col_nombre else {}
        dic_agrupacion[col_monto] = "sum"
        df_filtrado = df_filtrado.groupby(col_cuit, as_index=False).agg(dic_agrupacion)
    elif col_cuit:
        df_filtrado = df_filtrado.drop_duplicates(subset=[col_cuit])

    mapeo_renombre = {}
    if col_cuit: mapeo_renombre[col_cuit] = "CUIT"
    if col_nombre: mapeo_renombre[col_nombre] = "Nombre"
    if col_monto: mapeo_renombre[col_monto] = "Monto"
    df_filtrado = df_filtrado.rename(columns=mapeo_renombre)

    return df_filtrado, unidad_productiva

config_columnas = {
    "#": st.column_config.Column("#", width=40),
    "CUIT": st.column_config.TextColumn("CUIT", width=150),
    "Nombre": st.column_config.TextColumn("Nombre", width=300),
    "Monto": st.column_config.TextColumn("Monto", width=150),
    "Página PDF": st.column_config.NumberColumn("Página PDF", width=100, format="%d")
}

rc = st.session_state["reset_count"]

# --- SECCIÓN 1: SELECCIÓN DE MODO (LOTE O MANUAL) ---
st.markdown("### 1. Cargar el listado de personas / transferencias")

modo_carga = st.radio(
    "Seleccioná el método de carga:", 
    ["📁 Carga Múltiple de Excels (Lote)", "📋 Pegar datos manualmente"],
    horizontal=True,
    key=f"modo_carga_{rc}"
)

dict_lotes = {}
uploaded_pdf = None

if modo_carga == "📁 Carga Múltiple de Excels (Lote)":
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        uploaded_excels = st.file_uploader(
            "📁 Subí una o varias planillas Excel/CSV", 
            type=["xlsx", "xls", "csv"], 
            accept_multiple_files=True,
            key=f"excel_uploader_{rc}"
        )
    with col_l2:
        uploaded_pdf = st.file_uploader(
            "📄 Subí el PDF único con los comprobantes", 
            type=["pdf"], 
            key=f"pdf_uploader_{rc}"
        )

    if uploaded_excels:
        st.markdown("#### 📋 Configuración de Planillas Cargadas")
        for idx_file, excel_file in enumerate(uploaded_excels):
            with st.expander(f"📄 Archivo: **{excel_file.name}**", expanded=True):
                if excel_file.name.endswith('.csv'):
                    df_parsed = pd.read_csv(excel_file)
                    unit_name = excel_file.name.replace(".csv", "")
                else:
                    xls = pd.ExcelFile(excel_file)
                    hojas = xls.sheet_names
                    hoja_def = next((h for h in hojas if "total" in h.lower()), hojas[0])
                    
                    c_hoja, c_up = st.columns([2, 2])
                    with c_hoja:
                        hoja_sel = st.selectbox(
                            f"Hoja a procesar ({excel_file.name}):", 
                            hojas, 
                            index=hojas.index(hoja_def), 
                            key=f"sel_hoja_{idx_file}_{rc}"
                        )
                    
                    df_parsed, up_detectada = procesar_hoja_excel(xls, hoja_sel)
                    
                    with c_up:
                        label_up = "Nombre Unidad Productiva ⚠️ (Requerido)" if not up_detectada else "Nombre Unidad Productiva:"
                        unit_name = st.text_input(
                            label_up, 
                            value=up_detectada, 
                            key=f"up_input_{idx_file}_{rc}"
                        )
                        if not unit_name.strip():
                            st.markdown("<span style='color:red; font-size:12px;'>⚠️ Asigná un nombre para identificar esta Unidad Productiva.</span>", unsafe_allow_html=True)
                
                if df_parsed is not None and not df_parsed.empty:
                    final_name = unit_name.strip() if unit_name.strip() else f"Lote_{idx_file+1}"
                    dict_lotes[excel_file.name] = {
                        "df": df_parsed,
                        "unidad_productiva": final_name,
                        "file_name": excel_file.name
                    }
                    df_vis = df_parsed.copy().reset_index(drop=True)
                    df_vis.insert(0, "#", range(1, len(df_vis) + 1))
                    df_vis["Monto"] = df_vis["Monto"].apply(lambda x: f"$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    st.dataframe(df_vis, use_container_width=True, hide_index=True, column_config=config_columnas)
                else:
                    st.error("No se pudieron extraer datos válidos de este archivo.")

else: # Modo Manual / Pegar datos
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        pasted_text = st.text_area(
            "Pegá aquí las celdas copiadas de Excel (CUIT \t Nombre \t Monto):",
            height=150,
            placeholder="20-12345678-9\tJuan Perez\t15000",
            key=f"pasted_area_{rc}"
        )
        unit_manual = st.text_input("Nombre de la Unidad Productiva:", value="General", key=f"up_manual_{rc}")
    with col_m2:
        uploaded_pdf = st.file_uploader(
            "📄 Subí el archivo PDF con los comprobantes", 
            type=["pdf"], 
            key=f"pdf_uploader_manual_{rc}"
        )
    
    if pasted_text.strip():
        try:
            df_temp = pd.read_csv(io.StringIO(pasted_text.strip()), sep="\t", header=None)
            columnas_estandar = ["CUIT", "Nombre", "Monto"]
            num_cols = df_temp.shape[1]
            if num_cols <= 3:
                df_temp.columns = columnas_estandar[:num_cols]
            else:
                df_temp.columns = columnas_estandar + [f"Columna_{i+1}" for i in range(3, num_cols)]
            
            df_temp["CUIT"] = df_temp["CUIT"].astype(str).str.replace(".0", "", regex=False).str.strip()
            if "Monto" in df_temp.columns:
                df_temp["Monto"] = df_temp["Monto"].apply(parsear_monto)
            
            dict_lotes["manual_input"] = {
                "df": df_temp,
                "unidad_productiva": unit_manual.strip() if unit_manual.strip() else "General",
                "file_name": "Pegado Manual"
            }
            st.success("¡Datos manuales cargados correctamente!")
        except Exception as e:
            st.error(f"No se pudieron interpretar los datos pegados: {e}")

st.divider()

# --- SECCIÓN 2: PROCESAMIENTO MULTICAPA INTELIGENTE ---
if st.button("🚀 Procesar Comprobantes", type="primary"):
    if not dict_lotes:
        st.warning("⚠️ Primero cargá o pegá los datos del listado.")
    elif uploaded_pdf is None:
        st.warning("⚠️ Primero subí el archivo PDF con los comprobantes.")
    else:
        with st.spinner("Procesando lotes con el motor multicapa inteligente..."):
            try:
                pdf_bytes = uploaded_pdf.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                fecha_pago = extraer_fecha_pago_pdf(doc)
                
                textos_paginas = []
                for i_pag in range(len(doc)):
                    texto_pag = doc[i_pag].get_text("text")
                    textos_paginas.append((i_pag, normalizar_texto(texto_pag)))
                
                resultados_lotes = {}
                
                for key_file, info_lote in dict_lotes.items():
                    df_listado = info_lote["df"]
                    up_nombre = info_lote["unidad_productiva"]
                    
                    encontrados = []
                    no_encontrados = []
                    
                    for idx, row in df_listado.iterrows():
                        cuit_orig = str(row.get("CUIT", ""))
                        nombre_orig = str(row.get("Nombre", "Sin Nombre")).strip()
                        nombre_norm = normalizar_texto(nombre_orig)
                        monto_val = float(row.get("Monto", 0))
                        
                        monto_entero_str = str(int(monto_val)) if monto_val > 0 else ""
                        hallado = False
                        
                        # CAPA 1: Match Perfecto
                        for num_pag, texto_norm in textos_paginas:
                            texto_limpio = texto_norm.replace(".", "").replace(",", "")
                            montoencontrado = False
                            if monto_entero_str and len(monto_entero_str) >= 2:
                                if monto_entero_str in texto_limpio:
                                    montoencontrado = True
                            
                            nombre_encontrado = False
                            if nombre_norm:
                                palabras_nombre = [p for p in re.findall(r'\b\w+\b', nombre_norm) if len(p) > 2]
                                coincidencias = sum(1 for p in palabras_nombre if re.search(r'\b' + re.escape(p) + r'\b', texto_norm))
                                if palabras_nombre and (coincidencias >= len(palabras_nombre)):
                                    nombre_encontrado = True
                            else:
                                nombre_encontrado = True
                            
                            if montoencontrado and nombre_encontrado:
                                hallado = True
                                encontrados.append({
                                    "df_idx": idx,
                                    "CUIT": cuit_orig,
                                    "Nombre": nombre_orig,
                                    "Monto": f"$ {monto_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                    "Página PDF": num_pag + 1,
                                    "Pag_Index": num_pag
                                })
                                break
                        
                        # CAPA 2: Rescate Inteligente por Monto + Fuzzy
                        if not hallado and monto_entero_str:
                            for num_pag, texto_norm in textos_paginas:
                                texto_limpio = texto_norm.replace(".", "").replace(",", "")
                                if monto_entero_str in texto_limpio:
                                    titular_pdf, _ = extraer_datos_pagina(doc[num_pag].get_text("text"))
                                    titular_pdf_norm = normalizar_texto(titular_pdf)
                                    similitud = calcular_similitud(nombre_norm, titular_pdf_norm)
                                    if similitud >= 0.50 or any(p in titular_pdf_norm for p in re.findall(r'\b\w+\b', nombre_norm) if len(p) > 3):
                                        hallado = True
                                        encontrados.append({
                                            "df_idx": idx,
                                            "CUIT": cuit_orig,
                                            "Nombre": nombre_orig,
                                            "Monto": f"$ {monto_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                            "Página PDF": num_pag + 1,
                                            "Pag_Index": num_pag
                                        })
                                        break
                        
                        # CAPA 3: No Encontrados (Sugerencias)
                        if not hallado:
                            sugerencia_pag = None
                            titular_en_pdf = ""
                            monto_en_pdf = ""
                            
                            if nombre_norm:
                                palabras_nombre = [p for p in re.findall(r'\b\w+\b', nombre_norm) if len(p) > 2]
                                for num_pag, texto_norm in textos_paginas:
                                    coincidencias = sum(1 for p in palabras_nombre if re.search(r'\b' + re.escape(p) + r'\b', texto_norm))
                                    if palabras_nombre and (coincidencias / len(palabras_nombre)) >= 0.75:
                                        sugerencia_pag = num_pag + 1
                                        titular_en_pdf, monto_en_pdf = extraer_datos_pagina(doc[num_pag].get_text("text"))
                                        break
                            
                            no_encontrados.append({
                                "df_idx": idx,
                                "CUIT": cuit_orig,
                                "Nombre": nombre_orig,
                                "Monto": f"$ {monto_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                "Sugerencia_Pag": sugerencia_pag,
                                "Titular_PDF": titular_en_pdf,
                                "Monto_PDF": monto_en_pdf
                            })
                    
                    resultados_lotes[key_file] = {
                        "unidad_productiva": up_nombre,
                        "encontrados": encontrados,
                        "no_encontrados": no_encontrados
                    }
                
                st.session_state["doc_pdf"] = doc
                st.session_state["resultados_lotes"] = resultados_lotes
                st.session_state["fecha_pago"] = fecha_pago

            except Exception as e:
                st.error(f"Error al procesar: {e}")

# --- SECCIÓN 3: RESULTADOS, SUGERENCIAS Y ACCIONES DE EXCEPCIÓN ---
if "resultados_lotes" in st.session_state:
    st.subheader("📊 Resultados y Gestión de Excepciones por Unidad Productiva")
    
    doc_actual = st.session_state["doc_pdf"]
    fecha_str = st.session_state["fecha_pago"]
    
    titulos_tabs = []
    keys_lotes = list(st.session_state["resultados_lotes"].keys())
    
    for k in keys_lotes:
        res = st.session_state["resultados_lotes"][k]
        up = res["unidad_productiva"]
        tiene_conflictos = len(res["no_encontrados"]) > 0
        icono = "⚠️" if tiene_conflictos else "✅"
        titulos_tabs.append(f"{icono} {up}")
        
    tabs_lotes = st.tabs(titulos_tabs)
    
    archivos_pdf_generados = []
    
    for idx_tab, key_file in enumerate(keys_lotes):
        with tabs_lotes[idx_tab]:
            res_lote = st.session_state["resultados_lotes"][key_file]
            up_nombre = res_lote["unidad_productiva"]
            encs = res_lote["encontrados"]
            no_encs = res_lote["no_encontrados"]
            
            c1, c2 = st.columns(2)
            
            with c1:
                st.markdown(f"### 🟢 Encontrados ({len(encs)})")
                if encs:
                    df_e = pd.DataFrame(sorted(encs, key=lambda x: x["df_idx"])).drop(columns=["df_idx", "Pag_Index"], errors="ignore")
                    df_e.insert(0, "#", range(1, len(df_e) + 1))
                    st.dataframe(df_e, use_container_width=True, hide_index=True, column_config=config_columnas)
                else:
                    st.info("Sin coincidencias automáticas.")
            
            with c2:
                st.markdown(f"### 🔴 No encontrados / Conflictos ({len(no_encs)})")
                if no_encs:
                    st.write("Resolvé los casos pendientes aceptando sugerencias, forzando páginas o descartando:")
                    
                    for i_no, item in enumerate(no_encs):
                        df_idx = item["df_idx"]
                        nombre = item["Nombre"]
                        cuit = item["CUIT"]
                        monto = item["Monto"]
                        sug_pag = item["Sugerencia_Pag"]
                        titular_pdf = item["Titular_PDF"]
                        monto_pdf = item["Monto_PDF"]
                        
                        with st.container():
                            if sug_pag:
                                st.warning(f"⚠️ **{nombre}** — Posible coincidencia en Pág. {sug_pag}")
                                # gap="small" y proporciones equilibradas para alinear texto y botón perfectamente
                                col_txt_sug, col_btn_sug = st.columns([3.5, 2.5], gap="small")
                                with col_txt_sug:
                                    st.markdown(f"<div style='padding-top: 6px;'>💡 <b>Titular PDF:</b> {titular_pdf} &nbsp;|&nbsp; <b>Monto PDF:</b> $ {monto_pdf}</div>", unsafe_allow_html=True)
                                with col_btn_sug:
                                    if st.button(f"➕ Aceptar Sugerencia (Pág. {sug_pag})", key=f"btn_sug_{key_file}_{i_no}"):
                                        p_idx = sug_pag - 1
                                        st.session_state["resultados_lotes"][key_file]["encontrados"].append({
                                            "df_idx": df_idx,
                                            "CUIT": cuit,
                                            "Nombre": nombre,
                                            "Monto": monto,
                                            "Página PDF": sug_pag,
                                            "Pag_Index": p_idx
                                        })
                                        st.session_state["resultados_lotes"][key_file]["no_encontrados"].pop(i_no)
                                        st.rerun()
                            else:
                                st.error(f"❌ **{nombre}** — Sin coincidencias automáticas")
                            
                            # Uso de gap="small" para eliminar el espacio excesivo entre controles
                            col_input, col_btn_forzar, col_btn_omitir = st.columns([2, 2, 2], gap="small")
                            with col_input:
                                pag_manual = st.number_input(
                                    f"N° Pág manual para {nombre}", 
                                    min_value=1, 
                                    max_value=len(doc_actual), 
                                    value=sug_pag if sug_pag else 1, 
                                    key=f"input_pag_{key_file}_{i_no}",
                                    label_visibility="collapsed"
                                )
                            with col_btn_forzar:
                                if st.button("➕ Forzar Página", key=f"btn_forzar_{key_file}_{i_no}"):
                                    p_idx = int(pag_manual) - 1
                                    st.session_state["resultados_lotes"][key_file]["encontrados"].append({
                                        "df_idx": df_idx,
                                        "CUIT": cuit,
                                        "Nombre": nombre,
                                        "Monto": monto,
                                        "Página PDF": int(pag_manual),
                                        "Pag_Index": p_idx
                                    })
                                    st.session_state["resultados_lotes"][key_file]["no_encontrados"].pop(i_no)
                                    st.rerun()
                            with col_btn_omitir:
                                if st.button("🗑️ Descartar", key=f"btn_omitir_{key_file}_{i_no}"):
                                    st.session_state["resultados_lotes"][key_file]["no_encontrados"].pop(i_no)
                                    st.rerun()
                            st.divider()
                else:
                    st.success("¡100% de coincidencias confirmadas o descartadas!")
            
            # Generar PDF individual para descarga
            if encs:
                pdf_individual = fitz.open()
                encs_sorted = sorted(encs, key=lambda x: x["df_idx"])
                for item in encs_sorted:
                    p = item["Pag_Index"]
                    pdf_individual.insert_pdf(doc_actual, from_page=p, to_page=p)
                
                pdf_out_bytes = pdf_individual.write()
                nombre_pdf_salida = f"Comprobantes RI - {up_nombre} ({fecha_str}).pdf"
                
                archivos_pdf_generados.append({
                    "filename": nombre_pdf_salida,
                    "bytes": pdf_out_bytes
                })
                
                st.download_button(
                    label=f"📥 Descargar PDF para {up_nombre}",
                    data=pdf_out_bytes,
                    file_name=nombre_pdf_salida,
                    mime="application/pdf",
                    key=f"down_single_{key_file}"
                )

    # --- SECCIÓN FINAL: DESCARGA ZIP ---
    if archivos_pdf_generados:
        st.markdown("---")
        st.markdown("### 📦 Descarga Completa por Lote")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for item in archivos_pdf_generados:
                zip_file.writestr(item["filename"], item["bytes"])
        
        zip_buffer.seek(0)
        
        # Uso de gap="small" para mantener juntos el botón de ZIP y el de reiniciar
        col_zip, col_reset = st.columns([2, 1], gap="small")
        with col_zip:
            st.download_button(
                label="📦 Descargar TODOS los PDFs en un .ZIP",
                data=zip_buffer,
                file_name=f"Comprobantes_RI_Lote_({fecha_str}).zip",
                mime="application/zip",
                type="primary"
            )
        with col_reset:
            if st.button("🔄 Reiniciar Todo", type="secondary"):
                st.session_state["reset_count"] += 1
                for k in list(st.session_state.keys()):
                    if k != "reset_count":
                        del st.session_state[k]
                st.rerun()