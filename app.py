import streamlit as st
import pandas as pd
import io
import fitz  # PyMuPDF
import unicodedata
import re
from difflib import SequenceMatcher

st.set_page_config(page_title="ProgramaDeFiacosos", page_icon="💸", layout="wide")

st.title("Ayuda vagos 3000 😎")
st.subheader("Extractor automático de comprobantes de transferencia")

# Control de reseteo para limpiar componentes de archivos
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

# --- SECCIÓN 1: CARGAR EL LISTADO DE PERSONAS (INTACTO Y SEGURO) ---
st.markdown("### 1. Cargar el listado de personas / transferencias")

tab1, tab2 = st.tabs(["📋 Pegar datos desde Excel", "📁 Subir archivo Excel/CSV"])

df_listado = None
rc = st.session_state["reset_count"]

with tab1:
    st.write("Pegá los datos directo desde tu Excel (con o sin encabezados):")
    pasted_text = st.text_area(
        "Pegá aquí las celdas copiadas de Excel:",
        height=150,
        placeholder="20-12345678-9\tJuan Perez\t15000\n27-98765432-1\tMaria Gomez\t25000",
        key=f"pasted_area_{rc}"
    )
    if pasted_text.strip():
        try:
            df_temp = pd.read_csv(io.StringIO(pasted_text.strip()), sep="\t", header=None)
            primera_fila = " ".join([str(val).lower() for val in df_temp.iloc[0].values])
            if any(k in primera_fila for k in ["cuit", "nombre", "monto", "apellido"]):
                df_temp = df_temp.iloc[1:].reset_index(drop=True)
            
            columnas_estandar = ["CUIT", "Nombre", "Monto"]
            num_cols = df_temp.shape[1]
            if num_cols <= 3:
                df_temp.columns = columnas_estandar[:num_cols]
            else:
                df_temp.columns = columnas_estandar + [f"Columna_{i+1}" for i in range(3, num_cols)]
            
            df_temp["CUIT"] = df_temp["CUIT"].astype(str).str.replace(".0", "", regex=False).str.strip()
            df_temp = df_temp[~df_temp["CUIT"].str.upper().str.contains("TOTAL|SUMA|SUBTOTAL", na=False)]
            
            if "Monto" in df_temp.columns:
                df_temp["Monto"] = df_temp["Monto"].apply(parsear_monto)
            
            df_listado = df_temp
            st.success("¡Datos pegados procesados correctamente!")
        except Exception as e:
            st.error(f"No se pudieron interpretar los datos pegados: {e}")

with tab2:
    uploaded_excel = st.file_uploader("Elegí el archivo Excel o CSV", type=["xlsx", "xls", "csv"], key=f"excel_uploader_{rc}")
    
    if uploaded_excel is not None:
        try:
            if uploaded_excel.name.endswith('.csv'):
                df_listado = pd.read_csv(uploaded_excel)
                st.success("¡CSV cargado correctamente!")
            else:
                xls = pd.ExcelFile(uploaded_excel)
                hojas_disponibles = xls.sheet_names
                hoja_sugerida = next((h for h in hojas_disponibles if "total" in h.lower()), hojas_disponibles[0])
                hoja_seleccionada = st.selectbox("📌 Seleccioná la hoja que querés procesar:", hojas_disponibles, index=hojas_disponibles.index(hoja_sugerida), key=f"sheet_sel_{rc}")
                
                if hoja_seleccionada:
                    df_raw = pd.read_excel(xls, sheet_name=hoja_seleccionada, header=None)
                    fila_encabezado = None
                    for idx, row in df_raw.iterrows():
                        texto_fila = " ".join([str(val).lower() for val in row.values])
                        if "cuit" in texto_fila and ("nombre" in texto_fila or "apellido" in texto_fila):
                            fila_encabezado = idx
                            break
                    
                    if fila_encabezado is not None:
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
                        if cols_a_conservar:
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
                                st.info("ℹ️ Se detectaron registros repetidos. Agrupando y sumando montos por CUIT...")
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
                            df_listado = df_filtrado
                            st.success(f"¡Hoja '{hoja_seleccionada}' procesada exitosamente!")
        except Exception as e:
            st.error(f"Error al leer el archivo Excel: {e}")

# CONFIGURACIÓN DE COLUMNAS CON ANCHOS ESPECÍFICOS
config_columnas = {
    "#": st.column_config.Column("#", width=40),
    "CUIT": st.column_config.TextColumn("CUIT", width=150),
    "Nombre": st.column_config.TextColumn("Nombre", width=300),
    "Monto": st.column_config.TextColumn("Monto", width=150),
    "Página PDF": st.column_config.NumberColumn("Página PDF", width=100, format="%d")
}

if df_listado is not None:
    df_mostrar = df_listado.reset_index(drop=True).copy()
    df_mostrar.insert(0, "#", range(1, len(df_mostrar) + 1))
    if "Monto" in df_mostrar.columns:
        df_mostrar["Monto"] = df_mostrar["Monto"].apply(lambda x: f"$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    st.write(f"📋 **Vista previa del listado limpio a buscar** (Total registros: {len(df_mostrar)}):")
    st.dataframe(
        df_mostrar, 
        use_container_width=True, 
        hide_index=True,
        column_config=config_columnas
    )

st.divider()

st.markdown("### 2. Subí el PDF con los comprobantes mezclados")
uploaded_pdf = st.file_uploader("Elegí el archivo PDF", type=["pdf"], key=f"pdf_uploader_{rc}")

if uploaded_pdf is not None:
    st.success("¡PDF de comprobantes cargado con éxito!")

st.divider()

# --- SECCIÓN 3: PROCESAMIENTO HÍBRIDO MULTICAPA (OPCIÓN 3) ---
if st.button("🚀 Procesar Comprobantes", type="primary"):
    if df_listado is None:
        st.warning("⚠️ Primero cargá o pegá el listado de personas a buscar.")
    elif uploaded_pdf is None:
        st.warning("⚠️ Primero subí el archivo PDF con los comprobantes.")
    else:
        with st.spinner("Analizando comprobantes con sistema multicapa inteligente..."):
            try:
                pdf_bytes = uploaded_pdf.read()
                doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                
                encontrados = []
                no_encontrados = []
                
                textos_paginas = []
                for i_pag in range(len(doc)):
                    texto_pag = doc[i_pag].get_text("text")
                    textos_paginas.append((i_pag, normalizar_texto(texto_pag)))
                
                for idx, row in df_listado.iterrows():
                    cuit_orig = str(row.get("CUIT", ""))
                    nombre_orig = str(row.get("Nombre", "Sin Nombre")).strip()
                    nombre_norm = normalizar_texto(nombre_orig)
                    monto_val = float(row.get("Monto", 0))
                    
                    monto_entero_str = str(int(monto_val)) if monto_val > 0 else ""
                    hallado = False
                    
                    # --- CAPA 1: MATCH PERFECTO (Nombre exacto + Monto) ---
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
                    
                    # --- CAPA 2: RESCATE INTELIGENTE (Anclaje por Monto Exacto + Similitud Tipográfica / Fuzzy) ---
                    if not hallado and monto_entero_str:
                        for num_pag, texto_norm in textos_paginas:
                            texto_limpio = texto_norm.replace(".", "").replace(",", "")
                            if monto_entero_str in texto_limpio:
                                # El monto coincide exactamente en esta página. Extraemos el titular del PDF para comparar.
                                titular_pdf, _ = extraer_datos_pagina(doc[num_pag].get_text("text"))
                                titular_pdf_norm = normalizar_texto(titular_pdf)
                                
                                # Comparamos similitud de texto (umbral del 50% para absorber FERRIRA/Ferreira, LEGIZAMON/Leguizamon, etc.)
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
                    
                    # --- CAPA 3: NO ENCONTRADO (Generar sugerencias o revisión manual) ---
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
                
                st.session_state["doc_pdf"] = doc
                st.session_state["encontrados"] = encontrados
                st.session_state["no_encontrados"] = no_encontrados

            except Exception as e:
                st.error(f"Error al procesar el PDF: {e}")

# --- MOSTRAR RESULTADOS Y GESTIÓN DE EXCEPCIONES ---
if "encontrados" in st.session_state:
    st.subheader("📊 Resultados del Procesamiento")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"### 🟢 Encontrados ({len(st.session_state['encontrados'])})")
        if st.session_state["encontrados"]:
            df_enc_sorted = sorted(st.session_state["encontrados"], key=lambda x: x["df_idx"])
            df_enc = pd.DataFrame(df_enc_sorted).drop(columns=["df_idx", "Pag_Index"], errors="ignore")
            df_enc.insert(0, "#", range(1, len(df_enc) + 1))
            st.dataframe(
                df_enc, 
                use_container_width=True, 
                hide_index=True,
                column_config=config_columnas
            )
        else:
            st.info("No se encontró ningún comprobante automático.")
    
    with col2:
        st.markdown(f"### 🔴 No encontrados ({len(st.session_state['no_encontrados'])})")
        if st.session_state["no_encontrados"]:
            st.write("Casos pendientes o con diferencias de monto:")
            
            for idx, item in enumerate(st.session_state["no_encontrados"]):
                df_idx = item["df_idx"]
                nombre = item["Nombre"]
                cuit = item["CUIT"]
                monto = item["Monto"]
                sug_pag = item["Sugerencia_Pag"]
                titular_pdf = item["Titular_PDF"]
                monto_pdf = item["Monto_PDF"]
                
                with st.container():
                    if sug_pag:
                        st.warning(f"⚠️ **{nombre}** — Posibles coincidencias encontradas")
                        
                        cols_sug = st.columns([3, 2])
                        with cols_sug[0]:
                            st.caption(f"💡 Sugerencia: Encontrado en **Página {sug_pag}**<br>• Titular: **{titular_pdf}** | Monto PDF: **$ {monto_pdf}**", unsafe_allow_html=True)
                        with cols_sug[1]:
                            if st.button(f"➕ Aceptar Sugerencia (Pág. {sug_pag})", key=f"btn_sug_{idx}"):
                                p_idx = sug_pag - 1
                                st.session_state["encontrados"].append({
                                    "df_idx": df_idx,
                                    "CUIT": cuit,
                                    "Nombre": nombre,
                                    "Monto": monto,
                                    "Página PDF": sug_pag,
                                    "Pag_Index": p_idx
                                })
                                st.session_state["no_encontrados"].pop(idx)
                                st.rerun()
                    else:
                        st.error(f"❌ **{nombre}** — Sin coincidencias")
                    
                    col_input, col_btn = st.columns([3, 2])
                    with col_input:
                        pag_manual = st.number_input(
                            f"N° de Página manual para {nombre}", 
                            min_value=1, 
                            max_value=len(st.session_state["doc_pdf"]), 
                            value=sug_pag if sug_pag else 1, 
                            key=f"input_pag_{idx}",
                            label_visibility="collapsed"
                        )
                    with col_btn:
                        if st.button("➕ Forzar Página", key=f"btn_forzar_{idx}"):
                            p_idx = int(pag_manual) - 1
                            st.session_state["encontrados"].append({
                                "df_idx": df_idx,
                                "CUIT": cuit,
                                "Nombre": nombre,
                                "Monto": monto,
                                "Página PDF": int(pag_manual),
                                "Pag_Index": p_idx
                            })
                            st.session_state["no_encontrados"].pop(idx)
                            st.rerun()
                    st.divider()
        else:
            st.success("¡Todos los registros fueron incluidos!")
    
    # --- SECCIÓN FINAL: DESCARGA Y REINICIO ---
    if st.session_state["encontrados"]:
        doc_actual = st.session_state["doc_pdf"]
        pdf_final = fitz.open()
        
        encontrados_ordenados = sorted(st.session_state["encontrados"], key=lambda x: x["df_idx"])
        for item in encontrados_ordenados:
            p = item["Pag_Index"]
            pdf_final.insert_pdf(doc_actual, from_page=p, to_page=p)
        
        pdf_bytes = pdf_final.write()
        st.markdown("---")
        
        col_down, col_res, col_space = st.columns([2, 1, 3])
        with col_down:
            st.download_button(
                label="📥 Descargar PDF con Comprobantes Filtrados",
                data=pdf_bytes,
                file_name="Comprobantes_Filtrados_ProgramaDeFiacosos.pdf",
                mime="application/pdf",
                type="primary"
            )
        with col_res:
            if st.button("🔄 Reiniciar", type="secondary"):
                st.session_state["reset_count"] += 1
                for key in list(st.session_state.keys()):
                    if key != "reset_count":
                        del st.session_state[key]
                st.rerun()