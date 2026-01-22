import streamlit as st
import pandas as pd
import re
import io
import zipfile

# Configuración de la página
st.set_page_config(page_title="Generador F30-1 por Faena", page_icon="📑", layout="wide")

# --- FUNCIONES DE LÓGICA Y UTILIDAD ---

def clean_filename(text):
    """Limpia nombres de texto para que sean compatibles con archivos del sistema."""
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text.replace('\n', ' ').strip().replace(" ", "_")

def obtener_periodo_anterior(periodo_actual):
    """Calcula el mes anterior para buscar la asignación de faenas en el Excel."""
    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
             "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    try:
        nombre_mes, año = periodo_actual.strip().split()
        año = int(año)
        indice = meses.index(nombre_mes.capitalize())
        if indice == 0:
            mes_anterior, año_anterior = meses[-1], año - 1
        else:
            mes_anterior, año_anterior = meses[indice - 1], año
        return f"{mes_anterior} {año_anterior}"
    except:
        return None

def generar_template_excel():
    """Genera un archivo Excel de ejemplo para que el usuario use como base."""
    output = io.BytesIO()
    df_template = pd.DataFrame({
        "periodo": ["Enero 2026", "Enero 2026"],
        "rut_trabajador": ["12.345.678-9", "11.222.333-4"],
        "nombre_trabajador": ["EJEMPLO JUAN PEREZ", "EJEMPLO MARIA SOTO"],
        "obra_faena_servicio": ["NOMBRE FAENA A", "NOMBRE FAENA B"]
    })
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Nomina')
    return output.getvalue()

# --- BARRA LATERAL (SIDEBAR) ---

with st.sidebar:
    st.header("🛠️ Recursos y Configuración")
    st.markdown("Descarga el formato base si no lo tienes:")
    
    template_data = generar_template_excel()
    st.download_button(
        label="📥 Descargar Template Excel",
        data=template_data,
        file_name="template_nomina_f30-1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.divider()
    st.info("""
    **Instrucciones:**
    1. Sube el Excel con la nómina histórica.
    2. Sube el CSV masivo de Previred.
    3. Selecciona el mes que vas a declarar. El sistema buscará las faenas en el mes anterior.
    """)

# --- CUERPO PRINCIPAL ---

st.title("📂 Procesador de Archivos Previred F30-1")
st.markdown("Herramienta para segmentar el CSV de Previred según la obra/faena asignada en la nómina.")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. Cargar Excel (Nómina)", type=["xlsx"])
with col2:
    csv_file = st.file_uploader("2. Cargar CSV (Previred)", type=["csv"])

if excel_file and csv_file:
    # Leer Excel y detectar periodos
    try:
        df_temp = pd.read_excel(excel_file)
        if 'periodo' not in df_temp.columns:
            st.error("El Excel debe tener una columna llamada 'periodo'")
        else:
            periodos_disponibles = sorted(df_temp['periodo'].dropna().unique().tolist())
            periodo_a_generar = st.selectbox("3. Seleccione el período a procesar", periodos_disponibles)
            
            if st.button("🚀 Iniciar Procesamiento"):
                periodo_para_buscar = obtener_periodo_anterior(periodo_a_generar)
                
                # Filtrar Excel por periodo anterior
                df_excel = df_temp[df_temp['periodo'] == periodo_para_buscar].copy()
                
                if df_excel.empty:
                    st.warning(f"No se encontró información para el periodo anterior: {periodo_para_buscar}")
                else:
                    # Limpieza de RUT en Excel
                    df_excel['rut_limpio'] = df_excel['rut_trabajador'].astype(str).str.replace(r'[^0-9]', '', regex=True).str[:-1]
                    
                    # Leer CSV Previred
                    df_csv = pd.read_csv(csv_file, encoding='latin1', sep=';', header=None)
                    df_csv.columns = [f"col_{i}" for i in range(df_csv.shape[1])]
                    # Limpieza de RUT en CSV (columna 0)
                    df_csv['rut_limpio'] = df_csv['col_0'].astype(str).str.replace(r'\D', '', regex=True)
                    
                    archivos_output = []
                    log_data = []

                    # Agrupar por obra y filtrar
                    for obra, sub_df in df_excel.groupby("obra_faena_servicio"):
                        ruts_limpios = sub_df['rut_limpio'].tolist()
                        df_filtrado = df_csv[df_csv['rut_limpio'].isin(ruts_limpios)]

                        if not df_filtrado.empty:
                            nombre_csv = f"{clean_filename(obra)} - {clean_filename(periodo_a_generar)}.csv"
                            content = df_filtrado.drop(columns=['rut_limpio']).to_csv(sep=';', index=False, header=False).encode('latin1')
                            archivos_output.append((nombre_csv, content))
                            
                            for rut in df_filtrado['rut_limpio']:
                                nombre_trab = sub_df[sub_df['rut_limpio'] == str(rut)]['nombre_trabajador'].values
                                nombre_trab = nombre_trab[0] if len(nombre_trab) > 0 else "N/A"
                                log_data.append({
                                    "Obra": obra, 
                                    "RUT": rut, 
                                    "Trabajador": nombre_trab,
                                    "Archivo": nombre_csv
                                })

                    if archivos_output:
                        st.success(f"✅ ¡Éxito! Se generaron {len(archivos_output)} archivos de faena.")
                        
                        # Crear ZIP
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                            for name, data in archivos_output:
                                zip_file.writestr(name, data)
                            
                            # Log de proceso
                            df_log = pd.DataFrame(log_data)
                            zip_file.writestr(f"log_proceso_{clean_filename(periodo_a_generar)}.csv", df_log.to_csv(index=False).encode('utf-8'))

                        st.download_button(
                            label="🎁 Descargar todos los archivos (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Archivos_F30-1_{clean_filename(periodo_a_generar)}.zip",
                            mime="application/zip"
                        )
                        
                        with st.expander("Ver detalle de trabajadores encontrados"):
                            st.dataframe(df_log)
                    else:
                        st.error("No hubo coincidencias de RUT entre el Excel (periodo anterior) y el CSV de Previred.")
    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")
else:
    st.info("Por favor, carga ambos archivos para comenzar.")
