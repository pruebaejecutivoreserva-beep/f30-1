import streamlit as st
import pandas as pd
import re
import io
import zipfile
import os

# 1. Configuración de la página
st.set_page_config(page_title="Generador F30-1 | APS", page_icon="📑", layout="wide")

# 2. Estilos Personalizados (CSS)
st.markdown("""
    <style>
        /* Unificación de fondos: Azul APS */
        .stApp, [data-testid="stSidebar"], [data-testid="stHeader"] {
            background-color: #003366 !important;
        }

        /* Texto siempre blanco */
        .stMarkdown, p, span, label, h1, h2, h3, .stWidgetLabel {
            color: white !important;
        }

        /* Botones: Siempre azules con borde blanco para resaltar */
        .stButton>button {
            background-color: #003366 !important;
            color: white !important;
            border: 2px solid white !important;
            border-radius: 5px !important;
            font-weight: bold !important;
            transition: none !important;
        }

        /* Evitar que el botón cambie de color al pasar el mouse o hacer clic */
        .stButton>button:hover, .stButton>button:active, .stButton>button:focus {
            background-color: #003366 !important;
            color: white !important;
            border: 2px solid white !important;
        }

        /* Cuadros de carga de archivos */
        div[data-testid="stFileUploader"] section {
            background-color: rgba(255, 255, 255, 0.05) !important;
            border: 1px dashed white !important;
        }
        
        /* Texto interno de los cargadores */
        div[data-testid="stFileUploader"] section div div {
            color: white !important;
        }

        /* Selectbox (Fondo blanco para lectura de opciones) */
        div[data-baseweb="select"] > div {
            background-color: white !important;
            color: black !important;
        }
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIONES DE LÓGICA ---
def clean_filename(text):
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text.replace('\n', ' ').strip().replace(" ", "_")

def obtener_periodo_anterior(periodo_actual):
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

# --- SIDEBAR ---
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.header("🛠️ Configuración")
    st.markdown("Descarga el formato base:")
    st.download_button(
        label="📥 Descargar Template Excel",
        data=generar_template_excel(),
        file_name="template_nomina_f30-1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.divider()
    st.info("El sistema busca faenas del mes anterior al seleccionado.")

# --- CUERPO PRINCIPAL ---
# Logo también en la pantalla principal
if os.path.exists("logo.png"):
    col_logo, _ = st.columns([1, 4])
    with col_logo:
        st.image("logo.png", width=200)

st.title("📂 Procesador de Archivos Previred F30-1")
st.write("Carga los archivos para iniciar la segmentación por obra.")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. Cargar Excel (Nómina)", type=["xlsx"])
with col2:
    csv_file = st.file_uploader("2. Cargar CSV (Previred)", type=["csv"])

if excel_file and csv_file:
    try:
        df_temp = pd.read_excel(excel_file)
        periodos_disponibles = sorted(df_temp['periodo'].dropna().unique().tolist())
        periodo_a_generar = st.selectbox("3. Seleccione el período a declarar", periodos_disponibles)
        
        if st.button("🚀 Iniciar Procesamiento"):
            periodo_para_buscar = obtener_periodo_anterior(periodo_a_generar)
            df_excel = df_temp[df_temp['periodo'] == periodo_para_buscar].copy()
            
            if df_excel.empty:
                st.warning(f"No hay datos registrados para {periodo_para_buscar}")
            else:
                # Limpieza RUT
                df_excel['rut_limpio'] = df_excel['rut_trabajador'].astype(str).str.replace(r'[^0-9]', '', regex=True).str[:-1]
                df_csv = pd.read_csv(csv_file, encoding='latin1', sep=';', header=None)
                df_csv['rut_limpio'] = df_csv[0].astype(str).str.replace(r'\D', '', regex=True)
                
                archivos_output = []
                log_data = []

                for obra, sub_df in df_excel.groupby("obra_faena_servicio"):
                    ruts = sub_df['rut_limpio'].tolist()
                    df_filtrado = df_csv[df_csv['rut_limpio'].isin(ruts)]

                    if not df_filtrado.empty:
                        nombre_csv = f"{clean_filename(obra)} - {clean_filename(periodo_a_generar)}.csv"
                        csv_bytes = df_filtrado.drop(columns=['rut_limpio']).to_csv(sep=';', index=False, header=False).encode('latin1')
                        archivos_output.append((nombre_csv, csv_bytes))
                        
                        for r in df_filtrado['rut_limpio']:
                            log_data.append({"Obra": obra, "RUT": r, "Estado": "Encontrado"})

                if archivos_output:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zip_f:
                        for n, d in archivos_output:
                            zip_f.writestr(n, d)
                        df_log = pd.DataFrame(log_data)
                        zip_f.writestr(f"log_{clean_filename(periodo_a_generar)}.csv", df_log.to_csv(index=False).encode('utf-8'))

                    st.success(f"✅ Se han generado {len(archivos_output)} archivos.")
                    st.download_button(
                        label="🎁 Descargar todo en un ZIP",
                        data=zip_buf.getvalue(),
                        file_name=f"F301_{clean_filename(periodo_a_generar)}.zip",
                        mime="application/zip"
                    )
                    with st.expander("Ver detalle de registros"):
                        st.dataframe(df_log)
                else:
                    st.error("Sin coincidencias entre Excel y CSV.")
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Esperando archivos...")
