import streamlit as st
import pandas as pd
import os
import re
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Generador F30-1 por Faena", layout="wide")

st.title("📂 Procesador de Archivos Previred F30-1")
st.markdown("Carga la nómina y el CSV de Previred para segmentar por obra.")

# 🧹 Funciones de utilidad (Mantenemos tu lógica original)
def clean_filename(text):
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text.replace('\n', ' ').strip().replace(" ", "_")

def obtener_periodo_anterior(periodo_actual):
    meses = [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    ]
    try:
        nombre_mes, año = periodo_actual.strip().split()
        año = int(año)
        indice = meses.index(nombre_mes.capitalize())
        if indice == 0:
            mes_anterior = meses[-1]
            año_anterior = año - 1
        else:
            mes_anterior = meses[indice - 1]
            año_anterior = año
        return f"{mes_anterior} {año_anterior}"
    except:
        return None

# --- INTERFAZ DE STREAMLIT ---

col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("1. Seleccionar nómina (Excel)", type=["xlsx"])

with col2:
    csv_file = st.file_uploader("2. Seleccionar CSV Previred", type=["csv"])

if excel_file and csv_file:
    # Leer Excel para obtener períodos
    df_temp = pd.read_excel(excel_file)
    periodos_disponibles = sorted(df_temp['periodo'].dropna().unique().tolist())
    
    periodo_a_generar = st.selectbox("3. Seleccione el período a procesar", periodos_disponibles)
    
    if st.button("🚀 Generar Archivos"):
        periodo_para_buscar = obtener_periodo_anterior(periodo_a_generar)
        
        # Procesamiento
        df_excel = df_temp[df_temp['periodo'] == periodo_para_buscar].copy()
        df_excel['rut_limpio'] = df_excel['rut_trabajador'].str.replace(r'[^0-9]', '', regex=True).str[:-1]
        
        # Leer CSV (Nota: Streamlit maneja los archivos en memoria)
        df_csv = pd.read_csv(csv_file, encoding='latin1', sep=';', header=None)
        df_csv.columns = [f"col_{i}" for i in range(df_csv.shape[1])]
        df_csv['rut_limpio'] = df_csv['col_0'].astype(str).str.replace(r'\D', '', regex=True)
        
        log_registros = []
        archivos_output = []

        # Lógica de filtrado por obra
        for obra, sub_df in df_excel.groupby("obra_faena_servicio"):
            ruts_limpios = sub_df['rut_limpio'].tolist()
            df_filtrado = df_csv[df_csv['rut_limpio'].isin(ruts_limpios)]

            if not df_filtrado.empty:
                nombre_archivo = f"{clean_filename(obra)} - {clean_filename(periodo_a_generar)}.csv"
                # Guardar en buffer para descarga
                csv_buffer = df_filtrado.drop(columns=['rut_limpio']).to_csv(sep=';', index=False, header=False).encode('latin1')
                archivos_output.append((nombre_archivo, csv_buffer))
                
                for rut in df_filtrado['rut_limpio']:
                    log_registros.append({
                        "obra_faena_servicio": obra,
                        "rut": rut,
                        "archivo": nombre_archivo
                    })

        if archivos_output:
            st.success(f"✅ Se han generado {len(archivos_output)} archivos.")
            
            # Mostrar resultados y permitir descarga
            for nombre, contenido in archivos_output:
                st.download_button(
                    label=f"⬇️ Descargar {nombre}",
                    data=contenido,
                    file_name=nombre,
                    mime="text/csv"
                )
        else:
            st.warning("No se encontraron coincidencias para los RUTs en el período seleccionado.")
