import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Ordenes Producibles", layout="wide")
st.title("📄 Análisis de Producción: COOIS, ZCO41 vs MB52")

# Subida de archivos
st.sidebar.header(":inbox_tray: Subir Archivos")
crossref_file = st.sidebar.file_uploader("Tabla de equivalencias (Custom vs Non Custom)", type="xlsx")
mb52_file = st.sidebar.file_uploader("MB52 - Inventario", type="xlsx")
coois_file = st.sidebar.file_uploader("COOIS - Órdenes fijas", type="xlsx")
zco41_file = st.sidebar.file_uploader("ZCO41 - Nueva demanda", type="xlsx")

# Sección manual para inputs diarios
st.sidebar.header("🛠️ Input Diario Manual")
lasers_available = st.sidebar.number_input("Lasers disponibles", min_value=0, value=40)
lasers_running = st.sidebar.number_input("Lasers corriendo", min_value=0, value=33)
forecast = st.sidebar.number_input("Forecast semanal", min_value=0, value=118000)
units_shipped = st.sidebar.number_input("Units shipped so far", min_value=0, value=102922)
balance = forecast - units_shipped

# Mostrar resumen diario en el panel principal
with st.expander("📊 Resumen Diario Manual"):
    resumen_diario = pd.DataFrame({
        'Concepto': [
            'Lasers disponibles', 'Lasers corriendo',
            'Forecast semanal', 'Units shipped so far', 'Balance'
        ],
        'Valor': [
            lasers_available, lasers_running,
            forecast, units_shipped, balance
        ]
    })
    st.dataframe(resumen_diario, use_container_width=True)

    # Si los archivos están cargados, hacer el análisis básico y agregarlo al Excel
    if crossref_file and mb52_file and coois_file and zco41_file:
        crossref = pd.read_excel(crossref_file)
        mb52 = pd.read_excel(mb52_file)
        coois = pd.read_excel(coois_file)
        zco41 = pd.read_excel(zco41_file)

        crossref = crossref.rename(columns={"Non Custom": "Material description", "Custom": "Custom Description"})
        mb52_grouped = mb52.groupby('Material description', as_index=False)['Open Quantity'].sum()
        mb52_custom = mb52_grouped.merge(crossref, on='Material description', how='left')

        coois = coois.rename(columns={'Material description': 'Custom Description'})
        zco41 = zco41.rename(columns={'Material description': 'Custom Description'})

        coois_sum = coois.groupby('Custom Description', as_index=False)['Order quantity (GMEIN)'].sum()
        zco41_sum = zco41.groupby('Custom Description', as_index=False)['Pln.Or Qty'].sum()

        full = mb52_custom[['Custom Description', 'Open Quantity']].merge(coois_sum, on='Custom Description', how='left')\
            .merge(zco41_sum, on='Custom Description', how='left').fillna(0)

        full['Available after COOIS'] = full['Open Quantity'] - full['Order quantity (GMEIN)']
        full['Available after ALL'] = full['Available after COOIS'] - full['Pln.Or Qty']

        faltantes = full[full['Available after ALL'] < 0].copy()
        faltantes['Cantidad Faltante'] = -faltantes['Available after ALL']
        faltantes = faltantes[['Custom Description', 'Cantidad Faltante']]
        faltantes = faltantes.sort_values(by='Cantidad Faltante', ascending=False)

        # Combinar resumen diario y faltantes en un solo Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            resumen_diario.to_excel(writer, sheet_name='Resumen Diario', index=False)
            faltantes.to_excel(writer, sheet_name='Materiales Faltantes', index=False)
        output.seek(0)

        st.download_button(
            label="📥 Descargar resumen en Excel",
            data=output,
            file_name=f"resumen_diario_{datetime.today().date()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Continuación lógica solo si hay archivos cargados
if crossref_file and mb52_file and coois_file and zco41_file:
    st.success("Archivos cargados correctamente")
else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
