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

# Si los archivos están cargados, hacer el análisis completo
if crossref_file and mb52_file and coois_file and zco41_file:
    st.success("Archivos cargados correctamente")

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

    # Mostrar resumen de materiales faltantes
    faltantes = full[full['Available after ALL'] < 0].copy()
    faltantes['Cantidad Faltante'] = -faltantes['Available after ALL']
    faltantes_resumen = faltantes[['Custom Description', 'Cantidad Faltante']].sort_values(by='Cantidad Faltante', ascending=False)

    st.subheader("📦 Material que hace falta para poder producir")
    st.dataframe(faltantes_resumen, use_container_width=True)

    # Materiales no encontrados en CrossReference
    missing_in_crossref_coois = coois[~coois['Custom Description'].isin(crossref['Custom Description'])]['Custom Description'].drop_duplicates()
    missing_in_crossref_zco41 = zco41[~zco41['Custom Description'].isin(crossref['Custom Description'])]['Custom Description'].drop_duplicates()

    if not missing_in_crossref_coois.empty:
        st.warning("🚫 Descripciones en COOIS que no están en el CrossReference:")
        st.dataframe(missing_in_crossref_coois, use_container_width=True)

    if not missing_in_crossref_zco41.empty:
        st.warning("🚫 Descripciones en ZCO41 que no están en el CrossReference:")
        st.dataframe(missing_in_crossref_zco41, use_container_width=True)

    # --- Sección de análisis ---
    from datetime import date
    today = pd.to_datetime(date.today())

    # COOIS producibles y no producibles
    coois_stock = coois.merge(full[['Custom Description', 'Open Quantity']], on='Custom Description', how='left')
    coois_stock['Enough'] = coois_stock['Order quantity (GMEIN)'] <= coois_stock['Open Quantity']
    coois_stock['Shortage'] = coois_stock['Order quantity (GMEIN)'] - coois_stock['Open Quantity']
    coois_explain = coois_stock[~coois_stock['Enough'] & (coois_stock['Shortage'] > 0)][['Sales Order', 'Custom Description', 'Order quantity (GMEIN)', 'Open Quantity', 'Shortage']]

    coois_group = coois_stock.groupby('Sales Order')['Enough'].all().reset_index()
    coois_group['Producible'] = coois_group['Enough']
    coois_final = coois_stock.merge(coois_group[['Sales Order', 'Producible']], on='Sales Order', how='left')

    # ZCO41 producibles y no producibles
    zco41_stock = zco41.merge(full[['Custom Description', 'Available after COOIS']], on='Custom Description', how='left')
    zco41_stock['Enough'] = zco41_stock['Pln.Or Qty'] <= zco41_stock['Available after COOIS']
    zco41_stock['Shortage'] = zco41_stock['Pln.Or Qty'] - zco41_stock['Available after COOIS']
    zco41_explain = zco41_stock[~zco41_stock['Enough'] & (zco41_stock['Shortage'] > 0)][['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Available after COOIS', 'Shortage']]

    zco41_group = zco41_stock.groupby('Sales Order')['Enough'].all().reset_index()
    zco41_group['Producible'] = zco41_group['Enough']
    zco41_final = zco41_stock.merge(zco41_group[['Sales Order', 'Producible']], on='Sales Order', how='left')

    # Mostrar resultados
    st.subheader("✅ Órdenes de ZCO41 que SÍ se pueden producir")
    st.dataframe(zco41_final[zco41_final['Producible']], use_container_width=True)

    st.subheader("❌ Órdenes de ZCO41 que NO se pueden producir")
    st.dataframe(zco41_final[~zco41_final['Producible']], use_container_width=True)

    st.subheader("❌ Órdenes de COOIS que NO se pueden producir")
    st.dataframe(coois_final[~coois_final['Producible']], use_container_width=True)

    # Explicación por qué no se puede producir
    st.subheader("🔍 Detalle de faltantes en COOIS")
    st.dataframe(coois_explain, use_container_width=True)

    st.subheader("🔍 Detalle de faltantes en ZCO41")
    st.dataframe(zco41_explain, use_container_width=True)

    # Past Due Orders (ZCO41 y COOIS)
    coois_past_due = coois_final[(~coois_final['Producible']) & (pd.to_datetime(coois_final['Est. Ship Date']) < today)]
    zco41_past_due = zco41_final[(~zco41_final['Producible']) & (pd.to_datetime(zco41_final['Estimated Ship Date']) < today)]

    st.subheader("⏰ Órdenes Past Due que NO se pueden producir (ZCO41)")
    st.dataframe(zco41_past_due, use_container_width=True)

    st.subheader("⏰ Órdenes Past Due que NO se pueden producir (COOIS)")
    st.dataframe(coois_past_due, use_container_width=True)

else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
