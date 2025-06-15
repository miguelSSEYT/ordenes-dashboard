import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="Ordenes Producibles", layout="wide")
st.title("📄 Análisis de Producción: COOIS, ZCO41 vs MB52")

# Subida de archivos
st.sidebar.header(":inbox_tray: Subir Archivos")
crossref_file = st.sidebar.file_uploader("Tabla de equivalencias (Custom vs Non Custom)", type="xlsx")
mb52_file = st.sidebar.file_uploader("MB52 - Inventario", type="xlsx")
coois_file = st.sidebar.file_uploader("COOIS - Órdenes fijas", type="xlsx")
zco41_file = st.sidebar.file_uploader("ZCO41 - Nueva demanda", type="xlsx")

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

    st.subheader("📦 Inventario Neto por Material")
    st.dataframe(full[['Custom Description', 'Open Quantity', 'Order quantity (GMEIN)', 'Pln.Or Qty', 'Available after COOIS', 'Available after ALL']], use_container_width=True)

    today = pd.to_datetime(date.today())

    coois = coois.merge(full[['Custom Description', 'Open Quantity']], on='Custom Description', how='left')
    zco41 = zco41.merge(full[['Custom Description', 'Available after COOIS']], on='Custom Description', how='left')

    coois['Enough'] = coois['Order quantity (GMEIN)'] <= coois['Open Quantity']
    zco41['Enough'] = zco41['Pln.Or Qty'] <= zco41['Available after COOIS']

    coois['Shortage'] = coois['Order quantity (GMEIN)'] - coois['Open Quantity']
    zco41['Shortage'] = zco41['Pln.Or Qty'] - zco41['Available after COOIS']

    st.subheader("✅ COOIS - Órdenes que se pueden producir")
    st.dataframe(coois[coois['Enough']][['Sales Order', 'Custom Description', 'Order quantity (GMEIN)', 'Open Quantity']], use_container_width=True)

    st.subheader("❌ COOIS - Órdenes que NO se pueden producir")
    st.dataframe(coois[~coois['Enough']][['Sales Order', 'Custom Description', 'Order quantity (GMEIN)', 'Open Quantity', 'Shortage']], use_container_width=True)

    st.subheader("✅ ZCO41 - Órdenes que se pueden producir")
    st.dataframe(zco41[zco41['Enough']][['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Available after COOIS']], use_container_width=True)

    st.subheader("❌ ZCO41 - Órdenes que NO se pueden producir")
    st.dataframe(zco41[~zco41['Enough']][['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Available after COOIS', 'Shortage']], use_container_width=True)

    st.subheader("⏰ Órdenes Past Due que NO se pueden producir (COOIS)")
    coois['Est. Ship Date'] = pd.to_datetime(coois['Est. Ship Date'], errors='coerce')
    past_due_coois = coois[(~coois['Enough']) & (coois['Est. Ship Date'] < today)]
    st.dataframe(past_due_coois[['Sales Order', 'Custom Description', 'Order quantity (GMEIN)', 'Est. Ship Date', 'Shortage']], use_container_width=True)

    st.subheader("⏰ Órdenes Past Due que NO se pueden producir (ZCO41)")
    zco41['Estimated Ship Date'] = pd.to_datetime(zco41['Estimated Ship Date'], errors='coerce')
    past_due_zco41 = zco41[(~zco41['Enough']) & (zco41['Estimated Ship Date'] < today)]
    st.dataframe(past_due_zco41[['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Estimated Ship Date', 'Shortage']], use_container_width=True)

else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
