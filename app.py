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

    today = pd.to_datetime(date.today())

    coois = coois.merge(full[['Custom Description', 'Open Quantity']], on='Custom Description', how='left')
    zco41 = zco41.merge(full[['Custom Description', 'Available after COOIS']], on='Custom Description', how='left')

    coois['Enough'] = coois['Order quantity (GMEIN)'] <= coois['Open Quantity']
    zco41['Enough'] = zco41['Pln.Or Qty'] <= zco41['Available after COOIS']

    with st.expander("✅ COOIS - Producibles"):
        for _, row in coois[coois['Enough']].iterrows():
            st.write(f"✅ This line can be produced: Order {row['Sales Order']} - {row['Custom Description']} - Qty: {row['Order quantity (GMEIN)']:.0f}")

    with st.expander("❌ COOIS - No Producibles"):
        for _, row in coois[~coois['Enough']].iterrows():
            shortage = row['Order quantity (GMEIN)'] - row['Open Quantity']
            st.write(f"❌ This line cannot be produced: Order {row['Sales Order']} - {row['Custom Description']} - Qty: {row['Order quantity (GMEIN)']:.0f}, Inventory: {row['Open Quantity']:.0f} → Shortage: {shortage:.0f}")

    with st.expander("✅ ZCO41 - Producibles"):
        for _, row in zco41[zco41['Enough']].iterrows():
            st.write(f"✅ This line can be produced: Order {row['Sales Order']} - {row['Custom Description']} - Qty: {row['Pln.Or Qty']:.0f}")

    with st.expander("❌ ZCO41 - No Producibles"):
        for _, row in zco41[~zco41['Enough']].iterrows():
            shortage = row['Pln.Or Qty'] - row['Available after COOIS']
            st.write(f"❌ This line cannot be produced: Order {row['Sales Order']} - {row['Custom Description']} - Qty: {row['Pln.Or Qty']:.0f}, Inventory: {row['Available after COOIS']:.0f} → Shortage: {shortage:.0f}")

else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
