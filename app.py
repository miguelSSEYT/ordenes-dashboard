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

if crossref_file and mb52_file and coois_file and zco41_file:
    crossref = pd.read_excel(crossref_file)
    mb52 = pd.read_excel(mb52_file)
    coois = pd.read_excel(coois_file)
    zco41 = pd.read_excel(zco41_file)

    crossref = crossref.rename(columns={"Non Custom": "Material description", "Custom": "Custom Description"})
    mb52 = mb52.rename(columns={"Material description": "Material description"})

    # Validar referencias existentes
    all_descriptions = set(mb52['Material description'].unique())
    missing_ref = all_descriptions - set(crossref['Material description'].unique())
    if missing_ref:
        st.error(f"Las siguientes descripciones de MB52 no están en la tabla de equivalencias:\n{missing_ref}")
        st.stop()

    coois = coois.rename(columns={"Material description": "Custom Description"})
    zco41 = zco41.rename(columns={"Material description": "Custom Description"})

    all_custom_desc = set(coois['Custom Description'].unique()).union(set(zco41['Custom Description'].unique()))
    missing_customs = all_custom_desc - set(crossref['Custom Description'].unique())
    if missing_customs:
        st.error(f"Las siguientes descripciones CUSTOM no están en la tabla de equivalencias:\n{missing_customs}")
        st.stop()

    # Convertir MB52 a Custom
    mb52 = mb52.merge(crossref, on='Material description', how='left')
    mb52_grouped = mb52.groupby('Custom Description', as_index=False)['Open Quantity'].sum()

    # COOIS y ZCO41 agrupados
    coois_grouped = coois.groupby(['Sales Order', 'Custom Description'], as_index=False)['Order quantity (GMEIN)'].sum()
    zco41_grouped = zco41.groupby(['Sales Order', 'Custom Description'], as_index=False)['Pln.Or Qty'].sum()

    today = pd.to_datetime(date.today())

    # ZCO41 análisis por línea
    zco41_merged = zco41_grouped.merge(mb52_grouped, on='Custom Description', how='left')
    zco41_merged['Disponibilidad'] = zco41_merged['Open Quantity'] - zco41_merged['Pln.Or Qty']
    zco41_merged['Puede importar'] = zco41_merged['Disponibilidad'] >= 0

    # COOIS análisis por línea
    coois_merged = coois_grouped.merge(mb52_grouped, on='Custom Description', how='left')
    coois_merged['Disponibilidad'] = coois_merged['Open Quantity'] - coois_merged['Order quantity (GMEIN)']
    coois_merged['Puede producir'] = coois_merged['Disponibilidad'] >= 0

    # ZCO41 past due y SS
    zco41_dates = zco41[['Sales Order', 'Custom Description', 'Estimated Ship Date']].drop_duplicates()
    zco41_dates['Estimated Ship Date'] = pd.to_datetime(zco41_dates['Estimated Ship Date'], errors='coerce')
    zco41_full = zco41_merged.merge(zco41_dates, on=['Sales Order', 'Custom Description'], how='left')
    zco41_past_due_ss = zco41_full[(~zco41_full['Puede importar']) & 
                                   (zco41_full['Custom Description'].str.endswith("SS")) &
                                   (zco41_full['Estimated Ship Date'] < today)]

    # Lista general de materiales faltantes
    coois_needs = coois_merged[~coois_merged['Puede producir']].copy()
    zco41_needs = zco41_merged[~zco41_merged['Puede importar']].copy()
    coois_needs['Faltante'] = coois_needs['Order quantity (GMEIN)'] - coois_needs['Open Quantity']
    zco41_needs['Faltante'] = zco41_needs['Pln.Or Qty'] - zco41_needs['Open Quantity']
    total_needs = pd.concat([
        coois_needs[['Custom Description', 'Faltante']],
        zco41_needs[['Custom Description', 'Faltante']]
    ]).groupby('Custom Description', as_index=False).sum()

    # Mostrar resultados
    with st.expander("✅ Órdenes que sí se pueden importar de ZCO41"):
        st.dataframe(zco41_merged[zco41_merged['Puede importar']], use_container_width=True)

    with st.expander("❌ Órdenes que NO se pueden importar de ZCO41"):
        for order in zco41_merged[~zco41_merged['Puede importar']]['Sales Order'].unique():
            sub = zco41_merged[(zco41_merged['Sales Order'] == order)]
            total_lines = len(sub)
            not_importable = sub[~sub['Puede importar']]
            st.write(f"Orden {order} tiene {total_lines} líneas, de las cuales {len(not_importable)} no se pueden importar por falta de inventario:")
            st.dataframe(not_importable, use_container_width=True)

    with st.expander("❌ Órdenes de COOIS que NO se pueden producir"):
        for order in coois_merged[~coois_merged['Puede producir']]['Sales Order'].unique():
            sub = coois_merged[(coois_merged['Sales Order'] == order)]
            total_lines = len(sub)
            not_producible = sub[~sub['Puede producir']]
            st.write(f"Orden {order} tiene {total_lines} líneas, de las cuales {len(not_producible)} no se pueden producir por falta de inventario:")
            st.dataframe(not_producible, use_container_width=True)

    with st.expander("⏰ Past Due de ZCO41 - Solo vasos SS que NO se pueden producir"):
        st.dataframe(zco41_past_due_ss[['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Open Quantity', 'Estimated Ship Date', 'Disponibilidad']], use_container_width=True)

    with st.expander("📦 Lista General de Productos Faltantes"):
        st.dataframe(total_needs, use_container_width=True)

else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
