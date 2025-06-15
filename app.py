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

if crossref_file and mb52_file and coois_file and zco41_file:
    st.success("Archivos cargados correctamente")

    today = pd.to_datetime(datetime.today().date())

    # Leer archivos
    crossref = pd.read_excel(crossref_file, sheet_name=0)
    mb52 = pd.read_excel(mb52_file, sheet_name=0)
    coois = pd.read_excel(coois_file, sheet_name=0)
    zco41 = pd.read_excel(zco41_file, sheet_name=0)

    # Clasificar DC y SS
    coois['Tipo'] = coois['Material description'].apply(lambda x: 'SS' if isinstance(x, str) and x.strip().endswith('SS') else 'DC')
    zco41['Tipo'] = zco41['Material description'].apply(lambda x: 'SS' if isinstance(x, str) and x.strip().endswith('SS') else 'DC')

    # Resumen por tipo y cantidad
    coois['Cantidad'] = coois['Order quantity (GMEIN)']
    zco41['Cantidad'] = zco41['Pln.Or Qty']

    coois_sum_qty = coois.groupby('Tipo')['Cantidad'].sum().reset_index()
    coois_sum_qty['Fuente'] = 'COOIS'
    coois_lineas = coois['Tipo'].value_counts().reset_index()
    coois_lineas.columns = ['Tipo', 'Lineas']

    zco41_sum_qty = zco41.groupby('Tipo')['Cantidad'].sum().reset_index()
    zco41_sum_qty['Fuente'] = 'ZCO41'
    zco41_lineas = zco41['Tipo'].value_counts().reset_index()
    zco41_lineas.columns = ['Tipo', 'Lineas']

    resumen_qty = pd.concat([coois_sum_qty, zco41_sum_qty], ignore_index=True)
    resumen_lineas = pd.concat([coois_lineas.assign(Fuente='COOIS'), zco41_lineas.assign(Fuente='ZCO41')], ignore_index=True)

    st.header("🔢 Resumen de Vasos SS vs DC")
    st.subheader("Cantidad Total de Vasos por Tipo")
    st.dataframe(resumen_qty)
    st.subheader("Cantidad de Líneas por Tipo")
    st.dataframe(resumen_lineas)

    # Preparar equivalencia
    crossref = crossref.rename(columns={"Non Custom": "Material description", "Custom": "Custom Description"})
    mb52_grouped = mb52.groupby('Material description', as_index=False)['Open Quantity'].sum()
    mb52_custom = mb52_grouped.merge(crossref, on='Material description', how='left')

    # Agrupar COOIS y ZCO41
    coois = coois.rename(columns={'Material description': 'Custom Description'})
    zco41 = zco41.rename(columns={'Material description': 'Custom Description'})

    coois_sum = coois.groupby('Custom Description', as_index=False)['Order quantity (GMEIN)'].sum()
    zco41_sum = zco41.groupby('Custom Description', as_index=False)['Pln.Or Qty'].sum()

    # Combinar todo
    full = mb52_custom[['Custom Description', 'Open Quantity']].merge(coois_sum, on='Custom Description', how='left')\
        .merge(zco41_sum, on='Custom Description', how='left').fillna(0)

    full['Available after COOIS'] = full['Open Quantity'] - full['Order quantity (GMEIN)']
    full['Available after ALL'] = full['Available after COOIS'] - full['Pln.Or Qty']

    # Evaluar ZCO41 por línea
    zco41_eval = zco41.merge(full[['Custom Description', 'Available after COOIS']], on='Custom Description', how='left')
    zco41_eval['Available after COOIS'] = zco41_eval['Available after COOIS'].fillna(0)
    zco41_eval['Can Produce'] = zco41_eval['Pln.Or Qty'] <= zco41_eval['Available after COOIS']

    # Evaluar COOIS por línea
    coois_eval = coois.merge(full[['Custom Description', 'Open Quantity']], on='Custom Description', how='left')
    coois_eval['Open Quantity'] = coois_eval['Open Quantity'].fillna(0)
    coois_eval['Can Produce'] = coois_eval['Order quantity (GMEIN)'] <= coois_eval['Open Quantity']

    # Evaluación por orden completa (Sales Order)
    zco41_orders = zco41_eval.groupby('Sales Order')['Can Produce'].all().reset_index()
    coois_orders = coois_eval.groupby('Sales Order')['Can Produce'].all().reset_index()

    zco41_eval = zco41_eval.merge(zco41_orders, on='Sales Order', suffixes=('', '_order'))
    coois_eval = coois_eval.merge(coois_orders, on='Sales Order', suffixes=('', '_order'))

    # Mostrar resultados
    st.header("📊 Resultados del Análisis")

    with st.expander("ZCO41 - Órdenes COMPLETAS que SÍ se pueden producir"):
        st.dataframe(zco41_eval[zco41_eval['Can Produce_order']])

    with st.expander("ZCO41 - Órdenes COMPLETAS que NO se pueden producir"):
        df = zco41_eval[~zco41_eval['Can Produce_order']].copy()
        df['Net Inventory'] = df['Available after COOIS'] - df['Pln.Or Qty']
        df['Reason'] = df.apply(lambda row: (
            "Sales Order " + str(row['Sales Order']) + " needs " + str(int(row['Pln.Or Qty'])) +
            " units of '" + row['Custom Description'] + "', but only " + str(int(row['Available after COOIS'])) +
            " are available. Shortage: " + str(int(row['Pln.Or Qty'] - row['Available after COOIS']))
        ) if row['Pln.Or Qty'] > row['Available after COOIS'] else (
            "Sales Order " + str(row['Sales Order']) + " has sufficient inventory for '" + row['Custom Description'] + "'."
        ), axis=1)
        st.dataframe(df[['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Available after COOIS', 'Net Inventory', 'Reason']])

    with st.expander("COOIS - Órdenes COMPLETAS que NO se pueden producir"):
        df = coois_eval[~coois_eval['Can Produce_order']].copy()
        df['Net Inventory'] = df['Open Quantity'] - df['Order quantity (GMEIN)']
        df['Reason'] = df.apply(lambda row: (
            "Sales Order " + str(row['Sales Order']) + " needs " + str(int(row['Order quantity (GMEIN)'])) +
            " units of '" + row['Custom Description'] + "', but only " + str(int(row['Open Quantity'])) +
            " are available. Shortage: " + str(int(row['Order quantity (GMEIN)'] - row['Open Quantity']))
        ) if row['Order quantity (GMEIN)'] > row['Open Quantity'] else (
            "Sales Order " + str(row['Sales Order']) + " has sufficient inventory for '" + row['Custom Description'] + "'."
        ), axis=1)
        st.dataframe(df[['Sales Order', 'Custom Description', 'Order quantity (GMEIN)', 'Open Quantity', 'Net Inventory', 'Reason']])

    with st.expander("⚠️ Past Due - ZCO41 y COOIS que NO se pueden producir"):
        zco41_past_due = zco41_eval[(~zco41_eval['Can Produce_order']) & (pd.to_datetime(zco41_eval['Estimated Ship Date']) < today)]
        coois_past_due = coois_eval[(~coois_eval['Can Produce_order']) & (pd.to_datetime(coois_eval['Est. Ship Date']) < today)]
        st.subheader("ZCO41 - Past Due")
        st.dataframe(zco41_past_due)
        st.subheader("COOIS - Past Due")
        st.dataframe(coois_past_due)
else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
