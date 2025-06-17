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
    coois['Tipo'] = coois['Master Material Description'].apply(lambda x: 'SS' if isinstance(x, str) and x.strip().endswith('SS') else 'DC')
    zco41['Tipo'] = zco41['Material description'].apply(lambda x: 'SS' if isinstance(x, str) and x.strip().endswith('SS') else 'DC')
    
    # Resumen por tipo y cantidad
    coois['Cantidad'] = coois['Order Quantity (Item)']
    zco41['Cantidad'] = zco41['Pln.Or Qty']
    
    coois_sum_qty = coois.groupby('Tipo')['Cantidad'].sum().reset_index()
    coois_sum_qty['Fuente'] = 'COOIS'
    
    zco41_sum_qty = zco41.groupby('Tipo')['Cantidad'].sum().reset_index()
    zco41_sum_qty['Fuente'] = 'ZCO41'
    
    resumen_qty = pd.concat([coois_sum_qty, zco41_sum_qty], ignore_index=True)
    
    coois['Sales office'] = coois['Sales office'].fillna('null')
    zco41['Sales office'] = zco41['Sales office'].fillna('null')
    
    coois_office_summary = coois.groupby('Sales office')['Order Quantity (Item)'].sum().reset_index()
    coois_office_summary['Fuente'] = 'COOIS'
    
    zco41_office_summary = zco41.groupby('Sales office')['Pln.Or Qty'].sum().reset_index()
    zco41_office_summary['Fuente'] = 'ZCO41'
    
    coois_office_summary = coois_office_summary.rename(columns={
        'Sales office': 'Tipo de Orden',
        'Order Quantity (Item)': 'Cantidad'
    })
    zco41_office_summary = zco41_office_summary.rename(columns={
        'Sales office': 'Tipo de Orden',
        'Pln.Or Qty': 'Cantidad'
    })
    
    resumen_ordenes = pd.concat([coois_office_summary, zco41_office_summary], ignore_index=True)
    
    st.header("🆚 Vasos SS vs DC")
    st.subheader("Cantidad Total de Vasos por Tipo")
    st.dataframe(resumen_qty)
    
    st.header("🔢 Tipo de Orden")
    st.subheader("Cantidad Total de Vasos por Tipo de Orden (Sales office)")
    st.dataframe(resumen_ordenes)
    
    # Preparar equivalencia
    crossref = crossref.rename(columns={"Non Custom": "Material description", "Custom": "Custom Description"})
    mb52_grouped = mb52.groupby('Material description', as_index=False)['Open Quantity'].sum()
    mb52_custom = mb52_grouped.merge(crossref, on='Material description', how='left')
    
    # Verificar referencias no encontradas
    referencias_faltantes_mb52 = mb52_custom[mb52_custom['Custom Description'].isna()]['Material description'].dropna().unique()
    referencias_cross = set(crossref['Custom Description'].unique())
    referencias_coois = set(coois['Master Material Description'].dropna().unique())
    referencias_zco41 = set(zco41['Material description'].dropna().unique())
    
    faltantes_coois = sorted(list(referencias_coois - referencias_cross))
    faltantes_zco41 = sorted(list(referencias_zco41 - referencias_cross))
    
    referencias_no_encontradas = pd.DataFrame({
        "Fuente": ["MB52"] * len(referencias_faltantes_mb52) + ["COOIS"] * len(faltantes_coois) + ["ZCO41"] * len(faltantes_zco41),
        "Descripción no encontrada": list(referencias_faltantes_mb52) + faltantes_coois + faltantes_zco41
    })
    
    if not referencias_no_encontradas.empty:
        st.warning("⚠️ Existen descripciones que no están en la tabla de equivalencias. Corrige esto antes de continuar.")
        st.dataframe(referencias_no_encontradas)
        st.stop()
    
    # Agrupar COOIS y ZCO41
    coois = coois.rename(columns={'Master Material Description': 'Custom Description'})
    zco41 = zco41.rename(columns={'Material description': 'Custom Description'})
    
    coois_sum = coois.groupby('Custom Description', as_index=False)['Order Quantity (Item)'].sum()
    zco41_sum = zco41.groupby('Custom Description', as_index=False)['Pln.Or Qty'].sum()
    
    full = mb52_custom[['Custom Description', 'Material description', 'Open Quantity']].merge(coois_sum, on='Custom Description', how='left')\
        .merge(zco41_sum, on='Custom Description', how='left').fillna(0)
    
    full['Available after COOIS'] = full['Open Quantity'] - full['Order Quantity (Item)']
    full['Available after ALL'] = full['Available after COOIS'] - full['Pln.Or Qty']
    
    zco41_eval = zco41.merge(full[['Custom Description', 'Available after COOIS']], on='Custom Description', how='left')
    zco41_eval['Available after COOIS'] = zco41_eval['Available after COOIS'].fillna(0)
    zco41_eval['Can Produce'] = zco41_eval['Pln.Or Qty'] <= zco41_eval['Available after COOIS']
    
    coois_eval = coois.merge(full[['Custom Description', 'Open Quantity']], on='Custom Description', how='left')
    coois_eval['Open Quantity'] = coois_eval['Open Quantity'].fillna(0)
    coois_eval['Can Produce'] = coois_eval['Order Quantity (Item)'] <= coois_eval['Open Quantity']
    
    zco41_orders = zco41_eval.groupby('Sales Order')['Can Produce'].all().reset_index()
    coois_orders = coois_eval.groupby('Sales document')['Can Produce'].all().reset_index()
    
    zco41_eval = zco41_eval.merge(zco41_orders, on='Sales Order', suffixes=('', '_order'))
    coois_eval = coois_eval.merge(coois_orders, on='Sales document', suffixes=('', '_order'))
    
    st.header("📊 Resultados del Análisis")
    
    with st.expander("ZCO41 - Órdenes COMPLETAS que SÍ se pueden producir"):
        st.dataframe(zco41_eval[zco41_eval['Can Produce_order']])

    with st.expander("ZCO41 - Órdenes COMPLETAS que NO se pueden producir"):
        df = zco41_eval[~zco41_eval['Can Produce_order']].copy()
        df['Net Inventory'] = df['Available after COOIS'] - df['Pln.Or Qty']
        df['Reason'] = df.apply(lambda row: (
            "Can be produced — enough inventory available"
            if row['Available after COOIS'] >= row['Pln.Or Qty']
            else "Cannot be produced — not enough inventory"
        ), axis=1)
        st.dataframe(df[['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Available after COOIS', 'Net Inventory', 'Reason']])
    
    with st.expander("COOIS - Órdenes COMPLETAS que NO se pueden producir"):
        df = coois_eval[~coois_eval['Can Produce_order']].copy()
        df['Net Inventory'] = df['Open Quantity'] - df['Order Quantity (Item)']
        df['Reason'] = df.apply(lambda row: (
    "Sales document " + str(row['Sales document']) +
    " needs " + (str(int(row['Order Quantity (Item)'])) if pd.notnull(row['Order Quantity (Item)']) else "N/A") +
    " units of '" + str(row['Custom Description']) + "', but only " +
    (str(int(row['Open Quantity'])) if pd.notnull(row['Open Quantity']) else "N/A") +
    " are available. Shortage: " +
    (str(int(row['Order Quantity (Item)'] - row['Open Quantity'])) if pd.notnull(row['Order Quantity (Item)']) and pd.notnull(row['Open Quantity']) else "N/A")
    ), axis=1)
        st.dataframe(df[['Sales document', 'Custom Description', 'Order Quantity (Item)', 'Open Quantity', 'Net Inventory', 'Reason']])
    
    with st.expander("Past Due - ZCO41 y COOIS que NO se pueden producir"):
        zco41_past_due = zco41_eval[(~zco41_eval['Can Produce_order']) & (pd.to_datetime(zco41_eval['Estimated Ship Date']) < today)]
        coois_past_due = coois_eval[(~coois_eval['Can Produce_order']) & (pd.to_datetime(coois_eval['Estimated Ship Date (header)']) < today)]
        st.subheader("ZCO41 - Past Due")
        st.dataframe(zco41_past_due)
        st.subheader("COOIS - Past Due")
        st.dataframe(coois_past_due)
    
    with st.expander("Material Requerido para Cumplir Producción"):
        coois_eval['Cantidad Faltante'] = coois_eval['Order Quantity (Item)'] - coois_eval['Open Quantity']
        zco41_eval['Cantidad Faltante'] = zco41_eval['Pln.Or Qty'] - zco41_eval['Available after COOIS']
    
        faltantes_total = pd.concat([
            coois_eval[~coois_eval['Can Produce']][['Custom Description', 'Cantidad Faltante']],
            zco41_eval[~zco41_eval['Can Produce']][['Custom Description', 'Cantidad Faltante']]
        ])
    
        faltantes_grouped = faltantes_total.groupby('Custom Description', as_index=False).sum()
        faltantes_grouped = faltantes_grouped[faltantes_grouped['Cantidad Faltante'] > 0]
    
        faltantes_con_non_custom = faltantes_grouped.merge(crossref, on='Custom Description', how='left')
        faltantes_con_non_custom = faltantes_con_non_custom[['Custom Description', 'Material description', 'Cantidad Faltante']]
        st.dataframe(faltantes_con_non_custom.sort_values(by='Cantidad Faltante', ascending=False))
    
    # Recalcular columnas necesarias para exportación
    zco41_ok = zco41_eval[zco41_eval['Can Produce_order']]
    zco41_nok = zco41_eval[~zco41_eval['Can Produce_order']].copy()
    zco41_nok['Net Inventory'] = zco41_nok['Available after COOIS'] - zco41_nok['Pln.Or Qty']
    zco41_nok['Reason'] = zco41_nok.apply(lambda row: (
        "Sales Order " + str(row['Sales Order']) + " needs " + str(int(row['Pln.Or Qty'])) +
        " units of '" + row['Custom Description'] + "', but only " + str(int(row['Available after COOIS'])) +
        " are available. Shortage: " + str(int(row['Pln.Or Qty'] - row['Available after COOIS']))
    ), axis=1)
    
    coois_nok = coois_eval[~coois_eval['Can Produce_order']].copy()
    coois_nok['Net Inventory'] = coois_nok['Open Quantity'] - coois_nok['Order Quantity (Item)']
    coois_nok['Reason'] = coois_nok.apply(lambda row: (
    "Sales document " + str(row['Sales document']) +
    " needs " + (str(int(row['Order Quantity (Item)'])) if pd.notnull(row['Order Quantity (Item)']) else "N/A") +
    " units of '" + str(row['Custom Description']) + "', but only " +
    (str(int(row['Open Quantity'])) if pd.notnull(row['Open Quantity']) else "N/A") +
    " are available. Shortage: " +
    (str(int(row['Order Quantity (Item)'] - row['Open Quantity'])) if pd.notnull(row['Order Quantity (Item)']) and pd.notnull(row['Open Quantity']) else "N/A")
    ), axis=1)
    faltantes_sorted = faltantes_con_non_custom.sort_values(by='Cantidad Faltante', ascending=False)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        zco41_ok.to_excel(writer, sheet_name='ZCO41 - OK', index=False)
        zco41_nok.to_excel(writer, sheet_name='ZCO41 - NOK', index=False)
        coois_nok.to_excel(writer, sheet_name='COOIS - NOK', index=False)
        zco41_past_due.to_excel(writer, sheet_name='ZCO41 Past Due', index=False)
        coois_past_due.to_excel(writer, sheet_name='COOIS Past Due', index=False)
        faltantes_sorted.to_excel(writer, sheet_name='Material Faltante', index=False)
        output.seek(0)
    
    st.download_button(
        label="Descargar análisis completo en Excel",
        data=output.getvalue(),
        file_name="analisis_produccion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
