from io import BytesIO
import streamlit as st
import pandas as pd
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
    coois = pd.read_excel(coois_file, sheet_name='Sheet1')
    crossref = pd.read_excel(crossref_file, sheet_name='CROSSREFERENCE SAP')
    mb52 = pd.read_excel(mb52_file, sheet_name='Sheet1')
    zco41 = pd.read_excel(zco41_file, sheet_name='Sheet1')

    # Clasificar DC y SS
    coois['Tipo'] = coois['Master Material Description'].apply(lambda x: 'SS' if isinstance(x, str) and x.strip().endswith('SS') else 'DC')
    zco41['Tipo'] = zco41['Material description'].apply(lambda x: 'SS' if isinstance(x, str) and x.strip().endswith('SS') else 'DC')

    def map_tipo_orden(valor):
        if valor in ['BDV', 'CTL']:
            return 'BDV'
        elif valor in ['ECM', 'YRD']:
            return 'ECM'
        else:
            return 'BDV'

    coois['Sales office'] = coois['Sales office'].fillna('null')
    zco41['Sales office'] = zco41['Sales office'].fillna('null')

    coois_office_summary = coois.groupby('Sales office')['Order Quantity (Item)'].sum().reset_index()
    coois_office_summary = coois_office_summary.rename(columns={'Sales office': 'Tipo de Orden', 'Order Quantity (Item)': 'Cantidad'})
    coois_office_summary['Tipo de Orden'] = coois_office_summary['Tipo de Orden'].apply(map_tipo_orden)
    coois_office_summary = coois_office_summary.groupby('Tipo de Orden', as_index=False)['Cantidad'].sum()
    coois_office_summary['Fuente'] = 'COOIS'

    zco41_office_summary = zco41.groupby('Sales office')['Pln.Or Qty'].sum().reset_index()
    zco41_office_summary = zco41_office_summary.rename(columns={'Sales office': 'Tipo de Orden', 'Pln.Or Qty': 'Cantidad'})
    zco41_office_summary['Tipo de Orden'] = zco41_office_summary['Tipo de Orden'].apply(map_tipo_orden)
    zco41_office_summary = zco41_office_summary.groupby('Tipo de Orden', as_index=False)['Cantidad'].sum()
    zco41_office_summary['Fuente'] = 'ZCO41'

    resumen_ordenes = pd.concat([coois_office_summary, zco41_office_summary], ignore_index=True)
    st.header("🔢 Tipo de Orden")
    st.subheader("Cantidad Total de Vasos por Tipo de Orden (Sales office)")
    st.dataframe(resumen_ordenes.style.format({"Cantidad": ":,.0f"}))

    selected_model = st.text_input("🔍 Filtrar por modelo (Custom Description):", value="")
    if selected_model:
        zco41 = zco41[zco41['Material description'].str.strip() == selected_model.strip()]

    crossref = crossref.rename(columns={"Non Custom": "Material description", "Custom": "Custom Description"})
    mb52_grouped = mb52.groupby('Material description', as_index=False)['Open Quantity'].sum()
    mb52_custom = mb52_grouped.merge(crossref, on='Material description', how='left')

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
    zco41_eval.sort_values(by=['Sales Order', 'Custom Description'], inplace=True)

    cumulative_inventory = full.set_index('Custom Description')['Available after COOIS'].to_dict()
    cumulative_tracker = cumulative_inventory.copy()

    producible = []
    grouped_orders = zco41_eval.groupby('Sales Order')

    for order, group in grouped_orders:
        local_tracker = cumulative_tracker.copy()
        order_producible = True
        for _, row in group.iterrows():
            material = row['Custom Description']
            qty = row['Pln.Or Qty']
            if local_tracker.get(material, 0) >= qty:
                local_tracker[material] -= qty
            else:
                order_producible = False
                break
        if order_producible:
            cumulative_tracker = local_tracker.copy()
        producible.extend([order_producible] * len(group))

    zco41_eval['Can Produce_order'] = producible

    st.header("📊 Resultados del Análisis")

    with st.expander("ZCO41 - Órdenes COMPLETAS que SÍ se pueden producir"):
        st.dataframe(zco41_eval[zco41_eval['Can Produce_order']])

    with st.expander("ZCO41 - Órdenes COMPLETAS que NO se pueden producir"):
        df = zco41_eval[~zco41_eval['Can Produce_order']].copy()
        df['Net Inventory'] = df['Available after COOIS'] - df['Pln.Or Qty']
        df['Reason'] = df.apply(lambda row: (
            "Sales Order " + str(row['Sales Order']) +
            " needs " + str(int(row['Pln.Or Qty'])) +
            " units of '" + row['Custom Description'] + "', but only " + str(int(row['Available after COOIS'])) +
            " are available. Shortage: " + str(int(row['Pln.Or Qty'] - row['Available after COOIS']))
        ), axis=1)
        st.dataframe(df[['Sales Order', 'Custom Description', 'Pln.Or Qty', 'Available after COOIS', 'Net Inventory', 'Reason']])

    zco41_ok = zco41_eval[zco41_eval['Can Produce_order']]
    zco41_nok = zco41_eval[~zco41_eval['Can Produce_order']].copy()
    zco41_nok['Net Inventory'] = zco41_nok['Available after COOIS'] - zco41_nok['Pln.Or Qty']
    zco41_nok['Reason'] = zco41_nok.apply(lambda row: (
        "Sales Order " + str(row['Sales Order']) + " needs " + str(int(row['Pln.Or Qty'])) +
        " units of '" + row['Custom Description'] + "', but only " + str(int(row['Available after COOIS'])) +
        " are available. Shortage: " + str(int(row['Pln.Or Qty'] - row['Available after COOIS']))
    ), axis=1)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        zco41_ok.to_excel(writer, sheet_name='ZCO41 - OK', index=False)
        zco41_nok.to_excel(writer, sheet_name='ZCO41 - NOK', index=False)
        output.seek(0)

    st.download_button(
        label="⬇️ Descargar análisis en Excel",
        data=output.getvalue(),
        file_name="ordenes_filtradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
