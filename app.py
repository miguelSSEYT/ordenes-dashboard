import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.backends.backend_pdf import PdfPages

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

    # Gráfico de barras
    fig, ax = plt.subplots(figsize=(6, 1.5))
    ax.barh(["Units shipped"], [units_shipped], color="green")
    ax.barh(["Units shipped"], [balance], left=[units_shipped], color="lightgray")
    ax.set_xlim(0, max(forecast, units_shipped + balance))
    ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.set_title("Progress vs Forecast", fontsize=10)
    st.pyplot(fig)

    # Botón para exportar a PDF
    if st.button("📤 Descargar resumen diario en PDF"):
        buffer = BytesIO()
        with PdfPages(buffer) as pdf:
            # Agregar gráfico
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)

            # Agregar tabla como texto
            fig_table, ax_table = plt.subplots(figsize=(6, 2))
            ax_table.axis('off')
            table_data = resumen_diario.values.tolist()
            col_labels = resumen_diario.columns.tolist()
            table = ax_table.table(cellText=table_data, colLabels=col_labels, loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1.2, 1.2)
            pdf.savefig(fig_table, bbox_inches='tight')
            plt.close(fig_table)

            # Si los archivos están cargados, hacer el análisis básico y agregarlo al PDF
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

                # Agregar faltantes como tabla
                fig_falt, ax_falt = plt.subplots(figsize=(6, min(10, 0.25*len(faltantes))))
                ax_falt.axis('off')
                table_data = faltantes.values.tolist()
                col_labels = faltantes.columns.tolist()
                table = ax_falt.table(cellText=table_data, colLabels=col_labels, loc='center')
                table.auto_set_font_size(False)
                table.set_fontsize(8)
                table.scale(1.2, 1.2)
                pdf.savefig(fig_falt, bbox_inches='tight')
                plt.close(fig_falt)

        buffer.seek(0)
        st.download_button(
            label="📥 Descargar PDF",
            data=buffer,
            file_name=f"resumen_diario_{datetime.today().date()}.pdf",
            mime="application/pdf"
        )

# Continuación lógica solo si hay archivos cargados
if crossref_file and mb52_file and coois_file and zco41_file:
    st.success("Archivos cargados correctamente")
else:
    st.info("Por favor, sube los cuatro archivos para iniciar el análisis.")
