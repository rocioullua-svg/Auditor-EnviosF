import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuración de la página web (Forzamos limpieza de caché visual)
st.set_page_config(page_title="Auditoría Flexit v2", page_icon="📊", layout="wide")

st.title("📊 Auditoría Inteligente de Envíos v2.0")
st.markdown("Verifica cobros, reclama sobreprecios y detecta zonas faltantes en tu sistema.")

# Tarifas Oficiales Flexit
TARIFA_CABA = 4610.99
TARIFA_GBA1 = 7370.99
TARIFA_GBA2 = 10245.99
TARIFAS_VALIDAS = [TARIFA_CABA, TARIFA_GBA1, TARIFA_GBA2]

def cargar_archivo(file_uploader):
    """Lector inteligente que busca la fila real de títulos ignorando basura arriba"""
    nombre_archivo = file_uploader.name
    archivo_bytes = file_uploader.read()
    
    if nombre_archivo.lower().endswith(('.xls', '.xlsx')):
        df = pd.read_excel(io.BytesIO(archivo_bytes), header=None)
    else:
        df = pd.read_csv(io.BytesIO(archivo_bytes), encoding='latin-1', header=None, on_bad_lines='skip')
        
    # Escáner: Busca inteligentemente la fila que contiene las columnas clave
    header_row_idx = None
    for i, row in df.head(15).iterrows():
        # Busca cualquiera de nuestras palabras clave en las primeras filas
        if row.astype(str).str.contains('Tracking|Pedido|venta ML|ID venta', case=False, na=False).any():
            header_row_idx = i
            break
            
    if header_row_idx is not None:
        df.columns = df.iloc[header_row_idx].fillna('Columna_Sin_Nombre')
        df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
        
    df.columns = df.columns.astype(str).str.strip()
    return df

def limpiar_texto(columna):
    return columna.astype(str).str.strip().str.replace('.0', '', regex=False).str.upper()

# --- INTERFAZ DE CARGA ---
col1, col2 = st.columns(2)
with col1:
    archivo_prov = st.file_uploader("📁 1. Archivo Proveedor (Flexit)", type=['csv', 'xls', 'xlsx'])
with col2:
    archivo_int = st.file_uploader("📁 2. Archivo Interno (Sistema)", type=['csv', 'xls', 'xlsx'])

# --- PROCESAMIENTO ---
if archivo_prov and archivo_int:
    with st.spinner("⚙️ Analizando datos, cruzando viajes y auditando cobros..."):
        
        # 1. Cargar datos
        df_prov = cargar_archivo(archivo_prov)
        df_int = cargar_archivo(archivo_int)
        
        # 2. Limpiar columnas clave equivalentes
        if 'Número Tracking' in df_prov.columns:
            df_prov['Número Tracking'] = limpiar_texto(df_prov['Número Tracking'])
        if 'ID venta ML' in df_prov.columns:
            df_prov['ID venta ML'] = limpiar_texto(df_prov['ID venta ML'])
        if 'CP' in df_prov.columns:
            df_prov['CP_Num'] = pd.to_numeric(df_prov['CP'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
            
        if 'Tracking Code' in df_int.columns:
            df_int['Tracking Code'] = limpiar_texto(df_int['Tracking Code'])
        if 'Nro Pedido' in df_int.columns:
            df_int['Nro Pedido'] = limpiar_texto(df_int['Nro Pedido'])

        if 'Pedido - Transportista' in df_int.columns:
            df_int_flexit = df_int[df_int['Pedido - Transportista'].astype(str).str.contains('FLEX', case=False, na=False)].copy()
        else:
            df_int_flexit = df_int.copy()
            
        # 3. Cruce Principal Seguro
        tiene_tracking = 'Tracking Code' in df_int_flexit.columns and 'Número Tracking' in df_prov.columns
        tiene_pedido = 'Nro Pedido' in df_int_flexit.columns and 'ID venta ML' in df_prov.columns

        if tiene_tracking:
            cruce = pd.merge(df_prov, df_int_flexit, left_on='Número Tracking', right_on='Tracking Code', how='outer', indicator=True)
            metodo_usado = "Trackings"
        elif tiene_pedido:
            cruce = pd.merge(df_prov, df_int_flexit, left_on='ID venta ML', right_on='Nro Pedido', how='outer', indicator=True)
            metodo_usado = "Nro Pedido (ID Venta ML)"
        else:
            st.error("⚠️ Error crítico: El sistema no detectó las columnas clave. Verifica los archivos.")
            st.stop()

        # 4. Cálculos Monetarios
        col_precio_prov = 'Precio Facturado' if 'Precio Facturado' in cruce.columns else 'Precio'
        
        for col in [col_precio_prov, 'Costo de Envío', 'Costo de Envío Cliente']:
            if col in cruce.columns:
                cruce[col] = pd.to_numeric(cruce[col].astype(str).str.replace(',','.'), errors='coerce').fillna(0)
            else:
                cruce[col] = 0

        cruce['Diferencia_vs_Sistema'] = (cruce[col_precio_prov] - cruce['Costo de Envío']).round(2)
        cruce['Costo_Absorbido_Empresa'] = (cruce[col_precio_prov] - cruce['Costo de Envío Cliente']).round(2)

        # 5. Clasificación Estricta para Auditoría
        def clasificar_estado(row):
            cobro_prov = row[col_precio_prov]
            costo_sist = row['Costo de Envío']
            cp = row.get('CP_Num', 0)
            
            if row['_merge'] == 'left_only':
                return 'Cobrado pero No en Sistema'
            elif row['_merge'] == 'right_only':
                return 'En Sistema pero No Cobrado'
            
            if costo_sist > 0:
                if row['Diferencia_vs_Sistema'] > 50:
                    return 'Cobro MAYOR al Sistema'
                elif row['Diferencia_vs_Sistema'] < -50:
                    return 'Cobro MENOR al Sistema'
                else:
                    return 'Coincide OK'
            else:
                if cobro_prov == 0:
                    return 'Falta Zona ($0) - Viaje Gratis'
                    
                es_caba = 1000 <= cp <= 1499
                if es_caba:
                    if (cobro_prov - TARIFA_CABA) > 50:
                        return 'Cobro MAYOR al Sistema (Sobreprecio CABA detectado)'
                    else:
                        return 'Falta Zona en Sistema (Cobro CABA OK)'
                
                es_tarifa_valida = any(abs(cobro_prov - tarifa) <= 10 for tarifa in TARIFAS_VALIDAS)
                if es_tarifa_valida:
                    return 'Falta Zona en Sistema (Tarifa Estándar OK)'
                else:
                    return 'Falta Zona (Cobro IRREGULAR / Número Inventado)'
                
        cruce['Estado_Auditoria'] = cruce.apply(clasificar_estado, axis=1)
        
        def ajustar_diferencia_reclamo(row):
            if row['Estado_Auditoria'] == 'Cobro MAYOR al Sistema (Sobreprecio CABA detectado)':
                return round(row[col_precio_prov] - TARIFA_CABA, 2)
            return row['Diferencia_vs_Sistema']
            
        cruce['Diferencia_vs_Sistema'] = cruce.apply(ajustar_diferencia_reclamo, axis=1)
        
        # 6. Preparar DataFrames específicos
        columnas_deseadas = ['Número Tracking', 'ID venta ML', 'Nro Pedido', 'Estado_Auditoria', 
                             col_precio_prov, 'Costo de Envío', 'Diferencia_vs_Sistema', 
                             'Costo de Envío Cliente', 'Costo_Absorbido_Empresa', 
                             'Fecha Venta', 'Localidad', 'CP', 'Provincia']
        columnas_existentes = [col for col in columnas_deseadas if col in cruce.columns]
        
        reporte_final = cruce[columnas_existentes].copy().fillna('N/A')
        
        condicion_reclamo = reporte_final['Estado_Auditoria'].str.contains('Cobro MAYOR', na=False)
        df_reclamos = reporte_final[condicion_reclamo].copy()
        total_a_reclamar = df_reclamos['Diferencia_vs_Sistema'].sum() if not df_reclamos.empty else 0

        condicion_zonas = reporte_final['Estado_Auditoria'].str.contains('Falta Zona', na=False)
        df_zonas = cruce[condicion_zonas].copy()
        if not df_zonas.empty and all(col in df_zonas.columns for col in ['Provincia', 'Localidad', 'CP']):
            resumen_zonas = df_zonas.groupby(['Provincia', 'Localidad', 'CP', 'Estado_Auditoria']).size().reset_index(name='Cantidad')
            resumen_zonas = resumen_zonas.sort_values(by='Cantidad', ascending=False)
        else:
            resumen_zonas = pd.DataFrame(columns=['Provincia', 'Localidad', 'CP', 'Estado_Auditoria', 'Cantidad'])

        # --- MÉTRICAS EN PANTALLA ---
        st.success(f"¡Auditoría completada exitosamente! Se cruzaron los datos utilizando: **{metodo_usado}**.")
        
        total_prov = cruce[cruce['_merge'] != 'right_only'][col_precio_prov].sum()
        total_absorbido = cruce[cruce['_merge'] != 'right_only']['Costo_Absorbido_Empresa'].sum()
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Facturado", f"$ {total_prov:,.2f}")
        m2.metric("Gasto Empresa", f"$ {total_absorbido:,.2f}")
        m3.metric("A Reclamar", f"$ {total_a_reclamar:,.2f}", delta="Revisar Factura", delta_color="inverse")
        m4.metric("Sin Zona", f"{len(df_zonas)} envíos")
        
        # --- GENERACIÓN DEL EXCEL ---
        wb = openpyxl.Workbook()
        ws_dash = wb.active
        ws_dash.title = "Dashboard"
        ws_data = wb.create_sheet(title="Auditoria Completa")
        ws_reclamos = wb.create_sheet(title="Reclamos a Flexit") 
        ws_zonas = wb.create_sheet(title="Zonas a Corregir")     
        
        header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
        
        def dar_formato_tabla(ws, dataframe):
            if dataframe.empty:
                ws.append(["No hay registros en esta categoría"])
                return
            for r in dataframe_to_rows(dataframe, index=False, header=True):
                ws.append(r)
            for col in ws.iter_cols(min_row=1, max_row=1, max_col=ws.max_column):
                for cell in col:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
                for cell in row:
                    cell.border = thin_border
                    col_name = str(ws.cell(row=1, column=cell.column).value)
                    if any(moneda in col_name for moneda in ['Precio', 'Costo', 'Diferencia', 'Empresa']):
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '$#,##0.00'
            for i, col in enumerate(dataframe.columns):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = 20

        dar_formato_tabla(ws_data, reporte_final)
        dar_formato_tabla(ws_reclamos, df_reclamos)
        dar_formato_tabla(ws_zonas, resumen_zonas)
        
        ws_dash.sheet_view.showGridLines = False
        ws_dash['B2'] = "Dashboard de Control - Flexit"
        ws_dash['B2'].font = Font(size=16, bold=True, color="2F4F4F")
        ws_dash['B4'] = "1. Total Facturado por Proveedor:"
        ws_dash['C4'] = total_prov
        ws_dash['B5'] = "2. Dinero a RECLAMAR (Sobrecobros):"
        ws_dash['C5'] = total_a_reclamar
        ws_dash['C5'].font = Font(color="B22222", bold=True)
        ws_dash['B6'] = "3. Costo Neto Absorbido por Empresa:"
        ws_dash['C6'] = total_absorbido
        
        for r in range(4, 7):
            ws_dash[f'B{r}'].font = Font(bold=True)
            ws_dash[f'C{r}'].number_format = '$#,##0.00'
            
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        st.subheader("📥 Descargar Reporte Final")
        st.download_button(
            label="Descargar Auditoría en Excel",
            data=excel_buffer,
            file_name="Auditoria_Flexit_v2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
