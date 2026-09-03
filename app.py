import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuración de la página
st.set_page_config(page_title="Auditoría Flexit v3", page_icon="📊", layout="wide")

st.title("📊 Auditoría Definitiva de Envíos v3.0")
st.markdown("Verifica cobros bajo tarifario oficial, excluye Snow Flex y detecta zonas faltantes.")

# Tarifas Oficiales Flexit
TARIFA_CABA = 4610.99
TARIFA_GBA1 = 7370.99
TARIFA_GBA2 = 10245.99
TARIFAS_VALIDAS = [TARIFA_CABA, TARIFA_GBA1, TARIFA_GBA2]

def cargar_archivo(file_uploader):
    """Escáner inteligente de archivos"""
    nombre_archivo = file_uploader.name
    archivo_bytes = file_uploader.read()
    
    if nombre_archivo.lower().endswith(('.xls', '.xlsx')):
        df = pd.read_excel(io.BytesIO(archivo_bytes), header=None)
    else:
        df = pd.read_csv(io.BytesIO(archivo_bytes), encoding='latin-1', header=None, on_bad_lines='skip')
        
    header_row_idx = None
    for i, row in df.head(15).iterrows():
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
    with st.spinner("⚙️ Analizando y auditando con Tarifario Oficial..."):
        
        df_prov = cargar_archivo(archivo_prov)
        df_int = cargar_archivo(archivo_int)
        
        # Limpieza Proveedor
        if 'Número Tracking' in df_prov.columns:
            df_prov['Número Tracking'] = limpiar_texto(df_prov['Número Tracking'])
        if 'ID venta ML' in df_prov.columns:
            df_prov['ID venta ML'] = limpiar_texto(df_prov['ID venta ML'])
        if 'CP' in df_prov.columns:
            df_prov['CP_Num'] = pd.to_numeric(df_prov['CP'].astype(str).str.extract(r'(\d+)')[0], errors='coerce').fillna(0)
            
        # Limpieza Sistema
        if 'Tracking Code' in df_int.columns:
            df_int['Tracking Code'] = limpiar_texto(df_int['Tracking Code'])
        if 'Nro Pedido' in df_int.columns:
            df_int['Nro Pedido'] = limpiar_texto(df_int['Nro Pedido'])

        # Filtro estricto de Transporte: Contiene FLEX pero NO SNOW
        if 'Pedido - Transportista' in df_int.columns:
            transporte = df_int['Pedido - Transportista'].astype(str).str.upper()
            condicion_flex = transporte.str.contains('FLEX') & ~transporte.str.contains('SNOW')
            df_int_flexit = df_int[condicion_flex].copy()
        else:
            df_int_flexit = df_int.copy()
            
        # Cruce Seguro
        tiene_tracking = 'Tracking Code' in df_int_flexit.columns and 'Número Tracking' in df_prov.columns
        tiene_pedido = 'Nro Pedido' in df_int_flexit.columns and 'ID venta ML' in df_prov.columns

        if tiene_tracking:
            cruce = pd.merge(df_prov, df_int_flexit, left_on='Número Tracking', right_on='Tracking Code', how='outer', indicator=True)
        elif tiene_pedido:
            cruce = pd.merge(df_prov, df_int_flexit, left_on='ID venta ML', right_on='Nro Pedido', how='outer', indicator=True)
        else:
            st.error("⚠️ Error: No se detectaron columnas clave ('Tracking Code' o 'Nro Pedido').")
            st.stop()

        # Configuración Monetaria
        col_precio_prov = 'Precio Facturado' if 'Precio Facturado' in cruce.columns else 'Precio'
        for col in [col_precio_prov, 'Costo de Envío', 'Costo de Envío Cliente']:
            if col in cruce.columns:
                cruce[col] = pd.to_numeric(cruce[col].astype(str).str.replace(',','.'), errors='coerce').fillna(0)
            else:
                cruce[col] = 0

        # Lógica 1: Auditoría a Flexit (Proveedor)
        def auditar_flexit(row):
            cobro = row[col_precio_prov]
            cp = row.get('CP_Num', 0)
            
            if row['_merge'] == 'left_only':
                return 'Fantasma (Facturado pero No en Sistema)'
            elif row['_merge'] == 'right_only':
                return 'Omitido (No facturado por Flexit)'
                
            es_caba = 1000 <= cp <= 1499
            es_tarifa_oficial = any(abs(cobro - tarifa) <= 10 for tarifa in TARIFAS_VALIDAS)
            
            if es_caba:
                if (cobro - TARIFA_CABA) > 50:
                    return 'Reclamo: Sobreprecio CABA'
                return 'Cobro OK (CABA)'
            else:
                if es_tarifa_oficial:
                    return 'Cobro OK (Tarifa Oficial)'
                else:
                    return 'Reclamo: Tarifa Irregular / Inventada'
                    
        cruce['Estado_Flexit'] = cruce.apply(auditar_flexit, axis=1)

        # Lógica 2: Auditoría del Sistema Interno (Zonas Faltantes)
        def auditar_sistema(row):
            if row['_merge'] == 'right_only':
                return 'N/A (No facturado)'
            if row['Costo de Envío'] == 0:
                return 'Falta Zona en Sistema ($0)'
            return 'Cotizado OK'
            
        cruce['Alerta_Sistema_Interno'] = cruce.apply(auditar_sistema, axis=1)

        # Lógica 3: Cálculo de Reclamos
        def calcular_reclamo(row):
            estado = row['Estado_Flexit']
            cobro = row[col_precio_prov]
            costo_sistema = row['Costo de Envío']
            
            if estado == 'Fantasma (Facturado pero No en Sistema)':
                return cobro
            elif estado == 'Reclamo: Sobreprecio CABA':
                return round(cobro - TARIFA_CABA, 2)
            elif estado == 'Reclamo: Tarifa Irregular / Inventada':
                if costo_sistema > 0:
                    return round(cobro - costo_sistema, 2)
                else:
                    # Si no hay costo en sistema, y la tarifa es inventada (ej. $15,000), 
                    # reclamamos lo que exceda la tarifa máxima oficial como red de seguridad.
                    return round(max(0, cobro - TARIFA_GBA2), 2)
            return 0.0

        cruce['Monto_a_Reclamar'] = cruce.apply(calcular_reclamo, axis=1)

        # Ordenar columnas para el reporte final
        columnas_deseadas = ['Número Tracking', 'ID venta ML', 'Nro Pedido', 
                             'Estado_Flexit', 'Alerta_Sistema_Interno', 'Monto_a_Reclamar', 
                             col_precio_prov, 'Costo de Envío', 'Costo de Envío Cliente', 
                             'Fecha Venta', 'Localidad', 'CP', 'Provincia']
        columnas_existentes = [col for col in columnas_deseadas if col in cruce.columns]
        
        reporte_final = cruce[columnas_existentes].copy().fillna('N/A')
        
        # Filtros de Pestañas
        df_reclamos = reporte_final[reporte_final['Monto_a_Reclamar'] > 0].copy()
        
        df_zonas = cruce[cruce['Alerta_Sistema_Interno'] == 'Falta Zona en Sistema ($0)'].copy()
        if not df_zonas.empty and all(col in df_zonas.columns for col in ['Provincia', 'Localidad', 'CP']):
            resumen_zonas = df_zonas.groupby(['Provincia', 'Localidad', 'CP']).size().reset_index(name='Veces_No_Cotizado')
            resumen_zonas = resumen_zonas.sort_values(by='Veces_No_Cotizado', ascending=False)
        else:
            resumen_zonas = pd.DataFrame(columns=['Provincia', 'Localidad', 'CP', 'Veces_No_Cotizado'])

        # --- MÉTRICAS EN PANTALLA ---
        total_prov = cruce[cruce['_merge'] != 'right_only'][col_precio_prov].sum()
        total_reclamos = df_reclamos['Monto_a_Reclamar'].sum() if not df_reclamos.empty else 0
        viajes_sin_zona = len(df_zonas)
        
        st.success("¡Auditoría completada exitosamente!")
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Facturado (Flexit)", f"$ {total_prov:,.2f}")
        m2.metric("Total a Reclamar", f"$ {total_reclamos:,.2f}", delta="Fantasmas y Sobreprecios", delta_color="inverse")
        m3.metric("Fallas del Sistema (Costo $0)", f"{viajes_sin_zona} envíos")
        
        # --- GENERACIÓN DEL EXCEL ---
        wb = openpyxl.Workbook()
        ws_dash = wb.active
        ws_dash.title = "Dashboard"
        ws_data = wb.create_sheet(title="Auditoria General")
        ws_reclamos = wb.create_sheet(title="Reclamos a Flexit") 
        ws_zonas = wb.create_sheet(title="Zonas Faltantes Internas")     
        
        header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
        
        def dar_formato_tabla(ws, dataframe):
            if dataframe.empty:
                ws.append(["No hay registros para mostrar."])
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
                    if any(moneda in col_name for moneda in ['Precio', 'Costo', 'Monto']):
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '$#,##0.00'
            for i, col in enumerate(dataframe.columns):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = 22

        dar_formato_tabla(ws_data, reporte_final)
        dar_formato_tabla(ws_reclamos, df_reclamos)
        dar_formato_tabla(ws_zonas, resumen_zonas)
        
        ws_dash.sheet_view.showGridLines = False
        ws_dash['B2'] = "Dashboard de Control - Flexit"
        ws_dash['B2'].font = Font(size=16, bold=True, color="2F4F4F")
        ws_dash['B4'] = "1. Total Facturado por Proveedor:"
        ws_dash['C4'] = total_prov
        ws_dash['B5'] = "2. Dinero Total a RECLAMAR:"
        ws_dash['C5'] = total_reclamos
        ws_dash['C5'].font = Font(color="B22222", bold=True)
        ws_dash['B6'] = "3. Viajes con Zona sin cargar ($0):"
        ws_dash['C6'] = viajes_sin_zona
        
        for r in range(4, 7):
            ws_dash[f'B{r}'].font = Font(bold=True)
            if r != 6: ws_dash[f'C{r}'].number_format = '$#,##0.00'
            
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        st.subheader("📥 Descargar Reporte Final")
        st.download_button(
            label="Descargar Auditoría en Excel",
            data=excel_buffer,
            file_name="Auditoria_Flexit_v3_Definitiva.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
