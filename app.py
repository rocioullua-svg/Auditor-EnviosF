import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuración de la página web
st.set_page_config(page_title="Auditoría Flexit", page_icon="📊", layout="wide")

st.title("📊 Auditoría Cruzada de Envíos - Flexit")
st.markdown("Sube el archivo del proveedor y tu reporte interno para validar los cobros automáticamente.")

def cargar_archivo_inteligente(file_uploader, saltar_filas=0):
    """Función a prueba de balas para leer archivos subidos a la web"""
    nombre_archivo = file_uploader.name
    archivo_bytes = file_uploader.read()
    
    if nombre_archivo.lower().endswith(('.xls', '.xlsx')):
        return pd.read_excel(io.BytesIO(archivo_bytes), skiprows=saltar_filas)
    
    separadores = [',', ';', '\t']
    codificaciones = ['utf-8', 'latin-1', 'cp1252']
    
    for sep in separadores:
        for enc in codificaciones:
            try:
                df = pd.read_csv(io.BytesIO(archivo_bytes), skiprows=saltar_filas, 
                                 encoding=enc, sep=sep, on_bad_lines='skip', engine='python')
                if len(df.columns) > 5: 
                    return df
            except Exception:
                continue
                
    return pd.read_csv(io.BytesIO(archivo_bytes), skiprows=saltar_filas, encoding='latin-1', on_bad_lines='skip')

# --- INTERFAZ DE USUARIO (Columnas de Carga) ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📁 1. Archivo del Proveedor")
    archivo_prov = st.file_uploader("Sube el reporte que te envía Flexit", type=['csv', 'xls', 'xlsx'], key="prov")

with col2:
    st.subheader("📁 2. Archivo Interno")
    archivo_int = st.file_uploader("Sube tu reporte de sistema", type=['csv', 'xls', 'xlsx'], key="int")

# --- PROCESAMIENTO ---
if archivo_prov and archivo_int:
    with st.spinner("⚙️ Procesando datos y cruzando información..."):
        
        # 1. Cargar datos
        df_prov = cargar_archivo_inteligente(archivo_prov, saltar_filas=4)
        df_int = cargar_archivo_inteligente(archivo_int, saltar_filas=0)
        
        df_prov.columns = df_prov.columns.str.strip()
        df_int.columns = df_int.columns.str.strip()
        
        # 2. Limpieza básica
        if 'Número Tracking' in df_prov.columns:
            df_prov['Número Tracking'] = df_prov['Número Tracking'].astype(str).str.strip().str.upper()
        if 'ID venta ML' in df_prov.columns:
            df_prov['ID venta ML'] = df_prov['ID venta ML'].astype(str).str.strip().str.replace('.0', '', regex=False)
        if 'Nro Pedido' in df_int.columns:
            df_int['Nro Pedido'] = df_int['Nro Pedido'].astype(str).str.strip().str.replace('.0', '', regex=False)
            
        if 'Pedido - Transportista' in df_int.columns:
            df_int_flexit = df_int[df_int['Pedido - Transportista'].astype(str).str.contains('FLEX', case=False, na=False)]
        else:
            df_int_flexit = df_int.copy()
            
        # 3. Cruce Inteligente
        columna_tracking_interno = 'tracking code'
        if columna_tracking_interno in df_int_flexit.columns:
            df_int_flexit[columna_tracking_interno] = df_int_flexit[columna_tracking_interno].astype(str).str.strip().str.upper()
            cruce = pd.merge(df_prov, df_int_flexit, left_on='Número Tracking', right_on=columna_tracking_interno, how='outer', indicator=True)
            metodo_cruce = "Número Tracking"
        else:
            cruce = pd.merge(df_prov, df_int_flexit, left_on='ID venta ML', right_on='Nro Pedido', how='outer', indicator=True)
            metodo_cruce = "ID venta ML / Nro Pedido (Plan B)"
            
        # 4. Cálculos Monetarios
        if 'Precio' in cruce.columns and 'Costo de Envío' in cruce.columns:
            cruce['Precio'] = pd.to_numeric(cruce['Precio'].astype(str).str.replace(',','.'), errors='coerce').fillna(0)
            cruce['Costo de Envío'] = pd.to_numeric(cruce['Costo de Envío'].astype(str).str.replace(',','.'), errors='coerce').fillna(0)
            cruce['Diferencia_Plata'] = (cruce['Precio'] - cruce['Costo de Envío']).round(2)
        else:
            cruce['Diferencia_Plata'] = 0
            
        # 5. Clasificación
        def clasificar_estado(row):
            if row['_merge'] == 'left_only':
                return 'Cobrado pero No en Sistema'
            elif row['_merge'] == 'right_only':
                return 'En Sistema pero No Cobrado'
            elif abs(row['Diferencia_Plata']) <= 50:
                return 'Coincide OK'
            elif row['Diferencia_Plata'] > 50:
                return 'Cobro MAYOR al Sistema'
            else:
                return 'Cobro MENOR al Sistema'
                
        cruce['Estado_Auditoria'] = cruce.apply(clasificar_estado, axis=1)
        
        columnas_deseadas = ['Número Tracking', 'ID venta ML', 'Nro Pedido', 'Estado_Auditoria', 'Precio', 'Costo de Envío', 'Diferencia_Plata', 'Fecha Venta', 'Localidad', 'Dirección de entrega']
        columnas_existentes = [col for col in columnas_deseadas if col in cruce.columns]
        
        reporte_final = cruce[columnas_existentes].copy()
        reporte_final.fillna('N/A', inplace=True)
        
        # --- MÉTRICAS EN PANTALLA ---
        st.success(f"¡Cruce finalizado con éxito usando el método: {metodo_cruce}!")
        
        total_prov = cruce['Precio'].sum() if 'Precio' in cruce.columns else 0
        total_sist = cruce['Costo de Envío'].sum() if 'Costo de Envío' in cruce.columns else 0
        dif_total = total_prov - total_sist
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Facturado Proveedor", f"$ {total_prov:,.2f}")
        m2.metric("Total Estimado Sistema", f"$ {total_sist:,.2f}")
        m3.metric("Diferencia Neta (A Reclamar)", f"$ {dif_total:,.2f}", delta=f"{dif_total:,.2f}", delta_color="inverse")
        
        # --- RESUMEN DE ESTADOS ---
        st.subheader("📊 Resumen de la Auditoría")
        resumen = reporte_final['Estado_Auditoria'].value_counts()
        st.bar_chart(resumen)
        
        # --- GENERACIÓN DEL EXCEL PARA DESCARGA ---
        wb = openpyxl.Workbook()
        ws_data = wb.active
        ws_data.title = "Auditoria Detalle"
        ws_dash = wb.create_sheet(title="Dashboard", index=0)
        
        for r in dataframe_to_rows(reporte_final, index=False, header=True):
            ws_data.append(r)
            
        header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'),
                             top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
                             
        for col in ws_data.iter_cols(min_row=1, max_row=1, max_col=ws_data.max_column):
            for cell in col:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        for row in ws_data.iter_rows(min_row=2, max_col=ws_data.max_column, max_row=ws_data.max_row):
            for cell in row:
                cell.border = thin_border
                col_name = str(ws_data.cell(row=1, column=cell.column).value)
                if "Precio" in col_name or "Costo" in col_name or "Diferencia" in col_name:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '$#,##0.00'
                        
        for i, col in enumerate(columnas_existentes):
            ws_data.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = 22
            
        ws_dash.sheet_view.showGridLines = False
        ws_dash['B2'] = "Dashboard de Auditoría - Control Proveedores"
        ws_dash['B2'].font = Font(size=18, bold=True, color="2F4F4F")
        
        ws_dash['B4'] = "Total Facturado Proveedor:"
        ws_dash['C4'] = total_prov
        ws_dash['C4'].number_format = '$#,##0.00'
        ws_dash['B5'] = "Total Estimado Sistema:"
        ws_dash['C5'] = total_sist
        ws_dash['C5'].number_format = '$#,##0.00'
        ws_dash['B6'] = "Diferencia Neta (A reclamar):"
        ws_dash['C6'] = dif_total
        ws_dash['C6'].number_format = '$#,##0.00'
        
        for r in range(4, 7):
            ws_dash[f'B{r}'].font = Font(bold=True)
            ws_dash[f'C{r}'].font = Font(bold=True, color="B22222" if dif_total > 0 and r==6 else "000000")
            
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        st.subheader("📥 Descargar Reporte")
        st.download_button(
            label="Descargar Reporte de Auditoría en Excel",
            data=excel_buffer,
            file_name="Auditoria_Flexit_Cruzada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.subheader("👀 Vista previa de los datos")
        st.dataframe(reporte_final)
