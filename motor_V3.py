import pandas as pd
import numpy as np
from google.colab import files
import io
from datetime import datetime

def ejecutar_conciliacion():
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    num_reporte = datetime.now().strftime("%H%M")
    
    print(f"🚀 MOTOR ACTIVO | Reporte: {num_reporte} | Fecha: {fecha_hoy}")
    print("Cargue los archivos de Contabilidad y Banco.")
    
    subidos = files.upload()
    if len(subidos) < 2:
        return print("❌ Error: Se requieren 2 archivos.")

    # Identificación de archivos
    for nombre in subidos.keys():
        if 'contabilidad' in nombre.lower():
            df_cont = pd.read_excel(io.BytesIO(subidos[nombre]))
            df_cont.columns = ['Fecha', 'Referencia', 'Descripcion', 'Monto']
        elif 'banco' in nombre.lower():
            df_banco = pd.read_excel(io.BytesIO(subidos[nombre]))
            df_banco.columns = ['Fecha_Operacion', 'Codigo_Transaccion', 'Detalle', 'Valor']

    # Lógica de cruce (Matching)
    df_cont['Referencia'] = df_cont['Referencia'].astype(str)
    df_banco['Codigo_Transaccion'] = df_banco['Codigo_Transaccion'].astype(str)
    conciliacion = pd.merge(df_cont, df_banco, left_on='Referencia', right_on='Codigo_Transaccion', how='outer')
    
    conciliacion['Estado'] = "Pendiente"
    conciliacion.loc[(conciliacion['Referencia'].notna()) & (conciliacion['Codigo_Transaccion'].notna()), 'Estado'] = 'Conciliado'
    conciliacion.loc[(conciliacion['Referencia'].notna()) & (conciliacion['Codigo_Transaccion'].isna()), 'Estado'] = 'Pendiente en Banco (Cheque/Giro)'
    conciliacion.loc[(conciliacion['Referencia'].isna()) & (conciliacion['Codigo_Transaccion'].notna()), 'Estado'] = 'No registrado en Contabilidad (Comisión/Nota)'

    # Generación de archivo
    nombre_salida = f'Reporte_{fecha_hoy}_ID_{num_reporte}.xlsx'
    with pd.ExcelWriter(nombre_salida, engine='xlsxwriter') as writer:
        conciliacion.to_excel(writer, sheet_name='Detalle', index=False)
        resumen_df = pd.DataFrame({
            'Concepto': ['FECHA', 'ID REPORTE', 'Saldo Libros', 'Saldo Banco'],
            'Valor': [fecha_hoy, num_reporte, df_cont['Monto'].sum(), df_banco['Valor'].sum()]
        })
        resumen_df.to_excel(writer, sheet_name='Resumen', index=False)

    print(f"✅ Reporte generado: {nombre_salida}")
    files.download(nombre_salida)
