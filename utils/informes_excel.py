# utils/informes_excel.py
import io
import pandas as pd

def generar_informe_excel(visitas, resumen, mes):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="xlsxwriter")

    # Hoja con visitas
    df_visitas = pd.DataFrame(visitas)
    df_visitas.to_excel(writer, sheet_name="Visitas", index=False)

    # Hoja con resumen
    df_resumen = pd.DataFrame([resumen])
    df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    writer.close()
    output.seek(0)
    return output

def exportar_excel_visitas(visitas, resumen, mes):
    return generar_informe_excel(visitas, resumen, mes)

