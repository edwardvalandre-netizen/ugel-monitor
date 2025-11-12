from flask import Blueprint, send_file, request
from utils.informes_excel import generar_informe_excel

visitas_bp = Blueprint('visitas', __name__)

@visitas_bp.route('/exportar_excel')
def exportar_excel():
    mes = request.args.get('mes')  # ej. '2025-11'
    visitas = listar_visitas(mes)  # tu función para obtener visitas
    resumen = {
        "total_visitas": total_visitas(),
        "visitas_mes": visitas_mes(mes),
        "meta_mensual": 30,
        "porcentaje_meta": visitas_mes(mes)/30 if 30 else 0,
    }
    excel_buffer = generar_informe_excel(visitas, resumen, mes)
    nombre = f"informe_visitas_{mes}.xlsx" if mes else "informe_visitas_todos.xlsx"
    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
