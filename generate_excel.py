import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime
import shutil

def create_amortization_sheet():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabla de Amortización"

    # Completamente desbloqueado / sin contraseñas
    ws.protection.sheet = False
    ws.protection.disable()

    font_family = "Calibri"
    
    # Paleta de colores profesionales
    c_navy = "1F4E78"        # Azul oscuro institucional
    c_accent_blue = "2E75B6" # Azul medio encabezados
    c_light_blue = "D9E1F2"  # Azul claro para totales
    c_soft_gray = "F2F2F2"   # Gris suave
    c_zebra = "F9FAFB"       # Fondo alternado
    c_border = "D0D7DE"      # Bordes
    c_green_light = "E2EFDA" # Verde suave resumen
    c_input_bg = "FFF2CC"    # Amarillo suave para celdas editables

    title_font = Font(name=font_family, size=15, bold=True, color="FFFFFF")
    subtitle_font = Font(name=font_family, size=10, italic=True, color="D9E1F2")
    section_font = Font(name=font_family, size=11, bold=True, color="1F4E78")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    bold_font = Font(name=font_family, size=10, bold=True)
    normal_font = Font(name=font_family, size=10)
    note_font = Font(name=font_family, size=9, italic=True, color="595959")

    thin_border_side = Side(style='thin', color=c_border)
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom_side = Side(style='medium', color="1F4E78")
    header_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thick_bottom_side)

    title_fill = PatternFill(start_color=c_navy, end_color=c_navy, fill_type="solid")
    header_fill = PatternFill(start_color=c_accent_blue, end_color=c_accent_blue, fill_type="solid")
    input_fill = PatternFill(start_color=c_input_bg, end_color=c_input_bg, fill_type="solid")
    calc_fill = PatternFill(start_color=c_soft_gray, end_color=c_soft_gray, fill_type="solid")
    summary_fill = PatternFill(start_color=c_green_light, end_color=c_green_light, fill_type="solid")
    zebra_fill = PatternFill(start_color=c_zebra, end_color=c_zebra, fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    # Encabezado Principal en Español
    ws.merge_cells("A1:G1")
    ws["A1"] = "TABLA DE AMORTIZACIÓN DE CRÉDITO (CUOTA FIJA)"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:G2")
    ws["A2"] = "Proyección para 180 Meses - 100% Desbloqueada y Modificable"
    ws["A2"].font = subtitle_font
    ws["A2"].fill = title_fill
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # Sección de Parámetros Modificables
    ws.merge_cells("A4:C4")
    ws["A4"] = "DATOS DEL PRÉSTAMO (MODIFICABLES)"
    ws["A4"].font = section_font
    
    ws.merge_cells("E4:G4")
    ws["E4"] = "RESUMEN DEL CRÉDITO"
    ws["E4"].font = section_font

    inputs = [
        ("A5", "Capital Inicial (Monto Solicitado):", "C5", 50000000, '"$"#,##0.00'),
        ("A6", "Tasa de Interés Nominal Anual (TNA):", "C6", 0.12, '0.00%'),
        ("A7", "Tasa de Interés Mensual:", "C7", "=C6/12", '0.0000%'),
        ("A8", "Plazo Total (en Meses):", "C8", 180, '#,##0'),
        ("A9", "Fecha Inicial (Primer Desembolso):", "C9", datetime.date(2026, 10, 1), 'DD/MM/YYYY'),
    ]

    for label_cell, label_text, val_cell, val, num_fmt in inputs:
        ws[label_cell] = label_text
        ws[label_cell].font = bold_font
        ws[label_cell].alignment = Alignment(horizontal="left", vertical="center")
        
        ws[val_cell] = val
        ws[val_cell].font = bold_font
        ws[val_cell].number_format = num_fmt
        ws[val_cell].alignment = Alignment(horizontal="right", vertical="center")
        ws[val_cell].border = thin_border
        
        # Resaltado amarillo en celdas de entrada
        if label_cell in ["A5", "A6", "A8", "A9"]:
            ws[val_cell].fill = input_fill
        else:
            ws[val_cell].fill = calc_fill

    # Resumen derecho
    summaries = [
        ("E5", "Fecha Final Estimada:", "G5", '=IF(C8>0, EDATE(C9, C8), "")', 'DD/MM/YYYY'),
        ("E6", "Valor de la Cuota Mensual Fija:", "G6", '=IF(AND(C5>0, C8>0, C7>0), PMT(C7, C8, -C5), IF(C8>0, C5/C8, 0))', '"$"#,##0.00'),
        ("E7", "Total de Intereses a Pagar:", "G7", '=SUM(F14:F193)', '"$"#,##0.00'),
        ("E8", "Total a Pagar (Capital + Intereses):", "G8", '=C5+G7', '"$"#,##0.00'),
        ("E9", "Porcentaje Total de Intereses:", "G9", '=IF(C5>0, G7/C5, 0)', '0.00%'),
    ]

    for label_cell, label_text, val_cell, formula, num_fmt in summaries:
        ws[label_cell] = label_text
        ws[label_cell].font = bold_font
        ws[label_cell].alignment = Alignment(horizontal="left", vertical="center")
        
        ws[val_cell] = formula
        ws[val_cell].font = bold_font
        ws[val_cell].number_format = num_fmt
        ws[val_cell].alignment = Alignment(horizontal="right", vertical="center")
        ws[val_cell].border = thin_border
        ws[val_cell].fill = summary_fill

    # Nota de uso
    ws.merge_cells("A11:G11")
    ws["A11"] = "ℹ️ Modifique las celdas amarillas (C5: Capital, C6: Tasa de Interés, C8: Plazo, C9: Fecha Inicial). Todo se recalcula de forma automática y puede grabar el archivo sin restricciones."
    ws["A11"].font = note_font
    ws["A11"].alignment = Alignment(horizontal="left", vertical="center")

    # Títulos de las Columnas en Español
    headers = [
        ("A13", "Número de Cuota"),
        ("B13", "Fecha"),
        ("C13", "Capital Inicial"),
        ("D13", "Cuota Total"),
        ("E13", "Cuota Capital"),
        ("F13", "Intereses"),
        ("G13", "Saldo del Capital")
    ]

    ws.row_dimensions[13].height = 26
    for pos, h_title in headers:
        ws[pos] = h_title
        ws[pos].font = header_font
        ws[pos].fill = header_fill
        ws[pos].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws[pos].border = header_border

    # Generación de las 180 filas (filas 14 a 193)
    total_months = 180
    start_row = 14

    for i in range(1, total_months + 1):
        row = start_row + i - 1
        ws.row_dimensions[row].height = 20
        fill_to_use = zebra_fill if i % 2 == 0 else white_fill

        # Col A: Número de la cuota
        ws[f"A{row}"] = i
        ws[f"A{row}"].font = normal_font
        ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"A{row}"].border = thin_border
        ws[f"A{row}"].fill = fill_to_use

        # Col B: Fecha
        ws[f"B{row}"] = f'=IF(A{row}<=$C$8, EDATE($C$9, A{row}), "")'
        ws[f"B{row}"].font = normal_font
        ws[f"B{row}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"B{row}"].number_format = 'DD/MM/YYYY'
        ws[f"B{row}"].border = thin_border
        ws[f"B{row}"].fill = fill_to_use

        # Col C: Capital Inicial (Mes 1 = C5, Meses siguientes = Saldo del mes anterior G[row-1])
        if i == 1:
            ws[f"C{row}"] = f'=IF(A{row}<=$C$8, $C$5, 0)'
        else:
            ws[f"C{row}"] = f'=IF(A{row}<=$C$8, G{row-1}, 0)'
        ws[f"C{row}"].font = normal_font
        ws[f"C{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"C{row}"].number_format = '"$"#,##0.00'
        ws[f"C{row}"].border = thin_border
        ws[f"C{row}"].fill = fill_to_use

        # Col F: Intereses (Capital Inicial * Tasa Mensual)
        ws[f"F{row}"] = f'=IF(AND(A{row}<=$C$8, C{row}>0), C{row}*$C$7, 0)'
        ws[f"F{row}"].font = normal_font
        ws[f"F{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"F{row}"].number_format = '"$"#,##0.00'
        ws[f"F{row}"].border = thin_border
        ws[f"F{row}"].fill = fill_to_use

        # Col D: Cuota Total (Cuota fija mensual)
        ws[f"D{row}"] = f'=IF(AND(A{row}<=$C$8, C{row}>0), IF(A{row}=$C$8, C{row}+F{row}, $G$6), 0)'
        ws[f"D{row}"].font = normal_font
        ws[f"D{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"D{row}"].number_format = '"$"#,##0.00'
        ws[f"D{row}"].border = thin_border
        ws[f"D{row}"].fill = fill_to_use

        # Col E: Cuota Capital (Amortización = Cuota Total - Intereses)
        ws[f"E{row}"] = f'=IF(AND(A{row}<=$C$8, C{row}>0), MIN(C{row}, D{row}-F{row}), 0)'
        ws[f"E{row}"].font = normal_font
        ws[f"E{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"E{row}"].number_format = '"$"#,##0.00'
        ws[f"E{row}"].border = thin_border
        ws[f"E{row}"].fill = fill_to_use

        # Col G: Saldo del Capital (Capital Inicial - Cuota Capital) -> Pasa como Capital Inicial de la siguiente fila
        ws[f"G{row}"] = f'=IF(AND(A{row}<=$C$8, C{row}>0), MAX(0, C{row}-E{row}), 0)'
        ws[f"G{row}"].font = normal_font
        ws[f"G{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"G{row}"].number_format = '"$"#,##0.00'
        ws[f"G{row}"].border = thin_border
        ws[f"G{row}"].fill = fill_to_use

    # Fila de Totales
    tot_row = start_row + total_months
    ws.row_dimensions[tot_row].height = 24
    ws[f"A{tot_row}"] = "TOTALES"
    ws[f"A{tot_row}"].font = bold_font
    ws[f"A{tot_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{tot_row}"].border = thin_border
    ws[f"A{tot_row}"].fill = PatternFill(start_color=c_light_blue, end_color=c_light_blue, fill_type="solid")

    for col in ["B", "C"]:
        ws[f"{col}{tot_row}"] = ""
        ws[f"{col}{tot_row}"].border = thin_border
        ws[f"{col}{tot_row}"].fill = PatternFill(start_color=c_light_blue, end_color=c_light_blue, fill_type="solid")

    ws[f"D{tot_row}"] = f'=SUM(D14:D{tot_row-1})'
    ws[f"D{tot_row}"].font = bold_font
    ws[f"D{tot_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"D{tot_row}"].number_format = '"$"#,##0.00'
    ws[f"D{tot_row}"].border = thin_border
    ws[f"D{tot_row}"].fill = PatternFill(start_color=c_light_blue, end_color=c_light_blue, fill_type="solid")

    ws[f"E{tot_row}"] = f'=SUM(E14:E{tot_row-1})'
    ws[f"E{tot_row}"].font = bold_font
    ws[f"E{tot_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"E{tot_row}"].number_format = '"$"#,##0.00'
    ws[f"E{tot_row}"].border = thin_border
    ws[f"E{tot_row}"].fill = PatternFill(start_color=c_light_blue, end_color=c_light_blue, fill_type="solid")

    ws[f"F{tot_row}"] = f'=SUM(F14:F{tot_row-1})'
    ws[f"F{tot_row}"].font = bold_font
    ws[f"F{tot_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"F{tot_row}"].number_format = '"$"#,##0.00'
    ws[f"F{tot_row}"].border = thin_border
    ws[f"F{tot_row}"].fill = PatternFill(start_color=c_light_blue, end_color=c_light_blue, fill_type="solid")

    ws[f"G{tot_row}"] = f'=G{tot_row-1}'
    ws[f"G{tot_row}"].font = bold_font
    ws[f"G{tot_row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"G{tot_row}"].number_format = '"$"#,##0.00'
    ws[f"G{tot_row}"].border = thin_border
    ws[f"G{tot_row}"].fill = PatternFill(start_color=c_light_blue, end_color=c_light_blue, fill_type="solid")

    # Anchos de Columna óptimos
    col_widths = {
        "A": 18, # Número de Cuota
        "B": 16, # Fecha
        "C": 22, # Capital Inicial
        "D": 20, # Cuota Total
        "E": 22, # Cuota Capital
        "F": 20, # Intereses
        "G": 22  # Saldo del Capital
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Rutas de guardado
    paths = [
        "/Users/nico/.gemini/antigravity/scratch/tabla_amortizacion/Tabla_Amortizacion_180_Meses.xlsx",
        "/Users/nico/Desktop/Tabla_Amortizacion_180_Meses.xlsx",
        "/Users/nico/Downloads/Tabla_Amortizacion_180_Meses.xlsx"
    ]
    
    # Guardar en la ruta principal y copiar a Desktop y Downloads
    wb.save(paths[0])
    shutil.copyfile(paths[0], paths[1])
    shutil.copyfile(paths[0], paths[2])
    print("Archivos Excel en español generados exitosamente en todas las rutas.")

if __name__ == "__main__":
    create_amortization_sheet()
