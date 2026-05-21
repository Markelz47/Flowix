from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

app = Flask(__name__)
CORS(app)

def build_excel(r):
    wb = Workbook()
    ws = wb.active
    ws.title = "MESOCICLO"

    MICROS = 6

    C_GREEN_DARK  = "00543A"
    C_GREEN_MID   = "007A52"
    C_GREEN_LIGHT = "E8F5F0"
    C_WHITE       = "FFFFFF"
    C_GRAY_LIGHT  = "F7F7F7"
    C_GRAY_MED    = "DDDDDD"
    C_ORANGE      = "C0392B"
    C_ORANGE_LIGHT= "FDECEA"
    C_DARK        = "1A1A2E"
    C_BLUE_LIGHT  = "EBF3FB"
    C_BLUE_MED    = "D6E8F7"

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    thin_side  = Side(style='thin',   color=C_GRAY_MED)
    med_green  = Side(style='medium', color=C_GREEN_DARK)
    med_orange = Side(style='medium', color=C_ORANGE)
    med_gray   = Side(style='medium', color="999999")
    none_side  = Side(style=None)

    ws.column_dimensions['A'].width = 52
    col = 2
    micro_cols = []
    for m in range(MICROS):
        ws.column_dimensions[get_column_letter(col)].width   = 20
        ws.column_dimensions[get_column_letter(col+1)].width = 13
        micro_cols.append((col, col+1, False, m+1))
        col += 2
    ws.column_dimensions[get_column_letter(col)].width   = 22
    ws.column_dimensions[get_column_letter(col+1)].width = 13
    micro_cols.append((col, col+1, True, 7))
    total_cols = col + 1

    row = 1

    # TITLE ROW
    ws.row_dimensions[row].height = 30
    c = ws.cell(row=row, column=1, value="MESOCICLO")
    c.font      = Font(bold=True, color=C_WHITE, size=15, name="Calibri")
    c.fill      = fill(C_GREEN_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border    = Border(left=med_green, top=med_green, bottom=med_green)
    for col_i in range(2, total_cols+1):
        cell = ws.cell(row=row, column=col_i)
        cell.fill   = fill(C_GREEN_DARK)
        cell.border = Border(top=med_green, bottom=med_green,
                             right=med_green if col_i==total_cols else none_side)
    cliente_col_start = total_cols - 3
    ws.cell(row=row, column=cliente_col_start).value = f"Cliente: {r.get('cliente','')}"
    ws.cell(row=row, column=cliente_col_start).font = Font(bold=True, color=C_WHITE, size=10, name="Calibri")
    ws.cell(row=row, column=cliente_col_start).alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(start_row=row, start_column=cliente_col_start, end_row=row, end_column=total_cols)
    row += 1

    # MICROCICLO HEADERS
    ws.row_dimensions[row].height = 20
    c = ws.cell(row=row, column=1, value="EJERCICIO")
    c.font      = Font(bold=True, color=C_WHITE, size=9, name="Calibri")
    c.fill      = fill(C_GREEN_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = Border(left=med_green, bottom=med_green, right=med_gray)
    for sc, kc, is_desc, num in micro_cols:
        label = f"MICROCICLO  {num}" if not is_desc else "MICROCICLO 7 ( DESCARGA )"
        bg    = C_ORANGE if is_desc else C_GREEN_DARK
        cs = ws.cell(row=row, column=sc, value=label)
        cs.font      = Font(bold=True, color=C_WHITE, size=8, name="Calibri")
        cs.fill      = fill(bg)
        cs.alignment = Alignment(horizontal="center", vertical="center")
        cs.border    = Border(left=med_gray, bottom=med_green if not is_desc else med_orange)
        ck = ws.cell(row=row, column=kc, value="KG / REPES")
        ck.font      = Font(bold=True, color=C_WHITE, size=8, name="Calibri")
        ck.fill      = fill(bg)
        ck.alignment = Alignment(horizontal="center", vertical="center")
        ck.border    = Border(right=med_orange if is_desc else med_gray,
                              bottom=med_green if not is_desc else med_orange)
    row += 1

    # DAYS
    for dia in r["dias"]:
        ws.row_dimensions[row].height = 22
        day_label = f"  ENTRENAMIENTO DIA. {dia['numero']}          {dia['nombre']}"
        c = ws.cell(row=row, column=1, value=day_label)
        c.font      = Font(bold=True, color=C_WHITE, size=9, name="Calibri")
        c.fill      = fill(C_GREEN_MID)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = Border(left=med_green, top=med_green, bottom=med_green, right=med_gray)
        for sc, kc, is_desc, num in micro_cols:
            label = f"MICROCICLO  {num}" if not is_desc else "MICROCICLO 7 ( DESCARGA )"
            bg    = C_ORANGE if is_desc else C_GREEN_MID
            cs = ws.cell(row=row, column=sc, value=label)
            cs.font      = Font(bold=True, color=C_WHITE, size=8, name="Calibri")
            cs.fill      = fill(bg)
            cs.alignment = Alignment(horizontal="center", vertical="center")
            cs.border    = Border(left=med_gray, top=med_green, bottom=med_green)
            ck = ws.cell(row=row, column=kc, value="KG / REPES")
            ck.font      = Font(bold=True, color=C_WHITE, size=8, name="Calibri")
            ck.fill      = fill(bg)
            ck.alignment = Alignment(horizontal="center", vertical="center")
            ck.border    = Border(top=med_green, bottom=med_green,
                                  right=med_orange if is_desc else med_gray)
        row += 1

        for ei, ej in enumerate(dia["ejercicios"]):
            ws.row_dimensions[row].height = 18
            bg_ej = C_WHITE if ei % 2 == 0 else C_GRAY_LIGHT
            c = ws.cell(row=row, column=1, value=ej["nombre"])
            c.font      = Font(color=C_DARK, size=9, name="Calibri")
            c.fill      = fill(bg_ej)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False, indent=1)
            c.border    = Border(left=med_green, bottom=thin_side, right=med_gray)
            # Build series label from array or legacy single value
            series_arr = ej.get("series", [])
            if isinstance(series_arr, list):
                parts = []
                for s in series_arr:
                    tipo = s.get("tipo","")
                    label = f"{tipo} {s.get('series','')}x? Rir{s.get('rir','0')}" if tipo else f"{s.get('series','')}x? Rir{s.get('rir','0')}"
                    parts.append(label.strip())
                series_label = " / ".join(parts)
                rir_val = int(series_arr[0].get("rir","0")) if series_arr else 0
            else:
                rir_val = int(ej.get("rir","0")) if str(ej.get("rir","0")).isdigit() else 0
                series_label = f"{series_arr}x? Rir{rir_val}"

            for i, (sc, kc, is_desc, num) in enumerate(micro_cols):
                if is_desc:
                    # Descarga: increase all RIR by 1
                    if isinstance(ej.get("series",[]), list):
                        desc_parts = []
                        for s in ej["series"]:
                            tipo = s.get("tipo","")
                            r_desc = min(int(s.get("rir","0"))+1, 3)
                            lbl = f"{tipo} {s.get('series','')}x? Rir{r_desc}" if tipo else f"{s.get('series','')}x? Rir{r_desc}"
                            desc_parts.append(lbl.strip())
                        cell_val = " / ".join(desc_parts)
                    else:
                        cell_val = f"{series_arr}x? Rir{min(rir_val+1,3)}"
                    bg_m = C_ORANGE_LIGHT
                elif i >= MICROS - 2:
                    if isinstance(ej.get("series",[]), list):
                        prog_parts = []
                        for s in ej["series"]:
                            tipo = s.get("tipo","")
                            r_prog = max(int(s.get("rir","0"))-1, 0)
                            lbl = f"{tipo} {s.get('series','')}x? Rir{r_prog}" if tipo else f"{s.get('series','')}x? Rir{r_prog}"
                            prog_parts.append(lbl.strip())
                        cell_val = " / ".join(prog_parts)
                    else:
                        cell_val = f"{series_arr}x? Rir{max(rir_val-1,0)}"
                    bg_m = C_BLUE_MED
                else:
                    cell_val = series_label
                    bg_m = C_BLUE_LIGHT if i % 2 == 0 else C_WHITE
                cs = ws.cell(row=row, column=sc, value=cell_val)
                cs.font      = Font(color="333333", size=9, name="Calibri")
                cs.fill      = fill(bg_m)
                cs.alignment = Alignment(horizontal="center", vertical="center")
                cs.border    = Border(left=med_gray, bottom=thin_side)
                ck = ws.cell(row=row, column=kc, value="")
                ck.fill   = fill(C_WHITE if not is_desc else "FFF5F5")
                ck.border = Border(bottom=thin_side, right=med_orange if is_desc else med_gray)
            row += 1

        ws.row_dimensions[row].height = 5
        for c_i in range(1, total_cols+1):
            cell = ws.cell(row=row, column=c_i)
            cell.fill   = fill("DDDDDD")
            cell.border = Border()
        row += 1

    # COACHING NOTES
    if r.get("notas_coaching"):
        ws.row_dimensions[row].height = 20
        c = ws.cell(row=row, column=1, value="  NOTAS DE COACHING")
        c.font      = Font(bold=True, color=C_WHITE, size=9, name="Calibri")
        c.fill      = fill(C_GREEN_DARK)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = Border(left=med_green, top=med_green, bottom=med_green)
        for c_i in range(2, total_cols+1):
            cell = ws.cell(row=row, column=c_i)
            cell.fill   = fill(C_GREEN_DARK)
            cell.border = Border(top=med_green, bottom=med_green,
                                 right=med_green if c_i==total_cols else none_side)
        row += 1
        ws.row_dimensions[row].height = 55
        c = ws.cell(row=row, column=1, value=r["notas_coaching"])
        c.font      = Font(color=C_DARK, size=9, name="Calibri", italic=True)
        c.fill      = fill(C_GREEN_LIGHT)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        c.border    = Border(left=med_green, bottom=med_green, right=med_green)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=total_cols)

    ws.freeze_panes = "B3"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/', methods=['GET'])
def health():
    return jsonify({"status": "ok", "service": "Flowix Excel Generator"})

@app.route('/excel', methods=['POST', 'OPTIONS'])
def generar_excel():
    if request.method == 'OPTIONS':
        return '', 204
    try:
        data = request.get_json()
        if not data or 'dias' not in data:
            return jsonify({"error": "JSON inválido"}), 400
        excel_file = build_excel(data)
        nombre = (data.get('cliente') or 'cliente').replace(' ', '_')
        return send_file(
            excel_file,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"Rutina_{nombre}_Flowix.xlsx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
