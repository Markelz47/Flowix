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

    C_GREEN_DARK   = "00543A"
    C_GREEN_MID    = "007A52"
    C_GREEN_LIGHT  = "E8F5F0"
    C_WHITE        = "FFFFFF"
    C_GRAY_MED     = "DDDDDD"
    C_ORANGE       = "C0392B"
    C_ORANGE_LIGHT = "FDECEA"
    C_DARK         = "1A1A2E"
    C_BLUE_LIGHT   = "EBF3FB"
    C_BLUE_MED     = "D6E8F7"
    C_EJ_BG        = "F8F8F8"
    C_SERIE_EVEN   = "FFFFFF"
    C_SERIE_ODD    = "F5F8FF"

    def fill(h):
        return PatternFill("solid", fgColor=h)

    thin  = Side(style='thin',   color=C_GRAY_MED)
    med_g = Side(style='medium', color=C_GREEN_DARK)
    med_o = Side(style='medium', color=C_ORANGE)
    med_gr= Side(style='medium', color="999999")
    none  = Side(style=None)

    # Column widths
    ws.column_dimensions['A'].width = 36
    ws.column_dimensions['B'].width = 20
    col = 3
    micro_cols = []
    for m in range(MICROS):
        ws.column_dimensions[get_column_letter(col)].width   = 14
        ws.column_dimensions[get_column_letter(col+1)].width = 12
        micro_cols.append((col, col+1, False, m+1))
        col += 2
    ws.column_dimensions[get_column_letter(col)].width   = 16
    ws.column_dimensions[get_column_letter(col+1)].width = 12
    micro_cols.append((col, col+1, True, 7))
    total_cols = col + 1

    row = 1

    # ── TITLE ──
    ws.row_dimensions[row].height = 28
    c = ws.cell(row=row, column=1, value="MESOCICLO")
    c.font = Font(bold=True, color=C_WHITE, size=14, name="Calibri")
    c.fill = fill(C_GREEN_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border = Border(left=med_g, top=med_g, bottom=med_g)
    for ci in range(2, total_cols+1):
        cell = ws.cell(row=row, column=ci)
        cell.fill = fill(C_GREEN_DARK)
        cell.border = Border(top=med_g, bottom=med_g,
                             right=med_g if ci==total_cols else none)
    cc = total_cols - 2
    ws.cell(row=row, column=cc).value = f"Cliente: {r.get('cliente','')}"
    ws.cell(row=row, column=cc).font = Font(bold=True, color=C_WHITE, size=10, name="Calibri")
    ws.cell(row=row, column=cc).alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(start_row=row, start_column=cc, end_row=row, end_column=total_cols)
    row += 1

    # ── HEADER ROW ──
    ws.row_dimensions[row].height = 18
    for ci, label in [(1,"EJERCICIO"), (2,"SERIE / REPS")]:
        c = ws.cell(row=row, column=ci, value=label)
        c.font = Font(bold=True, color=C_WHITE, size=8, name="Calibri")
        c.fill = fill(C_GREEN_DARK)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(left=med_g if ci==1 else med_gr,
                          bottom=med_g, right=med_gr)
    for sc, kc, is_desc, num in micro_cols:
        bg = C_ORANGE if is_desc else C_GREEN_DARK
        label = f"MICROCICLO {num}" if not is_desc else "MC 7 DESCARGA"
        cs = ws.cell(row=row, column=sc, value=label)
        cs.font = Font(bold=True, color=C_WHITE, size=7.5, name="Calibri")
        cs.fill = fill(bg)
        cs.alignment = Alignment(horizontal="center", vertical="center")
        cs.border = Border(left=med_gr,
                           bottom=med_g if not is_desc else med_o)
        ck = ws.cell(row=row, column=kc, value="KG / REPS")
        ck.font = Font(bold=True, color=C_WHITE, size=7.5, name="Calibri")
        ck.fill = fill(bg)
        ck.alignment = Alignment(horizontal="center", vertical="center")
        ck.border = Border(right=med_o if is_desc else med_gr,
                           bottom=med_g if not is_desc else med_o)
    row += 1

    # ── DAYS ──
    for dia in r["dias"]:
        # Day header
        ws.row_dimensions[row].height = 20
        c = ws.cell(row=row, column=1,
                    value=f"  ENTRENAMIENTO DÍA {dia['numero']}  —  {dia['nombre']}")
        c.font = Font(bold=True, color=C_WHITE, size=9, name="Calibri")
        c.fill = fill(C_GREEN_MID)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = Border(left=med_g, top=med_g, bottom=med_g)
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=2)
        for sc, kc, is_desc, num in micro_cols:
            bg = C_ORANGE if is_desc else C_GREEN_MID
            for ci in [sc, kc]:
                cell = ws.cell(row=row, column=ci)
                cell.fill = fill(bg)
                cell.border = Border(
                    top=med_g, bottom=med_g,
                    left=med_gr if ci==sc else none,
                    right=med_o if (is_desc and ci==kc) else
                          (med_gr if ci==kc else none))
        row += 1

        # Exercises
        for ej in dia["ejercicios"]:
            series_list = ej.get("series", [])
            if not isinstance(series_list, list):
                series_list = [{"tipo": "", "series": str(series_list),
                                "reps": ej.get("reps",""),
                                "rir":  ej.get("rir","0")}]
            num_series = len(series_list)

            for si, s in enumerate(series_list):
                ws.row_dimensions[row].height = 16
                is_first = si == 0

                # Col A — ejercicio name (merged across all its series rows)
                if is_first:
                    c = ws.cell(row=row, column=1, value=ej["nombre"])
                    c.font = Font(bold=True, color=C_DARK, size=8.5, name="Calibri")
                    c.fill = fill(C_EJ_BG)
                    c.alignment = Alignment(horizontal="left", vertical="center",
                                            indent=1, wrap_text=False)
                    c.border = Border(left=med_g, top=thin,
                                      bottom=thin, right=med_gr)
                    if num_series > 1:
                        ws.merge_cells(start_row=row, start_column=1,
                                       end_row=row+num_series-1, end_column=1)

                # Col B — serie type + reps
                tipo  = s.get("tipo", "") or ("Serie única" if num_series==1 else f"Serie {si+1}")
                reps  = s.get("reps", "")
                rir   = s.get("rir", "0")
                num_s = s.get("series", "")
                bg_s  = C_SERIE_EVEN if si % 2 == 0 else C_SERIE_ODD

                cb = ws.cell(row=row, column=2,
                             value=f"{tipo}   {num_s}x{reps}")
                cb.font = Font(color="444444", size=8, name="Calibri",
                               italic=bool(s.get("tipo","")))
                cb.fill = fill(bg_s)
                cb.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                cb.border = Border(left=med_gr, top=thin, bottom=thin, right=med_gr)

                # Microciclo columns
                rir_i = int(rir) if str(rir).isdigit() else 0
                for i, (sc, kc, is_desc, num) in enumerate(micro_cols):
                    if is_desc:
                        rir_m = min(rir_i+1, 3); bg_m = C_ORANGE_LIGHT
                    elif i >= MICROS-2:
                        rir_m = max(rir_i-1, 0); bg_m = C_BLUE_MED
                    else:
                        rir_m = rir_i; bg_m = C_BLUE_LIGHT if i%2==0 else C_WHITE

                    cs = ws.cell(row=row, column=sc, value=f"Rir{rir_m}")
                    cs.font = Font(color="333333", size=8, name="Calibri")
                    cs.fill = fill(bg_m)
                    cs.alignment = Alignment(horizontal="center", vertical="center")
                    cs.border = Border(left=med_gr, top=thin, bottom=thin)

                    ck = ws.cell(row=row, column=kc, value="")
                    ck.fill = fill(C_WHITE if not is_desc else "FFF5F5")
                    ck.border = Border(bottom=thin,
                                       right=med_o if is_desc else med_gr)
                row += 1

            # Thin spacer between exercises
            ws.row_dimensions[row].height = 3
            for ci in range(1, total_cols+1):
                ws.cell(row=row, column=ci).fill = fill("EEEEEE")
            row += 1

        # Spacer between days
        ws.row_dimensions[row].height = 6
        for ci in range(1, total_cols+1):
            ws.cell(row=row, column=ci).fill = fill("CCCCCC")
        row += 1

    # ── COACHING NOTES ──
    if r.get("notas_coaching"):
        ws.row_dimensions[row].height = 18
        c = ws.cell(row=row, column=1, value="  NOTAS DE COACHING")
        c.font = Font(bold=True, color=C_WHITE, size=9, name="Calibri")
        c.fill = fill(C_GREEN_DARK)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = Border(left=med_g, top=med_g, bottom=med_g)
        for ci in range(2, total_cols+1):
            cell = ws.cell(row=row, column=ci)
            cell.fill = fill(C_GREEN_DARK)
            cell.border = Border(top=med_g, bottom=med_g,
                                 right=med_g if ci==total_cols else none)
        row += 1
        ws.row_dimensions[row].height = 55
        c = ws.cell(row=row, column=1, value=r["notas_coaching"])
        c.font = Font(color=C_DARK, size=8.5, name="Calibri", italic=True)
        c.fill = fill(C_GREEN_LIGHT)
        c.alignment = Alignment(horizontal="left", vertical="top",
                                wrap_text=True, indent=1)
        c.border = Border(left=med_g, bottom=med_g, right=med_g)
        ws.merge_cells(start_row=row, start_column=1,
                       end_row=row, end_column=total_cols)

    ws.freeze_panes = "C3"

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
