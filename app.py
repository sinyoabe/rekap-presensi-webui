from flask import Flask, render_template, request, send_file
from datetime import datetime
from weasyprint import HTML
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import base64
import io

app = Flask(__name__)

NAMA_HARI = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU", "MINGGU"]

def format_tanggal_indo(date_str):
    if not date_str:
        return ""
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        hari = NAMA_HARI[dt.weekday()]
        return f"{hari}, {dt.strftime('%d-%m-%Y')}"
    except ValueError:
        return date_str

DEFAULT_ROWS = [
    {"date_raw": "2026-09-01", "hadir": "06:38:31", "diff_hadir": "+01:04:52", "pulang": "16:05:58", "diff_pulang": "+00:10:58", "status": ""},
    {"date_raw": "2026-09-02", "hadir": "06:52:39", "diff_hadir": "+01:20:37", "pulang": "16:07:26", "diff_pulang": "+00:12:26", "status": ""},
    {"date_raw": "2026-09-03", "hadir": "06:53:26", "diff_hadir": "+01:13:48", "pulang": "16:14:07", "diff_pulang": "+00:19:07", "status": ""},
    {"date_raw": "2026-09-04", "hadir": "06:25:50", "diff_hadir": "+01:17:07", "pulang": "15:44:46", "diff_pulang": "+04:14:46", "status": ""},
    {"date_raw": "2026-09-07", "hadir": "06:45:32", "diff_hadir": "+01:28:16", "pulang": "16:11:02", "diff_pulang": "+00:16:02", "status": ""},
]

for r in DEFAULT_ROWS:
    r["tanggal"] = format_tanggal_indo(r["date_raw"])

DEFAULT_DATA = {
    "instansi_baris1": "Rekap Presensi Pegawai SDN",
    "instansi_baris2": "Sendangmulyo 04 Semarang",
    "bulan_tahun": "Bulan September 2026",
    "nip": "197908112025211039",
    "nama": "AGUS SUWITO, S.E.",
    "jabatan": "Penata Layanan Operasional",
    "hari_kerja": "22",
    "hadir": "22",
    "tidak_hadir": "0",
    "konfirmasi": "0",
    "keterlambatan": "00:00:00",
    "printed_time": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
    "rows": DEFAULT_ROWS
}

def extract_form_data(req):
    data = DEFAULT_DATA.copy()
    logo_base64 = req.form.get("logo_base64_hidden") or None

    data["instansi_baris1"] = req.form.get("instansi_baris1", "")
    data["instansi_baris2"] = req.form.get("instansi_baris2", "")
    data["bulan_tahun"] = req.form.get("bulan_tahun", "")
    data["nip"] = req.form.get("nip", "")
    data["nama"] = req.form.get("nama", "")
    data["jabatan"] = req.form.get("jabatan", "")
    data["hari_kerja"] = req.form.get("hari_kerja", "0")
    data["hadir"] = req.form.get("hadir", "0")
    data["tidak_hadir"] = req.form.get("tidak_hadir", "0")
    data["konfirmasi"] = req.form.get("konfirmasi", "0")
    data["keterlambatan"] = req.form.get("keterlambatan", "00:00:00")
    data["printed_time"] = req.form.get("printed_time") or datetime.now().strftime("%d-%m-%Y %H:%M:%S")

    logo_file = req.files.get("logo")
    if logo_file and logo_file.filename != "":
        encoded = base64.b64encode(logo_file.read()).decode("utf-8")
        logo_base64 = f"data:{logo_file.mimetype};base64,{encoded}"

    date_raws = req.form.getlist("tanggal[]")
    hadirs = req.form.getlist("jam_hadir[]")
    diff_hadirs = req.form.getlist("diff_hadir[]")
    pulangs = req.form.getlist("jam_pulang[]")
    diff_pulangs = req.form.getlist("diff_pulang[]")
    statuses = req.form.getlist("status[]")

    rows = []
    for i in range(len(date_raws)):
        if date_raws[i].strip():
            rows.append({
                "date_raw": date_raws[i],
                "tanggal": format_tanggal_indo(date_raws[i]),
                "hadir": hadirs[i],
                "diff_hadir": diff_hadirs[i],
                "pulang": pulangs[i],
                "diff_pulang": diff_pulangs[i],
                "status": statuses[i]
            })
    data["rows"] = rows
    return data, logo_base64

def set_fixed_table_layout(table, col_widths_mm):
    """Mengunci lebar kolom agar tidak berubah di Microsoft Word maupun Google Docs"""
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    tblPr.append(layout)

    # Set tblGrid secara eksplisit
    tblGrid = table._tbl.tblGrid
    for gridCol in list(tblGrid):
        tblGrid.remove(gridCol)
    for w_mm in col_widths_mm:
        gridCol = OxmlElement('w:gridCol')
        gridCol.set(qn('w:w'), str(int(w_mm * 56.7))) # 1 mm ≈ 56.7 twips
        tblGrid.append(gridCol)

    for row in table.rows:
        for idx, w_mm in enumerate(col_widths_mm):
            cell = row.cells[idx]
            cell.width = Mm(w_mm)
            tcPr = cell._tc.get_or_add_tcPr()
            for child in list(tcPr):
                if child.tag.endswith('tcW'):
                    tcPr.remove(child)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(int(w_mm * 56.7)))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)

def set_all_table_borders(table):
    """Memberikan garis kotak penuh (atas, bawah, kiri, kanan, dalam) pada tabel"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                        '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        '<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        '<w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        '<w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                        '</w:tblBorders>')
    tblPr.append(borders)

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
                        '</w:tblBorders>')
    tblPr.append(borders)

def build_docx(data, logo_base64):
    doc = Document()
    
    # Ukuran Kertas A4 (210 x 297 mm) & Margin kiri/kanan 18 mm -> Lebar area isi = 174 mm
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(15)
    section.bottom_margin = Mm(15)
    section.left_margin = Mm(18)
    section.right_margin = Mm(18)

    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10)

    # 1. Kop Surat (Total lebar = 174 mm)
    kop_table = doc.add_table(rows=1, cols=3)
    kop_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(kop_table)
    set_fixed_table_layout(kop_table, [22, 130, 22])

    cell_logo = kop_table.cell(0, 0)
    if logo_base64 and "," in logo_base64:
        _, b64_data = logo_base64.split(",", 1)
        img_stream = io.BytesIO(base64.b64decode(b64_data))
        p_logo = cell_logo.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_logo = p_logo.add_run()
        run_logo.add_picture(img_stream, width=Mm(13))

    cell_kop = kop_table.cell(0, 1)
    p_kop = cell_kop.paragraphs[0]
    p_kop.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kop.paragraph_format.space_after = Pt(14)
    p_kop.paragraph_format.line_spacing = 1.15
    for idx, line in enumerate([data["instansi_baris1"], data["instansi_baris2"], data["bulan_tahun"]]):
        run = p_kop.add_run(line + ("\n" if idx < 2 else ""))
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)

    # 2. Identitas Pegawai (Total lebar = 174 mm -> 33 mm + 4 mm + 137 mm)
    id_table = doc.add_table(rows=3, cols=3)
    id_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(id_table)
    set_fixed_table_layout(id_table, [33, 4, 137])

    id_rows = [
        ("NIP", "", data["nip"]),
        ("Nama Pegawai", ":", data["nama"]),
        ("Jabatan", ":", data["jabatan"])
    ]
    for r_idx, (label, sep, val) in enumerate(id_rows):
        for c_idx, text in enumerate([label, sep, val]):
            cell = id_table.cell(r_idx, c_idx)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            run = p.add_run(text)
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(4)
    p_space.paragraph_format.space_after = Pt(4)

    # 3. Tabel Presensi (Total lebar = 174 mm -> 10 + 48 + 22 + 22 + 22 + 22 + 28 = 174 mm)
    headers = ["No", "Hari/Tanggal", "Jam Hadir", "+/-", "Jam Pulang", "+/-", "Status"]
    col_widths_mm = [10, 48, 22, 22, 22, 22, 28]
    
    table = doc.add_table(rows=1 + len(data["rows"]), cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_all_table_borders(table)
    set_fixed_table_layout(table, col_widths_mm)

    # Header Tabel
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(h_text)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)

    # Isi Tabel
    for r_idx, row in enumerate(data["rows"], start=1):
        vals = [str(r_idx), row["tanggal"], row["hadir"], row["diff_hadir"], row["pulang"], row["diff_pulang"], row["status"]]
        for c_idx, val in enumerate(vals):
            cell = table.cell(r_idx, c_idx)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
            if c_idx == 1:
                p.paragraph_format.left_indent = Mm(1.5)
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            run = p.add_run(val)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)

    # 4. Rekap Bawah
    p_rekap = doc.add_paragraph()
    p_rekap.paragraph_format.space_before = Pt(12)
    p_rekap.paragraph_format.line_spacing = 1.25
    rekap_lines = [
        f"Jumlah Hari Kerja: {data['hari_kerja']}",
        f"Jumlah Hadir: {data['hadir']}",
        f"Jumlah Tidak Hadir: {data['tidak_hadir']}",
        f"Jumlah Konfirmasi: {data['konfirmasi']}",
        f"Jam Keterlambatan: {data['keterlambatan']}"
    ]
    for idx, line in enumerate(rekap_lines):
        run = p_rekap.add_run(line + ("\n" if idx < len(rekap_lines) - 1 else ""))
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)

    # 5. Footer System
    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_footer.paragraph_format.space_before = Pt(16)
    run_footer = p_footer.add_run(f"Printed by System on {data['printed_time']}")
    run_footer.bold = True
    run_footer.font.name = 'Times New Roman'
    run_footer.font.size = Pt(9.5)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        data, logo_base64 = extract_form_data(request)
    else:
        data = DEFAULT_DATA.copy()
        logo_base64 = None
    return render_template("index.html", data=data, logo_base64=logo_base64)

@app.route("/download-pdf", methods=["POST"])
def download_pdf():
    data, logo_base64 = extract_form_data(request)
    rendered_html = render_template("pdf_template.html", data=data, logo_base64=logo_base64)
    
    pdf_bytes = HTML(string=rendered_html).write_pdf()
    nama_bersih = data["nama"].replace(" ", "_").replace(",", "").replace(".", "")
    filename = f"Rekap_Presensi_{nama_bersih}.pdf"
    
    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename
    )

@app.route("/download-docx", methods=["POST"])
def download_docx():
    data, logo_base64 = extract_form_data(request)
    docx_stream = build_docx(data, logo_base64)
    nama_bersih = data["nama"].replace(" ", "_").replace(",", "").replace(".", "")
    filename = f"Rekap_Presensi_{nama_bersih}.docx"
    
    return send_file(
        docx_stream,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
