"""Generate editable, ATS-safe .docx versions of the CV.

Rules this follows so Word files survive an ATS:
  - no text boxes, no tables for layout, no header/footer contact block
  - real Word styles (Heading 1/2, List Bullet) so parsers see structure
  - one column, top to bottom
  - no images, no columns, no line shapes
  - contact details in the body, as the first paragraph

Two files are produced: English and Indonesian. Both are editable in Word,
Google Docs, or LibreOffice with no layout breakage.

The type scale is tuned so the result is one A4 page, matching the PDFs.
Run cv/build-docx.py, then verify with _docxcheck.py and a Word page count.
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GREEN = RGBColor(0x1F, 0x4D, 0x3A)
RUST = RGBColor(0xA6, 0x5A, 0x1B)
INK = RGBColor(0x18, 0x22, 0x1D)
SOFT = RGBColor(0x34, 0x42, 0x3B)
MUTED = RGBColor(0x4A, 0x5A, 0x52)

FONT = "Calibri"

# Type scale, tuned to land on one A4 page. Sized in points.
SZ_NAME = 17
SZ_ROLE = 9.5
SZ_CONTACT = 8.5
SZ_HEAD = 8.5
SZ_BODY = 8.6
SZ_JOB = 10
SZ_SKILL = 8.6
LINE = 0.98

# vertical rhythm
GAP_SECTION = 6
GAP_JOB = 3.5
GAP_BULLET = 0.8


def set_font(run, size=None, bold=None, color=None, italic=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.append(rf)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(attr), FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def para(doc, space_before=0, space_after=0, line=LINE):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    pf.widow_control = False
    return p


def bottom_rule(p, color="1F4D3A", size=8):
    """A paragraph border, not a shape -- ATS-safe and stays put when edited."""
    pPr = p._element.get_or_add_pPr()
    bdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    bdr.append(bottom)
    pPr.append(bdr)


def section_heading(doc, text):
    p = para(doc, space_before=GAP_SECTION, space_after=2)
    r = p.add_run(text.upper())
    set_font(r, SZ_HEAD, bold=True, color=GREEN)
    rPr = r._element.get_or_add_rPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:val'), '16')
    rPr.append(sp)
    bottom_rule(p)
    return p


def role_line(doc, title, dates):
    p = para(doc, space_before=GAP_JOB, space_after=0)
    r = p.add_run(title)
    set_font(r, SZ_JOB, bold=True, color=INK)
    r2 = p.add_run("\t" + dates)
    set_font(r2, SZ_BODY, color=SOFT)
    # right-aligned tab stop so dates sit flush right without a table
    pPr = p._element.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:pos'), str(int(6.85 * 1440)))
    tabs.append(tab)
    pPr.append(tabs)
    return p


def sub_line(doc, text):
    p = para(doc, space_before=0, space_after=1)
    r = p.add_run(text)
    set_font(r, SZ_BODY, italic=True, color=SOFT)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(GAP_BULLET)
    pf.line_spacing = LINE
    pf.left_indent = Inches(0.20)
    pf.first_line_indent = Inches(-0.13)
    pf.widow_control = False
    r = p.add_run(text)
    set_font(r, SZ_BODY, color=SOFT)
    return p


def note_line(doc, text):
    p = para(doc, space_before=0, space_after=1)
    r = p.add_run(text)
    set_font(r, SZ_BODY, color=SOFT)
    return p


def skill_line(doc, label, value):
    p = para(doc, space_before=0, space_after=0.8, line=1.0)
    r = p.add_run(label + ": ")
    set_font(r, SZ_SKILL, bold=True, color=GREEN)
    r2 = p.add_run(value)
    set_font(r2, SZ_SKILL, color=SOFT)
    return p


def build(lang, out_path):
    en = lang == 'en'
    doc = Document()

    # A4 with tight margins, matching the PDF
    s = doc.sections[0]
    s.page_width = Inches(8.27)
    s.page_height = Inches(11.69)
    s.top_margin = Inches(0.34)
    s.bottom_margin = Inches(0.34)
    s.left_margin = Inches(0.43)
    s.right_margin = Inches(0.43)

    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = Pt(SZ_BODY)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    # ---- header -------------------------------------------------------
    p = para(doc, space_after=1)
    r = p.add_run("Hilmy Abdurrahman Darmawan")
    set_font(r, SZ_NAME, bold=True, color=GREEN)

    p = para(doc, space_after=1)
    role = ("Mobile Developer · Analis Sistem Bersertifikat (CSA) · Perakitan & Pengkabelan Panel"
            if not en else
            "Mobile Developer · Certified System Analyst (CSA) · Control Panel Assembly & Wiring")
    r = p.add_run(role)
    set_font(r, SZ_ROLE, bold=True, color=RUST)

    p = para(doc, space_after=0)
    r = p.add_run("Bandung, " + ("West Java" if en else "Jawa Barat")
                  + ", Indonesia | +62 838-9189-9401 | adhilmy159@gmail.com")
    set_font(r, SZ_CONTACT, color=MUTED)

    p = para(doc, space_after=2)
    r = p.add_run("linkedin.com/in/hilmy-abdurrahman-darmawan | github.com/hilmy112233 | hilmy112233.github.io")
    set_font(r, SZ_CONTACT, color=MUTED)
    bottom_rule(p, size=10)

    # ---- summary ------------------------------------------------------
    section_heading(doc, "Summary" if en else "Ringkasan")
    summary = (
        "Informatics graduate (GPA 3.47/4.00) and BNSP-certified System Analyst: Flutter mobile, "
        "systems analysis, and interface design — a field-used Flutter inventory app, SRS and UML "
        "for a streaming platform, and the interface of a wall-mounted monitoring panel. Hands-on "
        "IoT work building and wiring greenhouse control panels, then testing them."
        if en else
        "Lulusan Teknik Informatika (IPK 3.47/4.00) dan Certified System Analyst (BNSP): mobile "
        "Flutter, analisis sistem, dan desain antarmuka — aplikasi inventaris Flutter yang dipakai "
        "di lapangan, SRS dan UML untuk platform streaming, serta antarmuka panel monitoring dinding. "
        "Pengalaman lapangan IoT dari merakit dan mengkabel panel kontrol greenhouse, lalu mengujinya."
    )
    p = para(doc, space_after=0, line=1.02)
    r = p.add_run(summary)
    set_font(r, SZ_BODY, color=SOFT)

    # ---- experience ---------------------------------------------------
    section_heading(doc, "Experience" if en else "Pengalaman Kerja")

    EXP = [
        {
            "title": "Frontend Mobile Developer (Intern)" if en else "Frontend Mobile Developer (Magang)",
            "dates": "Aug 2025 – Jan 2026",
            "sub": "Leapfactor · PT Lattice Teknologi Mandiri — Bandung, Indonesia",
            "bullets": [
                "Built Arjuna Mobile inventory flows in Flutter/Dart on a modular architecture, using Provider for auth state, navigation, and item mutation cycles."
                if en else
                "Membangun alur inventaris Arjuna Mobile dengan Flutter/Dart di atas arsitektur modular, memakai Provider untuk siklus state autentikasi, navigasi, dan mutasi item.",
                "Integrated real-time barcode/QR scanning with mobile_scanner — releasing the camera on route exit and debouncing duplicate scans — then wired the UI to REST endpoints for token auth and item mutation sync."
                if en else
                "Mengintegrasikan pemindaian barcode/QR lewat mobile_scanner — kamera dilepas saat keluar route, debounce anti scan-ganda — lalu menyambungkan UI ke REST API untuk autentikasi token dan sinkronisasi mutasi barang.",
            ],
        },
        {
            # Job titles stay in English in both languages -- they are the
            # terms an ATS keyword-matches on, and the PDF does the same.
            "title": "Panel Assembler, Wirer & Tester (Contract)" if en else "Panel Assembler, Wirer & Tester (Kontrak)",
            "dates": "Jan 2023 – Sep 2025",
            "sub": ("Smart Greenhouse Project, Ministry of Agriculture · Team Aligness — Subang & Sumedang"
                    if en else
                    "Proyek Smart Greenhouse, Kementerian Pertanian · Team Aligness — Subang & Sumedang"),
            "bullets": [
                "Assembled and wired five control panels for five 18 × 18 m greenhouse sites — plate cutting and drilling, then soldering, crimping and clamping every harness."
                if en else
                "Merakit dan mengkabel lima panel kontrol untuk lima greenhouse 18 × 18 m — potong dan lubang pelat, lalu solder, crimp, dan klem seluruh harness.",
                "Tested mode switching and fault behaviour across all five panels and their CO2/light sensor nodes, reading every 10 seconds — no relay chatter, no uncommanded actuation."
                if en else
                "Menguji peralihan mode dan perilaku fault pada lima panel serta node sensor CO2/cahaya yang membaca tiap 10 detik — tanpa relay chatter atau aktuasi tak terkendali.",
                "Reported that four of the five sites reverted to manual — the finding that shaped the project's conclusion, and why the manual fail-safe is what still runs daily."
                if en else
                "Melaporkan bahwa empat dari lima situs kembali ke manual — temuan yang membentuk kesimpulan proyek, dan alasan jalur fail-safe manual masih dipakai tiap hari.",
            ],
        },
        {
            "title": "UI/UX Designer & Independent Contractor",
            "dates": ("2024 – Present" if en else "2024 – Sekarang"),
            "sub": None,
            "note": ("Projects.co.id — wireframes, user flows and high-fidelity UI in Figma for tour booking and e-commerce, plus a design system (reusable components, colour tokens, type hierarchy)."
                     if en else
                     "Projects.co.id — wireframe, user flow, dan UI high-fidelity di Figma untuk pemesanan tur dan e-commerce, plus design system (komponen reusable, token warna, hierarki tipografi)."),
        },
    ]

    for job in EXP:
        role_line(doc, job["title"], job["dates"])
        if job.get("sub"):
            sub_line(doc, job["sub"])
        if job.get("note"):
            note_line(doc, job["note"])
        for b in job.get("bullets", []):
            bullet(doc, b)

    # ---- projects -----------------------------------------------------
    section_heading(doc, "Selected Projects" if en else "Proyek Pilihan")

    PROJ = [
        {
            "title": ("IoT Fuel Consumption Monitoring for Carburettor Motorcycles" if en
                      else "Monitoring Konsumsi BBM Motor Karburator Berbasis IoT"),
            "dates": "2025 – 2026",
            "sub": ("S1 Capstone Thesis · UNIKOM — Hardware & Software Developer" if en
                    else "Tugas Akhir S1 · UNIKOM — Hardware & Software Developer"),
            "bullets": [
                "Replaced an imprecise analog bar indicator with real-time ESP32 telemetry, after cutting engine-vibration and fuel-slosh noise with a three-stage filter: 7-sample median > EMA (alpha = 0.1) > 0.03 L hysteresis."
                if en else
                "Menggantikan indikator analog yang tidak presisi dengan telemetri digital real-time ESP32, setelah meredam noise getaran mesin dan fuel sloshing lewat filter tiga lapis: median 7 sampel > EMA (alpha = 0.1) > histeresis 0.03 L.",
                "Read the tank float via a 16-bit ADS1115 ADC, accumulated distance from NEO-6M GPS with the Haversine formula, and pushed per-second telemetry over MQTT/TLS to a Django + PostgreSQL dashboard."
                if en else
                "Membaca pelampung tangki via ADC 16-bit ADS1115, mengakumulasi jarak dari GPS NEO-6M dengan formula Haversine, lalu mengirim telemetri per detik lewat MQTT/TLS ke dashboard Django + PostgreSQL.",
            ],
        },
        {
            "title": ("Flood & Water Tank — Panel & App Interface Design" if en
                      else "Panel Banjir & Tangki Air — Desain Antarmuka Panel & Aplikasi"),
            "dates": "2026 · " + ("team" if en else "tim"),
            "sub": ("Aligness — UI/UX Designer for both surfaces; firmware, sensor network and enclosure built by other team members"
                    if en else
                    "Aligness — UI/UX Designer untuk kedua antarmuka; firmware, jaringan sensor, dan enclosure dikerjakan anggota tim lain"),
            "bullets": [
                "Designed a 7-inch ESP32-P4 wall panel for two non-technical on-site operators: an always-visible dashboard and two oversized menu buttons, because they would never open an app while standing at the tank."
                if en else
                "Merancang panel dinding ESP32-P4 7 inci untuk dua pengguna non-teknis di lokasi: dashboard yang selalu terlihat dan dua tombol menu besar, karena mereka tidak akan membuka aplikasi saat berdiri di depan tangki.",
                "Designed the companion mobile app in full — 25 Figma frames across authentication, home status, device pairing and settings — reusing the panel's status vocabulary so the same condition reads the same on both screens."
                if en else
                "Merancang aplikasi pendamping seluler secara utuh: 25 layar Figma untuk autentikasi, status rumah, pemasangan perangkat, dan pengaturan, dengan kosakata status yang sama seperti panel.",
                "Encoded status as word + colour + icon rather than a chart, sized every target for thumb use at a 1.5–2 m viewing distance, and logged the remaining gaps as the next pass."
                if en else
                "Menyusun status sebagai kata + warna + ikon alih-alih grafik, menetapkan ukuran target sentuh untuk jempol pada jarak baca 1,5–2 m, dan mencatat celah desain.",
            ],
        },
        {
            "title": "CineStream — Video Streaming Platform" if en else "CineStream — Platform Streaming Video",
            "dates": "2025 · " + ("team of 6" if en else "tim 6 orang"),
            "sub": "Project Manager & System Analyst — " + ("SRS, UML, Laravel" if en else "SKPL/SRS, UML, Laravel"),
            "bullets": [
                "Led a six-person team as PM and System Analyst: wrote the SRS, the full UML set (use case, activity, sequence, class) and the ERD before any code was written, then built the PHP/Laravel app consuming public movie-catalogue APIs."
                if en else
                "Memimpin tim enam orang sebagai PM sekaligus System Analyst: menulis SKPL/SRS, seluruh diagram UML (use case, activity, sequence, class), dan ERD sebelum satu baris kode ditulis, lalu membangun aplikasi PHP/Laravel-nya.",
            ],
        },
    ]

    for job in PROJ:
        role_line(doc, job["title"], job["dates"])
        if job.get("sub"):
            sub_line(doc, job["sub"])
        for b in job.get("bullets", []):
            bullet(doc, b)

    # ---- skills -------------------------------------------------------
    section_heading(doc, "Technical Skills" if en else "Keahlian Teknis")
    SKILLS = [
        ("Programming" if en else "Bahasa Pemrograman",
         "C/C++, Python, Dart, PHP, SQL, JavaScript, HTML/CSS"),
        ("Mobile & Web",
         "Flutter (Dart), Provider, mobile_scanner, REST APIs, Laravel, Django"),
        ("Hardware & IoT" if en else "Perangkat Keras & IoT",
         ("panel assembly & plate work (cut/drill), wiring & soldering/crimping/clamping, panel testing, relays, ESP32 / Raspberry Pi (test & support level), ThingsBoard, MQTT"
          if en else
          "perakitan panel & pelat (potong/lubang), pengkabelan & solder/crimp/klem, uji panel, relay, ESP32 / Raspberry Pi (level uji & dukung), ThingsBoard, MQTT")),
        ("Analysis & Design" if en else "Analisis & Desain",
         ("SRS/SyRS, UML, ERD, DFD, requirement gathering, Waterfall & Agile, Figma, interface design"
          if en else
          "SRS/SyRS, UML, ERD, DFD, requirement gathering, Waterfall & Agile, Figma, desain antarmuka")),
        ("Data & Tools" if en else "Data & Alat",
         "PostgreSQL, MySQL, Docker, Git & GitHub, VS Code, PlatformIO, Linux / Raspberry Pi OS, Postman, Pentaho"),
        ("Languages" if en else "Bahasa",
         ("Indonesian (native) · English (professional — technical, documentation, work communication)"
          if en else
          "Indonesia (asli) · Inggris (profesional — teknis, dokumentasi, komunikasi kerja)")),
    ]
    for label, value in SKILLS:
        skill_line(doc, label, value)

    # ---- education ----------------------------------------------------
    section_heading(doc, "Education & Certification" if en else "Pendidikan & Sertifikasi")
    role_line(doc, ("S.Kom — Informatics, Universitas Komputer Indonesia (UNIKOM)" if en
                    else "S.Kom — Teknik Informatika, Universitas Komputer Indonesia (UNIKOM)"),
              "2022 – 2026")
    note_line(doc, ("Bandung, Indonesia — GPA 3.47 / 4.00 — Bachelor of Computer Science" if en
                    else "Bandung, Indonesia — IPK 3.47 / 4.00 — setara Bachelor of Computer Science"))

    role_line(doc, "Certified System Analyst (CSA)", "2026 – 2029")
    note_line(doc, ("BNSP · LSP Informatika — SRS/SyRS, UML/ERD/DFD, system feasibility verification." if en
                    else "BNSP · LSP Informatika — SRS/SyRS, UML/ERD/DFD, verifikasi kelayakan sistem."))

    cp = doc.core_properties
    cp.author = "Hilmy Abdurrahman Darmawan"
    cp.title = "Hilmy Abdurrahman Darmawan — CV"
    cp.subject = "Curriculum Vitae"
    cp.keywords = "Flutter, System Analyst, IoT, ESP32, UI/UX, Figma, SRS, UML"

    doc.save(out_path)
    print(f"wrote {out_path}")


if __name__ == "__main__":
    build('en', 'assets/CV_Hilmy_Abdurrahman_Darmawan.docx')
    build('id', 'assets/CV_Hilmy_Abdurrahman_Darmawan_ID.docx')
