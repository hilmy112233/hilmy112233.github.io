"""Generate the IT Support Engineer variant of the CV as editable .docx files.

Same ATS rules as build-docx.py: no text boxes, no layout tables, no
header/footer content, real List Bullet style, one column, contact in the body.

Content is kept token-for-token identical to cv/cv-itsupport.html so the PDF and
the Word file extract the same text. Run cv/build-docx-itsupport.py, then check
with _docxcheck.py and a Word page count.
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

SZ_NAME = 17
SZ_ROLE = 9.5
SZ_CONTACT = 8.5
SZ_HEAD = 8.5
SZ_BODY = 8.6
SZ_JOB = 10
SZ_SKILL = 8.6
LINE = 0.98

GAP_SECTION = 4
GAP_JOB = 3.2
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
    role = ("IT Support Engineer · Perangkat Keras, Perangkat & Dukungan Lapangan"
            if not en else
            "IT Support Engineer · Hardware, Devices & Field Support")
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
        "Informatics graduate (GPA 3.47/4.00) with two years of hands-on device and field "
        "support: installed, wired and tested seven control panels across seven sites, then "
        "diagnosed the faults that took them down — mode switching that would not latch, relay "
        "chatter, sensor nodes dropping off — and restored the manual fail-safe where automation "
        "failed. Configures and supports ESP32 and Raspberry Pi devices, runs MQTT/TLS telemetry, "
        "and builds the internal tools field staff actually use. Comfortable on Linux, Docker and "
        "PostgreSQL/MySQL, with clear written handover."
        if en else
        "Lulusan Teknik Informatika (IPK 3.47/4.00) dengan dua tahun dukungan perangkat keras di "
        "lapangan: memasang, mengkabel, dan menguji tujuh panel kontrol di tujuh lokasi, lalu "
        "mendiagnosis gangguannya — peralihan mode tidak mengunci, relay chatter, node sensor "
        "putus — dan memulihkan jalur fail-safe manual saat otomasi gagal. Mengonfigurasi "
        "perangkat ESP32 dan Raspberry Pi, menjalankan telemetri MQTT/TLS, membangun tools "
        "internal yang dipakai staf lapangan. Terbiasa Linux, Docker, PostgreSQL/MySQL, dan serah "
        "terima tertulis."
    )
    p = para(doc, space_after=0, line=1.02)
    r = p.add_run(summary)
    set_font(r, SZ_BODY, color=SOFT)

    # ---- experience ---------------------------------------------------
    section_heading(doc, "Experience" if en else "Pengalaman Kerja")

    EXP = [
        {
            "title": "Panel Assembler, Wirer & Tester/Support (Contract)" if en else "Panel Assembler, Wirer & Tester/Support (Kontrak)",
            "dates": "Jan 2023 – Sep 2025",
            "sub": ("Team Aligness · Smart Greenhouse Project, Ministry of Agriculture — Subang & Sumedang"
                    if en else
                    "Team Aligness · Proyek Smart Greenhouse, Kementerian Pertanian — Subang & Sumedang"),
            "bullets": [
                "Installed, wired and commissioned seven control panels across seven 18 × 18 m greenhouse sites — cutting and drilling the mounting plate, fitting every component, then soldering, crimping and clamping the full harness."
                if en else
                "Memasang, mengkabel, dan mengomisioning tujuh panel kontrol di tujuh lokasi greenhouse 18 × 18 m — potong dan lubang pelat dudukan, pasang komponen, lalu solder, crimp, dan klem seluruh harness.",
                "Diagnosed on-site faults down to root cause: mode switching that would not latch, relay chatter, and sensor nodes dropping their connection — then fixed the wiring or replaced the part and re-tested until the panel held."
                if en else
                "Mendiagnosis gangguan di lokasi sampai akar masalahnya: peralihan mode tidak mengunci, relay chatter, node sensor lepas koneksi — lalu perbaiki jalur kabel atau ganti komponen dan uji ulang sampai stabil.",
                "Configured and tested the CO2/light sensor nodes reporting every 10 seconds, plus the relay and fail-safe paths behind them, verifying no uncommanded actuation after every change."
                if en else
                "Mengonfigurasi dan menguji node sensor CO2/cahaya yang melapor tiap 10 detik, beserta jalur relay dan fail-safe di belakangnya, memastikan tidak ada aktuasi tak terkendali setelah tiap perubahan.",
                "Documented and escalated that two of the seven sites could not hold automated mode, then recommended the manual fail-safe — which is what still runs daily."
                if en else
                "Mendokumentasikan dan melaporkan bahwa dua dari tujuh situs tidak mampu bertahan di mode otomatis, lalu merekomendasikan jalur fail-safe manual — yang sampai sekarang masih dipakai harian.",
            ],
        },
        {
            "title": "Frontend Mobile Developer (Intern)" if en else "Frontend Mobile Developer (Magang)",
            "dates": "Aug 2025 – Jan 2026",
            "sub": "Leapfactor · PT Lattice Teknologi Mandiri — Bandung, Indonesia",
            "bullets": [
                "Supported and extended Arjuna Mobile, the inventory app field staff run on their phones — fixed camera-lifecycle bugs, added debounce against duplicate scans, and kept REST sync consistent with the backend."
                if en else
                "Mendukung dan mengembangkan Arjuna Mobile, aplikasi inventaris yang dipakai staf lapangan di ponselnya — membenahi bug siklus kamera, menambah debounce anti scan-ganda, dan menjaga sinkronisasi REST dengan backend.",
                "Integrated barcode/QR scanner hardware through mobile_scanner, including releasing the camera on route exit so the device did not hang."
                if en else
                "Mengintegrasikan pemindai barcode/QR lewat mobile_scanner, termasuk melepas kamera saat keluar route supaya perangkat tidak menggantung.",
            ],
        },
    ]

    for job in EXP:
        role_line(doc, job["title"], job["dates"])
        if job.get("sub"):
            sub_line(doc, job["sub"])
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
                "Integrated the sensor chain end to end — a 16-bit ADS1115 ADC on the tank float, NEO-6M GPS over serial, an ESP32 pushing MQTT/TLS — then traced the noise that made readings unusable to engine vibration and fuel slosh."
                if en else
                "Merangkai jalur sensor dari ujung ke ujung — ADC 16-bit ADS1115 pada pelampung tangki, GPS NEO-6M lewat serial, ESP32 mengirim MQTT/TLS — lalu menelusuri noise pembacaan ke getaran mesin dan fuel slosh.",
                "Fixed it in firmware with a three-stage filter (7-sample median > EMA alpha 0.1 > 0.03 L hysteresis) and streamed telemetry every second to a Django + PostgreSQL dashboard."
                if en else
                "Memperbaikinya di firmware dengan filter tiga lapis (median 7 sampel > EMA alpha 0.1 > histeresis 0.03 L) dan mengalirkan telemetri per detik ke dashboard Django + PostgreSQL.",
            ],
        },
        {
            "title": ("Flood & Water-Tank Monitoring — Panel & App Interface" if en
                      else "Flood & Water-Tank Monitoring — Antarmuka Panel & Aplikasi"),
            "dates": ("2026 – Present" if en else "2026 – Sekarang"),
            "sub": "Team Aligness — UI/UX Designer",
            "bullets": [
                "Designed the companion mobile app (25 Figma frames) on the same status vocabulary as the panel, so a fault reads the same on the wall as on the user's phone."
                if en else
                "Merancang aplikasi pendamping seluler (25 layar Figma) memakai kosakata status yang sama dengan panel, supaya satu gangguan terbaca sama di dinding maupun di ponsel pengguna.",
            ],
        },
        {
            "title": "CineStream — Video Streaming Platform" if en else "CineStream — Platform Streaming Video",
            "dates": "2025",
            "sub": "Project Manager & System Analyst — " + ("SRS, UML, Laravel" if en else "SKPL/SRS, UML, Laravel"),
            "bullets": [
                "Wrote the SRS, the full UML set and the ERD before any code was written, then built the PHP/Laravel app with a team of six — the same documentation habit used for handover notes and fault reports."
                if en else
                "Menulis SKPL/SRS, seluruh diagram UML, dan ERD sebelum satu baris kode ditulis, lalu membangun aplikasi PHP/Laravel-nya bersama tim enam orang — kebiasaan dokumentasi yang sama dipakai untuk catatan serah terima dan laporan gangguan.",
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
        ("Hardware & Device Support" if en else "Dukungan Perangkat Keras",
         ("panel installation & commissioning, wiring (solder/crimp/clamp), component fitting, "
          "fault diagnosis & root-cause tracing, corrective maintenance & repair, functional testing, device configuration, "
          "relays, fail-safe paths"
          if en else
          "instalasi & komisioning panel, pengkabelan (solder/crimp/klem), pemasangan komponen, "
          "diagnosis gangguan & pelacakan akar masalah, perawatan korektif & perbaikan, uji fungsi, konfigurasi perangkat, relay, "
          "jalur fail-safe")),
        ("Systems & Tools" if en else "Sistem & Perkakas",
         "Linux / Raspberry Pi OS, Windows, Docker, Git & GitHub, VS Code, PlatformIO, Postman, remote troubleshooting"),
        ("Instrumentation & IoT" if en else "Instrumentasi & IoT",
         ("ESP32, Raspberry Pi, CO2/light sensor nodes, ADS1115 ADC, NEO-6M GPS, MQTT/TLS, ThingsBoard"
          if en else
          "ESP32, Raspberry Pi, node sensor CO2/cahaya, ADC ADS1115, GPS NEO-6M, MQTT/TLS, ThingsBoard")),
        ("Data & Backend",
         "PostgreSQL, MySQL, Django, REST APIs, Laravel"),
        ("Programming & Mobile" if en else "Pemrograman & Mobile",
         "Python, C/C++, SQL, Dart, PHP, JavaScript, HTML/CSS, Flutter, Provider, mobile_scanner"),
        ("Documentation & Analysis" if en else "Dokumentasi & Analisis",
         ("SRS/SyRS, UML, ERD, DFD, requirement gathering, Agile & Waterfall, Figma"
          if en else
          "SRS/SyRS, UML, ERD, DFD, requirement gathering, Agile & Waterfall, Figma")),
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
    note_line(doc, "Bandung, Indonesia — GPA 3.47 / 4.00" if en
              else "Bandung, Indonesia — IPK 3.47 / 4.00")

    role_line(doc, "Certified System Analyst (CSA)", "2026 – 2029")
    note_line(doc, ("BNSP · LSP Informatika — SRS/SyRS, UML/ERD/DFD, system feasibility verification." if en
                    else "BNSP · LSP Informatika — SRS/SyRS, UML/ERD/DFD, verifikasi kelayakan sistem."))

    cp = doc.core_properties
    cp.author = "Hilmy Abdurrahman Darmawan"
    cp.title = "Hilmy Abdurrahman Darmawan — CV (IT Support Engineer)"
    cp.subject = "Curriculum Vitae"
    cp.keywords = ("IT Support, Hardware Support, Troubleshooting, Device Configuration, "
                   "Linux, Windows, Docker, ESP32, MQTT")

    doc.save(out_path)
    print(f"wrote {out_path}")


if __name__ == "__main__":
    build('en', 'assets/CV_Hilmy_Abdurrahman_Darmawan_ITSupport.docx')
    build('id', 'assets/CV_Hilmy_Abdurrahman_Darmawan_ITSupport_ID.docx')
