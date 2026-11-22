import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client
from pypdf import PdfReader, PdfWriter, PageObject, Transformation


def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def set_cell_borders(cell, color="E2E8F0", sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)


def set_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def add_page_number_to_footer(section):
    section.different_first_page_header_footer = True
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

    r_pre = p.add_run("ITST 301 Portfolio Guide  •  Maribel Bravo  •  Page ")
    r_pre.font.name = "Segoe UI"
    r_pre.font.size = Pt(8.5)
    r_pre.font.color.rgb = RGBColor(100, 116, 139)

    fldSimple = OxmlElement("w:fldSimple")
    fldSimple.set(qn("w:instr"), "PAGE")
    r_pg = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Segoe UI")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "17")  # 8.5 pt
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "7C3AED")
    b = OxmlElement("w:b")
    rPr.append(rFonts)
    rPr.append(sz)
    rPr.append(color)
    rPr.append(b)
    r_pg.append(rPr)
    fldSimple.append(r_pg)
    p._p.append(fldSimple)


def create_brochure_manual():
    doc = docx.Document()

    # Configure IPT2 Finished Booklet Size: 6.5 x 8.5 in (half of 8.5 x 13 in long bond)
    for section in doc.sections:
        section.page_width = Inches(6.5)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)   # extra space at fold
        section.right_margin = Inches(0.5)
        section.footer_distance = Inches(0.28)
        # Enable mirror margins for duplex booklet folding
        sectPr = section._sectPr
        mirror = OxmlElement("w:mirrorMargins")
        sectPr.append(mirror)
        add_page_number_to_footer(section)

    C_NAVY = RGBColor(15, 23, 42)       # #0F172A
    C_PURPLE = RGBColor(124, 58, 237)   # #7C3AED
    C_CYAN = RGBColor(6, 182, 212)      # #06B6D4
    C_GRAY = RGBColor(100, 116, 139)    # #64748B
    C_DARK = RGBColor(30, 41, 59)       # #1E293B
    C_WHITE = RGBColor(255, 255, 255)

    def add_section_banner(kicker, title):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_no_borders(tbl)
        c = tbl.rows[0].cells[0]
        c.width = Inches(5.4)
        set_cell_bg(c, "0F172A")
        set_cell_margins(c, top=90, bottom=90, left=140, right=140)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        rk = p.add_run(kicker.upper() + "\n")
        rk.font.name = "Segoe UI"
        rk.font.size = Pt(8)
        rk.bold = True
        rk.font.color.rgb = C_CYAN
        rt = p.add_run(title)
        rt.font.name = "Segoe UI"
        rt.font.size = Pt(14)
        rt.bold = True
        rt.font.color.rgb = C_WHITE

        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    def add_subhead(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(text)
        r.font.name = "Segoe UI"
        r.font.size = Pt(11.5)
        r.bold = True
        r.font.color.rgb = C_PURPLE
        return p

    def add_body(text, bold_prefix="", space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            rb = p.add_run(bold_prefix + " ")
            rb.font.name = "Segoe UI"
            rb.font.size = Pt(11)
            rb.bold = True
            rb.font.color.rgb = C_NAVY
        r = p.add_run(text)
        r.font.name = "Segoe UI"
        r.font.size = Pt(11)
        r.font.color.rgb = C_DARK
        return p

    def add_callout_card(title, body, bg="F8FAFC", accent="7C3AED"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = tbl.rows[0].cells[0]
        c.width = Inches(5.4)
        set_cell_bg(c, bg)
        set_cell_margins(c, top=70, bottom=70, left=130, right=120)
        tcPr = c._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'<w:left w:val="single" w:sz="18" w:space="0" w:color="{accent}"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.12
        rt = p.add_run(title + "\n")
        rt.font.name = "Segoe UI"
        rt.font.size = Pt(10.5)
        rt.bold = True
        rt.font.color.rgb = C_NAVY
        rb = p.add_run(body)
        rb.font.name = "Segoe UI"
        rb.font.size = Pt(10)
        rb.font.color.rgb = C_DARK
        sp = doc.add_paragraph()
        sp.paragraph_format.space_after = Pt(3)

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 1: BROCHURE FRONT COVER
    # ══════════════════════════════════════════════════════════════════════
    p_univ = doc.add_paragraph()
    p_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_univ.paragraph_format.space_before = Pt(4)
    p_univ.paragraph_format.space_after = Pt(10)
    ru1 = p_univ.add_run("REPUBLIC OF THE PHILIPPINES\nLAGUNA STATE POLYTECHNIC UNIVERSITY\n")
    ru1.font.name = "Segoe UI"
    ru1.font.size = Pt(10)
    ru1.bold = True
    ru1.font.color.rgb = C_NAVY
    ru2 = p_univ.add_run("Siniloan Campus • College of Computer Studies")
    ru2.font.name = "Segoe UI"
    ru2.font.size = Pt(9.5)
    ru2.font.color.rgb = C_GRAY

    # Hero Brochure Cover Panel
    cover_tbl = doc.add_table(rows=1, cols=1)
    cover_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_no_borders(cover_tbl)
    cc = cover_tbl.rows[0].cells[0]
    cc.width = Inches(5.4)
    set_cell_bg(cc, "0F172A")
    set_cell_margins(cc, top=240, bottom=240, left=180, right=180)
    cp = cc.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_after = Pt(0)

    r_badge = cp.add_run("ITST 301  •  PRINCIPLES OF WEB DESIGN\n\n")
    r_badge.font.name = "Segoe UI"
    r_badge.font.size = Pt(9.5)
    r_badge.bold = True
    r_badge.font.color.rgb = C_CYAN

    r_main = cp.add_run("PERSONAL PORTFOLIO\nBROCHURE & SYSTEM GUIDE\n\n")
    r_main.font.name = "Segoe UI"
    r_main.font.size = Pt(19)
    r_main.bold = True
    r_main.font.color.rgb = C_WHITE

    r_desc = cp.add_run("Planning, Semantic HTML5, Responsive Bootstrap 5,\nAccessibility & Live Deployment Manual")
    r_desc.font.name = "Segoe UI"
    r_desc.font.size = Pt(10.5)
    r_desc.font.color.rgb = RGBColor(203, 213, 225)

    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_after = Pt(8)

    # Developer Photo + Info Brochure Strip
    info_tbl = doc.add_table(rows=1, cols=2)
    info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_no_borders(info_tbl)
    c_img, c_txt = info_tbl.rows[0].cells[0], info_tbl.rows[0].cells[1]
    c_img.width = Inches(1.5)
    c_txt.width = Inches(3.9)
    set_cell_bg(c_img, "F8FAFC")
    set_cell_bg(c_txt, "F8FAFC")
    set_cell_margins(c_img, 100, 100, 100, 60)
    set_cell_margins(c_txt, 100, 100, 60, 100)

    prof_path = r"C:\Users\moral\OneDrive\Practice\basic-portfolio\assets\profile.jpg"
    p_im = c_img.paragraphs[0]
    p_im.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(prof_path):
        p_im.add_run().add_picture(prof_path, width=Inches(1.15))

    p_tx = c_txt.paragraphs[0]
    p_tx.paragraph_format.line_spacing = 1.2
    for lbl, val in [
        ("Developer: ", "Maribel Bravo\n"),
        ("Student ID: ", "0224-0865\n"),
        ("Section: ", "BSIT 3A WMAD\n"),
        ("Instructor: ", "Prof. Julius A. Derla\n"),
        ("Subject: ", "ITST 301 / IPT2 Final Project"),
    ]:
        rl = p_tx.add_run(lbl)
        rl.font.name = "Segoe UI"
        rl.font.size = Pt(10)
        rl.bold = True
        rl.font.color.rgb = C_PURPLE
        rv = p_tx.add_run(val)
        rv.font.name = "Segoe UI"
        rv.font.size = Pt(10)
        rv.font.color.rgb = C_NAVY

    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foot.paragraph_format.space_before = Pt(16)
    rf = p_foot.add_run(
        "Live Portfolio: https://hiddengemofbell.github.io/basic-portfolio/\n"
        "GitHub Repo: https://github.com/hiddengemofbell/basic-portfolio\n"
        "Academic Year 2026–2027"
    )
    rf.font.name = "Segoe UI"
    rf.font.size = Pt(9)
    rf.font.color.rgb = C_GRAY

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 2: TABLE OF CONTENTS & DEVELOPER PROFILE
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Guide Overview", "Table of Contents & Developer Profile")

    toc_rows = [
        ("01", "Developer Profile & Core Skills", "p. 2"),
        ("02", "Portfolio Purpose & Section Guide", "p. 3"),
        ("03", "Personal Logo & Color Palette (HEX)", "p. 4"),
        ("04", "Desktop & Mobile UI vs. Figma Designs", "p. 5"),
        ("05", "Assets, Images & External Credits", "p. 6"),
        ("06", "Technology Stack & JS Interactions", "p. 7"),
        ("07", "GitHub Repository & Live Deployment", "p. 8"),
        ("08", "Technical Learning Reflection", "p. 9"),
        ("09", "Acceptance Tests Matrix (T01 – T06)", "p. 10"),
        ("10", "Grading Rubric & Print Checklist", "p. 11–12"),
    ]
    t_toc = doc.add_table(rows=len(toc_rows), cols=3)
    t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (num, title, pg) in enumerate(toc_rows):
        row = t_toc.rows[i]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(0.55), Inches(4.1), Inches(0.75)
        bg = "F8FAFC" if i % 2 == 0 else "FFFFFF"
        for c in (c0, c1, c2):
            set_cell_bg(c, bg)
            set_cell_margins(c, 35, 35, 60, 60)
            set_cell_borders(c, "E2E8F0", "4")
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(num)
        r0.font.name, r0.font.size, r0.bold, r0.font.color.rgb = "Segoe UI", Pt(9.5), True, C_PURPLE

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(title)
        r1.font.name, r1.font.size, r1.font.color.rgb = "Segoe UI", Pt(9.5), C_DARK

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(pg)
        r2.font.name, r2.font.size, r2.bold, r2.font.color.rgb = "Segoe UI", Pt(9.5), True, C_NAVY

    add_subhead("1. Developer Profile")
    add_body(
        "I am Maribel Bravo (ID: 0224-0865), a 3rd-year BSIT student majoring in Web and Mobile Application Development (BSIT 3A WMAD) at Laguna State Polytechnic University – Siniloan Campus."
    )
    add_body(
        "HTML5, CSS3, Bootstrap 5, JavaScript (ES6+), Python (Flask), Java, SQL, PostgreSQL, React, and Git/GitHub.",
        bold_prefix="Explainable Skills:"
    )
    add_body(
        "moralesmaribel262@gmail.com  |  github.com/hiddengemofbell  |  linkedin.com/in/maribel-bravo-876655374",
        bold_prefix="Professional Links:"
    )

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 3: PORTFOLIO DESCRIPTION & PARTS
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 02", "Portfolio Description & Main Parts")
    add_body(
        "This portfolio presents my academic background, technical skills, and software projects to IT instructors, internship recruiters, and fellow developers."
    )

    parts_grid = [
        ("Header & Sticky Navbar", "Brand monogram 'MB' and responsive navigation links that collapse into a mobile drawer on narrow screens."),
        ("Hero Introduction", "Dynamic typewriter headline, availability status indicator, CV download button, and direct social links."),
        ("About Me & Highlights", "Academic profile (BSIT 3A WMAD, LSPU), optimized portrait photo, and 3 specialization highlight cards."),
        ("Technical Skills", "Categorized skill pills separating core languages, frontend styling, tools, and AI dev environments."),
        ("Featured Projects (WEB 6)", "Real-world works: Cake Ordering Platform, Queue Ordering System, and EduFlow Task Manager PWA."),
        ("Certificates & Contact", "Verified cybersecurity/IT credentials plus an interactive Formspree contact form with live feedback."),
    ]

    t_parts = doc.add_table(rows=3, cols=2)
    t_parts.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, (pt_title, pt_desc) in enumerate(parts_grid):
        r, c = divmod(idx, 2)
        cell = t_parts.rows[r].cells[c]
        cell.width = Inches(2.7)
        set_cell_bg(cell, "F8FAFC")
        set_cell_margins(cell, 70, 70, 90, 90)
        set_cell_borders(cell, "CBD5E1", "4")
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.12
        rt = p.add_run(pt_title + "\n")
        rt.font.name, rt.font.size, rt.bold, rt.font.color.rgb = "Segoe UI", Pt(10.5), True, C_PURPLE
        rd = p.add_run(pt_desc)
        rd.font.name, rd.font.size, rd.font.color.rgb = "Segoe UI", Pt(9.5), C_DARK

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 4: LOGO & COLOR PALETTE
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 03", "Personal Logo & Color Palette")

    add_callout_card(
        "Personal Monogram Wordmark: 'MB' + Cyan Beacon",
        "The navbar and browser favicon feature the bold 'MB' initials paired with an electric-cyan dot (#06B6D4), symbolizing availability and modern web craftsmanship."
    )

    add_subhead("Color Palette & HEX Application Table")
    palette_rows = [
        ("Swatch", "Token & HEX", "Where Applied in Portfolio"),
        ("060610", "--bg-primary\n#060610", "Root page background & dark canvas base"),
        ("0D0D1F", "--bg-secondary\n#0D0D1F", "Alternating section depth & subtle gradients"),
        ("7C3AED", "--purple\n#7C3AED", "Primary CTA buttons, brand mark, active accents"),
        ("06B6D4", "--cyan\n#06B6D4", "Focus-visible outline, scroll progress bar, status dot"),
        ("F8FAFC", "--text-primary\n#F8FAFC", "Main headings & primary copy (18.2:1 contrast)"),
        ("94A3B8", "--text-secondary\n#94A3B8", "Supporting paragraphs, labels & card metadata"),
    ]
    t_pal = doc.add_table(rows=len(palette_rows), cols=3)
    t_pal.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (sw, tok, where) in enumerate(palette_rows):
        row = t_pal.rows[i]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(0.7), Inches(1.6), Inches(3.1)
        for c in (c0, c1, c2):
            set_cell_margins(c, 50, 50, 70, 70)
            set_cell_borders(c, "CBD5E1", "4")
        if i == 0:
            for c, txt in zip((c0, c1, c2), (sw, tok, where)):
                set_cell_bg(c, "0F172A")
                p = c.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(txt)
                r.font.name, r.font.size, r.bold, r.font.color.rgb = "Segoe UI", Pt(9.5), True, C_WHITE
        else:
            set_cell_bg(c0, sw)
            c0.paragraphs[0].paragraph_format.space_after = Pt(0)
            set_cell_bg(c1, "F8FAFC")
            set_cell_bg(c2, "FFFFFF")
            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_after = Pt(0)
            r1 = p1.add_run(tok)
            r1.font.name, r1.font.size, r1.bold, r1.font.color.rgb = "Segoe UI", Pt(9), True, C_NAVY
            p2 = c2.paragraphs[0]
            p2.paragraph_format.space_after = Pt(0)
            r2 = p2.add_run(where)
            r2.font.name, r2.font.size, r2.font.color.rgb = "Segoe UI", Pt(9.5), C_DARK

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 5: DESKTOP & MOBILE FIGMA DESIGNS
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 04", "Desktop & Mobile Designs (Figma Alignment)")
    add_body("Figma Prototype Link: https://www.figma.com/@hiddengemofbell", bold_prefix="Figma URL:")

    desk_path = r"C:\Users\moral\OneDrive\Practice\basic-portfolio\assets\screenshots\desktop.png"
    mob_path = r"C:\Users\moral\OneDrive\Practice\basic-portfolio\assets\screenshots\mobile.png"

    t_scr = doc.add_table(rows=1, cols=2)
    t_scr.alignment = WD_TABLE_ALIGNMENT.CENTER
    cd, cm = t_scr.rows[0].cells[0], t_scr.rows[0].cells[1]
    cd.width, cm.width = Inches(3.45), Inches(1.95)
    for c in (cd, cm):
        set_cell_bg(c, "F8FAFC")
        set_cell_margins(c, 50, 50, 50, 50)
        set_cell_borders(c, "CBD5E1", "4")

    pd = cd.paragraphs[0]
    pd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(desk_path):
        pd.add_run().add_picture(desk_path, width=Inches(3.25))
    pd_cap = cd.add_paragraph()
    pd_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pd_cap.paragraph_format.space_after = Pt(0)
    rdc = pd_cap.add_run("Fig 1. Desktop Layout (1280px)")
    rdc.font.name, rdc.font.size, rdc.bold, rdc.font.color.rgb = "Segoe UI", Pt(9), True, C_NAVY

    pm = cm.paragraphs[0]
    pm.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(mob_path):
        pm.add_run().add_picture(mob_path, width=Inches(1.65))
    pm_cap = cm.add_paragraph()
    pm_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pm_cap.paragraph_format.space_after = Pt(0)
    rmc = pm_cap.add_run("Fig 2. Mobile Layout (375px)")
    rmc.font.name, rmc.font.size, rmc.bold, rmc.font.color.rgb = "Segoe UI", Pt(9), True, C_NAVY

    add_subhead("Design-to-Implementation Mapping")
    add_body("Desktop multi-column project/contact grids collapse cleanly into single-column cards at 375px, while the horizontal navbar switches to a Bootstrap 5 toggle drawer.")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 6: ASSETS & IMAGES
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 05", "Assets, Images & External Credits")
    add_body("All visual assets are either original student work or properly credited open-source resources:")

    ast_rows = [
        ("Asset File", "Origin / Credit", "Purpose & Optimization"),
        ("assets/profile.jpg", "Original photo (Maribel Bravo)", "Developer portrait; compressed from 6.6 MB to 68 KB (99% smaller)."),
        ("favicon.ico & favicon.svg", "Original vector design", "MB monogram tab icon & mobile home-screen shortcut."),
        ("assets/maribel-bravo-cv.pdf", "Original authored CV", "Downloadable resume linked from the Hero CTA."),
        ("Bootstrap 5.3.3 & Icons", "Bootstrap Team (MIT)", "Responsive grid system, navbar collapse, and UI icons."),
        ("Inter Font Family", "Google Fonts (SIL OFL)", "Clean sans-serif typography across all viewports."),
        ("Formspree API", "Formspree Inc.", "Handles Contact Me form submissions asynchronously."),
    ]
    t_ast = doc.add_table(rows=len(ast_rows), cols=3)
    t_ast.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (f, o, p_txt) in enumerate(ast_rows):
        row = t_ast.rows[i]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(1.55), Inches(1.55), Inches(2.3)
        for c in (c0, c1, c2):
            set_cell_margins(c, 45, 45, 60, 60)
            set_cell_borders(c, "CBD5E1", "4")
        bg = "0F172A" if i == 0 else ("F8FAFC" if i % 2 == 0 else "FFFFFF")
        for idx_c, (c, txt) in enumerate(zip((c0, c1, c2), (f, o, p_txt))):
            set_cell_bg(c, bg)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(txt)
            r.font.name = "Segoe UI"
            r.font.size = Pt(9.5)
            if i == 0:
                r.bold, r.font.color.rgb = True, C_WHITE
            elif idx_c == 0:
                r.bold, r.font.color.rgb = True, C_PURPLE
            else:
                r.font.color.rgb = C_DARK

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 7: TECHNOLOGY & INTERACTION
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 06", "Technology Stack & JS Interactions")
    add_body(
        "Site Build Tools: Semantic HTML5, CSS3 Custom Properties, Bootstrap 5.3.3, and Vanilla ES6+ JavaScript (distinct from backend languages listed in my profile such as Python, Java, and PostgreSQL)."
    )

    js_rows = [
        ("Feature", "Trigger / Event", "Visible Result"),
        ("Typewriter Hero", "DOMContentLoaded + setTimeout", "Cycles career roles letter-by-letter in hero subtitle."),
        ("Scroll Progress Bar", "window 'scroll' event", "Top #scroll-progress bar fills 0% to 100% width."),
        ("ScrollSpy Navbar", "IntersectionObserver", "Highlights active section link while scrolling."),
        ("Scroll Reveal", "IntersectionObserver (12%)", "Cards fade and slide upward into view smoothly."),
        ("Contact Form", "form 'submit' (fetch POST)", "Shows loading spinner, sends to Formspree, prints confirmation."),
    ]
    t_js = doc.add_table(rows=len(js_rows), cols=3)
    t_js.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (ft, tr, res) in enumerate(js_rows):
        row = t_js.rows[i]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.width, c1.width, c2.width = Inches(1.4), Inches(1.65), Inches(2.35)
        for c in (c0, c1, c2):
            set_cell_margins(c, 45, 45, 60, 60)
            set_cell_borders(c, "CBD5E1", "4")
        bg = "0F172A" if i == 0 else ("F8FAFC" if i % 2 == 0 else "FFFFFF")
        for idx_c, (c, txt) in enumerate(zip((c0, c1, c2), (ft, tr, res))):
            set_cell_bg(c, bg)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(txt)
            r.font.name, r.font.size = "Segoe UI", Pt(9.5)
            if i == 0:
                r.bold, r.font.color.rgb = True, C_WHITE
            elif idx_c == 0:
                r.bold, r.font.color.rgb = True, C_NAVY
            else:
                r.font.color.rgb = C_DARK

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 8: REPOSITORY & DEPLOYED WEBSITE
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 07", "GitHub Repository & Live Deployment")

    add_callout_card(
        "Public Links (Verified Accessible)",
        "Deployed Site: https://hiddengemofbell.github.io/basic-portfolio/\n"
        "GitHub Repo:   https://github.com/hiddengemofbell/basic-portfolio"
    )

    add_subhead("How to Open & Locate Main Files")
    add_body("index.html — Semantic HTML5 entry page, SEO meta tags, and ARIA landmarks.", bold_prefix="•")
    add_body("css/style.css — Design tokens (:root), responsive breakpoints, and :focus-visible styles.", bold_prefix="•")
    add_body("js/main.js — Typewriter, scroll progress, ScrollSpy, reveal, and Formspree AJAX.", bold_prefix="•")
    add_body("assets/ — Optimized profile.jpg (68 KB), favicon.svg, and maribel-bravo-cv.pdf.", bold_prefix="•")
    add_body("README.md — Full project overview, Figma link, and acceptance testing summary.", bold_prefix="•")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 9: LEARNING REFLECTION
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 08", "Learning Reflection")

    add_callout_card(
        "Challenge: Balancing Dark Aesthetics with Speed & Accessibility",
        "My initial high-resolution portrait (6.6 MB) slowed mobile page loads, and default browser focus outlines were hard to see against dark glassmorphic cards."
    )

    add_subhead("How I Addressed It")
    add_body("I compressed the portrait image by 99% (down to 68 KB), added a keyboard skip-to-content link, wrapped all sections inside <main>, and styled a high-contrast 2px cyan :focus-visible ring across all interactive elements.")

    add_subhead("What I Learned")
    add_body("Good web design is not only visual styling—it requires semantic HTML5 structure, fast asset delivery, and keyboard accessibility so every visitor can navigate the site reliably.")

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 10: ACCEPTANCE TESTS (T01 - T06)
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Section 09", "Acceptance Tests Matrix (T01 – T06)")
    tests_rows = [
        ("ID", "Action & Screen Size", "Expected & Actual Result", "Status"),
        ("T01", "Private browser (Desktop)", "Site loads with styles & images without login.", "PASS"),
        ("T02", "Click all nav & project links", "All section anchors & GitHub links match labels.", "PASS"),
        ("T03", "375px, 768px, 1280px widths", "Usable nav, fitting images, no horizontal scroll.", "PASS"),
        ("T04", "Trigger JS features repeatedly", "Typewriter, progress bar & form work reliably.", "PASS"),
        ("T05", "Tab / Shift+Tab keyboard nav", "Visible cyan focus ring, skip link & alt text.", "PASS"),
        ("T06", "GitHub, Figma & Booklet check", "Repo, README, Figma & 12-p booklet verified.", "PASS"),
    ]
    t_tst = doc.add_table(rows=len(tests_rows), cols=4)
    t_tst.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (tid, act, res, st) in enumerate(tests_rows):
        row = t_tst.rows[i]
        c0, c1, c2, c3 = row.cells[0], row.cells[1], row.cells[2], row.cells[3]
        c0.width, c1.width, c2.width, c3.width = Inches(0.55), Inches(1.8), Inches(2.4), Inches(0.65)
        for c in (c0, c1, c2, c3):
            set_cell_margins(c, 45, 45, 50, 50)
            set_cell_borders(c, "CBD5E1", "4")
        bg = "0F172A" if i == 0 else ("F8FAFC" if i % 2 == 0 else "FFFFFF")
        for idx_c, (c, txt) in enumerate(zip((c0, c1, c2, c3), (tid, act, res, st))):
            set_cell_bg(c, bg)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(txt)
            r.font.name, r.font.size = "Segoe UI", Pt(9.5)
            if i == 0:
                r.bold, r.font.color.rgb = True, C_WHITE
            elif idx_c == 3:
                r.bold, r.font.color.rgb = True, RGBColor(22, 101, 52)
            elif idx_c == 0:
                r.bold, r.font.color.rgb = True, C_PURPLE
            else:
                r.font.color.rgb = C_DARK

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 11: RUBRIC & TABLE OF SPECIFICATIONS
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Appendix A", "Evaluation Rubric & Specifications")
    rub_rows = [
        ("Criterion", "Evidence Evaluated", "Max", "Score"),
        ("Subject Matter", "Semantic HTML5, CSS, Bootstrap 5, JS event/result (WEB 1–4).", "25", "___"),
        ("Creativity", "Personal MB logo, palette, Figma hierarchy & UI assets.", "25", "___"),
        ("Final Product", "Published site, navigation, responsive views, JS checks (T01–05).", "25", "___"),
        ("Structure & Flow", "README, GitHub history, Figma link, folded manual & PDF (T06).", "25", "___"),
        ("TOTAL", "Combined Final Project Rating", "100", "___ / 100"),
    ]
    t_rub = doc.add_table(rows=len(rub_rows), cols=4)
    t_rub.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (cr, ev, mx, sc) in enumerate(rub_rows):
        row = t_rub.rows[i]
        c0, c1, c2, c3 = row.cells[0], row.cells[1], row.cells[2], row.cells[3]
        c0.width, c1.width, c2.width, c3.width = Inches(1.25), Inches(2.75), Inches(0.55), Inches(0.85)
        for c in (c0, c1, c2, c3):
            set_cell_margins(c, 50, 50, 55, 55)
            set_cell_borders(c, "CBD5E1", "4")
        bg = "0F172A" if i == 0 else ("E2E8F0" if i == len(rub_rows) - 1 else ("F8FAFC" if i % 2 == 0 else "FFFFFF"))
        for idx_c, (c, txt) in enumerate(zip((c0, c1, c2, c3), (cr, ev, mx, sc))):
            set_cell_bg(c, bg)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(txt)
            r.font.name, r.font.size = "Segoe UI", Pt(9.5)
            if i == 0:
                r.bold, r.font.color.rgb = True, C_WHITE
            elif i == len(rub_rows) - 1 or idx_c == 0:
                r.bold, r.font.color.rgb = True, C_NAVY
            else:
                r.font.color.rgb = C_DARK

    add_subhead("Evaluator Sign-Off")
    add_body("Instructor: Prof. Julius A. Derla    Date: _______________", space_after=2)

    doc.add_page_break()

    # ══════════════════════════════════════════════════════════════════════
    # PAGE 12: BACK COVER / BOOKLET PRINT CHECKLIST
    # ══════════════════════════════════════════════════════════════════════
    add_section_banner("Back Cover", "IPT2 Booklet Format & Print Checklist")

    checks = [
        ("Paper Size:", "Long bond paper (8.5 × 13 inches), printed duplex (both sides)."),
        ("Fold & Size:", "Landscape sheet folded in half → 6.5 × 8.5 inch portrait booklet."),
        ("Imposition:", "12 pages total (multiple of 4); reads Pages 1–12 sequentially."),
        ("Typography:", "11 pt readable body text, 0.5–0.6 in margins clear of the fold."),
        ("Assembly:", "Align all 3 sheets, fold at center, and staple along the spine."),
    ]
    for k, v in checks:
        add_body(v, bold_prefix="[✔] " + k, space_after=4)

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(40)
    re1 = p_end.add_run("Maribel Bravo  •  BSIT 3A WMAD\n")
    re1.font.name, re1.font.size, re1.bold, re1.font.color.rgb = "Segoe UI", Pt(11), True, C_PURPLE
    re2 = p_end.add_run("Laguna State Polytechnic University – Siniloan Campus\nCollege of Computer Studies  •  2026")
    re2.font.name, re2.font.size, re2.font.color.rgb = "Segoe UI", Pt(9.5), C_GRAY

    out_dl = r"C:\Users\moral\Downloads\ITST_301_Personal_Portfolio_Manual_Maribel_Bravo.docx"
    out_repo = r"C:\Users\moral\OneDrive\Practice\basic-portfolio\ITST_301_Personal_Portfolio_Manual_Maribel_Bravo.docx"
    doc.save(out_dl)
    doc.save(out_repo)

    # Export sequential PDF via Word COM
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc_com = word.Documents.Open(os.path.abspath(out_dl))
    pdf_dl = r"C:\Users\moral\Downloads\ITST_301_Personal_Portfolio_Manual_Maribel_Bravo.pdf"
    pdf_repo = r"C:\Users\moral\OneDrive\Practice\basic-portfolio\ITST_301_Personal_Portfolio_Manual_Maribel_Bravo.pdf"
    doc_com.SaveAs(os.path.abspath(pdf_dl), FileFormat=17)
    doc_com.SaveAs(os.path.abspath(pdf_repo), FileFormat=17)
    doc_com.Close()
    word.Quit()

    # Impose onto 8.5 x 13 inch Landscape sheets (13 in wide x 8.5 in tall)
    impose_booklet(
        pdf_dl,
        r"C:\Users\moral\Downloads\ITST_301_Brochure_Booklet_PRINT_READY_8.5x13.pdf",
    )
    impose_booklet(
        pdf_dl,
        r"C:\Users\moral\OneDrive\Practice\basic-portfolio\ITST_301_Brochure_Booklet_PRINT_READY_8.5x13.pdf",
    )


def impose_booklet(src_pdf_path, dst_pdf_path):
    """Impose sequential 6.5x8.5 pages onto 13x8.5 landscape sheets for duplex center-fold."""
    reader = PdfReader(src_pdf_path)
    n = len(reader.pages)
    # Pad to multiple of 4 if needed
    pad = (4 - (n % 4)) % 4
    total = n + pad

    writer = PdfWriter()
    sheet_w = 13.0 * 72.0   # 936 pt
    sheet_h = 8.5 * 72.0    # 612 pt
    half_w = 6.5 * 72.0     # 468 pt

    # Saddle-stitch booklet pairs: (left_idx, right_idx) 0-based
    for i in range(total // 4):
        # Front of sheet: [total - 1 - 2*i, 2*i]
        # Back of sheet:  [2*i + 1, total - 2 - 2*i]
        pairs = [
            (total - 1 - 2 * i, 2 * i),
            (2 * i + 1, total - 2 - 2 * i),
        ]
        for left_i, right_i in pairs:
            sheet = PageObject.create_blank_page(width=sheet_w, height=sheet_h)
            if left_i < n:
                sheet.merge_transformed_page(
                    reader.pages[left_i],
                    Transformation().translate(0, 0),
                )
            if right_i < n:
                sheet.merge_transformed_page(
                    reader.pages[right_i],
                    Transformation().translate(half_w, 0),
                )
            writer.add_page(sheet)

    with open(dst_pdf_path, "wb") as f:
        writer.write(f)
    print(f"Imposed print-ready 13x8.5 booklet PDF ({len(writer.pages)} sides / {len(writer.pages)//2} sheets): {dst_pdf_path}")


if __name__ == "__main__":
    create_brochure_manual()
