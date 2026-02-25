from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Colour palette ────────────────────────────────────────────────────────────
EVIDENT_BLUE   = RGBColor(0x00, 0x5B, 0xAA)   # Evident brand blue
SECTION_BLUE   = RGBColor(0xE8, 0xF1, 0xFA)   # light blue cell shading
HEADER_DARK    = RGBColor(0x00, 0x3A, 0x70)   # dark navy for table headers
WHITE          = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT      = RGBColor(0x44, 0x44, 0x44)
NOTE_BG        = RGBColor(0xFF, 0xF8, 0xE1)   # soft amber for writer notes
NOTE_BORDER    = RGBColor(0xFF, 0xB3, 0x00)

# ── Style helpers ─────────────────────────────────────────────────────────────
def set_cell_bg(cell, rgb: RGBColor):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "{:02X}{:02X}{:02X}".format(*rgb))
    tcPr.append(shd)

def set_cell_border(cell, **edges):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge, attrs in edges.items():
        el = OxmlElement(f"w:{edge}")
        for k, v in attrs.items():
            el.set(qn(f"w:{k}"), v)
        tcBorders.append(el)
    tcPr.append(tcBorders)

def para_space(para, before=0, after=0):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)

def add_heading(doc, text, level=1, color=EVIDENT_BLUE):
    p   = doc.add_heading(text, level=level)
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.color.rgb = color
    run.font.bold      = True
    font_sizes = {1: 18, 2: 14, 3: 12}
    run.font.size = Pt(font_sizes.get(level, 11))
    para_space(p, before=14 if level == 1 else 10, after=4)
    return p

def add_body(doc, text, italic=False, color=GREY_TEXT, indent=False):
    p   = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size      = Pt(10.5)
    run.font.color.rgb = color
    run.font.italic    = italic
    para_space(p, before=2, after=4)
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    return p

def add_bullet(doc, text, level=0, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.font.bold      = True
        r.font.size      = Pt(10.5)
        r.font.color.rgb = GREY_TEXT
        r2 = p.add_run(text)
        r2.font.size      = Pt(10.5)
        r2.font.color.rgb = GREY_TEXT
    else:
        r = p.add_run(text)
        r.font.size      = Pt(10.5)
        r.font.color.rgb = GREY_TEXT
    p.paragraph_format.left_indent   = Cm(0.5 + level * 0.5)
    p.paragraph_format.space_after   = Pt(2)
    p.paragraph_format.space_before  = Pt(2)
    return p

def add_writer_note(doc, text):
    """Amber callout box for writer-facing instructions."""
    p   = doc.add_paragraph()
    run = p.add_run("✏  Writer note: " + text)
    run.font.size      = Pt(9.5)
    run.font.italic    = True
    run.font.color.rgb = RGBColor(0x7A, 0x5C, 0x00)
    p.paragraph_format.left_indent  = Cm(0.4)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    # light amber shading via direct XML (paragraph-level shading not standard in python-docx)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "FFF8E1")
    pPr.append(shd)
    return p

def make_table(doc, headers, rows, col_widths=None):
    """Generic styled table. headers = list of str, rows = list of lists."""
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style          = "Table Grid"
    t.alignment      = WD_TABLE_ALIGNMENT.LEFT
    # header row
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        set_cell_bg(cell, HEADER_DARK)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p   = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(h)
        run.font.bold      = True
        run.font.color.rgb = WHITE
        run.font.size      = Pt(9.5)
    # data rows
    for ri, row in enumerate(rows):
        tr = t.rows[ri + 1]
        bg = SECTION_BLUE if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row):
            cell = tr.cells[ci]
            set_cell_bg(cell, bg)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p   = cell.paragraphs[0]
            # support **bold** within cell text
            if "**" in str(val):
                parts = str(val).split("**")
                for idx, part in enumerate(parts):
                    if part:
                        run           = p.add_run(part)
                        run.font.bold = (idx % 2 == 1)
                        run.font.size = Pt(9.5)
                        run.font.color.rgb = GREY_TEXT
            else:
                run = p.add_run(str(val))
                run.font.size      = Pt(9.5)
                run.font.color.rgb = GREY_TEXT
    # optional column widths
    if col_widths:
        for ci, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[ci].width = Cm(w)
    doc.add_paragraph()  # spacer
    return t

def add_faq(doc, q, a, link_text=None):
    # Question
    p = doc.add_paragraph()
    r = p.add_run("Q: " + q)
    r.font.bold      = True
    r.font.size      = Pt(10.5)
    r.font.color.rgb = EVIDENT_BLUE
    para_space(p, before=8, after=2)
    # Answer
    p2 = doc.add_paragraph()
    r2 = p2.add_run(a)
    r2.font.size      = Pt(10.5)
    r2.font.color.rgb = GREY_TEXT
    para_space(p2, before=0, after=2)
    if link_text:
        p3 = doc.add_paragraph()
        r3 = p3.add_run("→ " + link_text)
        r3.font.italic     = True
        r3.font.size       = Pt(10)
        r3.font.color.rgb  = EVIDENT_BLUE
        para_space(p3, before=0, after=6)

def add_divider(doc):
    p   = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pb  = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"),   "single")
    bot.set(qn("w:sz"),    "6")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "005BAA")
    pb.append(bot)
    pPr.append(pb)
    para_space(p, before=6, after=6)

# ══════════════════════════════════════════════════════════════════════════════
#  DOCUMENT START
# ══════════════════════════════════════════════════════════════════════════════

# ── Cover block ───────────────────────────────────────────────────────────────
title = doc.add_heading("Content Brief: Research Slide Scanners", 0)
for run in title.runs:
    run.font.color.rgb = EVIDENT_BLUE
    run.font.size      = Pt(22)
title.alignment = WD_ALIGN_PARAGRAPH.LEFT
para_space(title, before=0, after=4)

sub = doc.add_paragraph()
r   = sub.add_run("Life Science Category Page  ·  Evident Scientific")
r.font.size      = Pt(11)
r.font.italic    = True
r.font.color.rgb = GREY_TEXT
para_space(sub, before=0, after=12)

add_divider(doc)

# ── Meta block ────────────────────────────────────────────────────────────────
add_heading(doc, "1.  SEO Tags", level=1)
add_writer_note(doc, "Pass these tags to your web/CMS team once copy is approved. Do not publish without updating the title tag — the current one still reads '| Olympus' which is outdated.")

make_table(doc,
    headers=["Tag", "Value"],
    rows=[
        ["Title Tag",        "Research Slide Scanners for Life Science | Evident Scientific"],
        ["Meta Description", "Digitize entire microscope slides with the SLIDEVIEW VS200 — Evident Scientific's whole slide imaging system for neuroscience, cancer, spatial biology, and drug discovery research.  (152 chars)"],
        ["H1",               "High-Resolution Whole Slide Imaging for Life Science Research"],
        ["URL Slug",         "/en/life-science-microscopes/slide-scanners  (keep as-is)"],
        ["Schema",           "FAQ schema on FAQ section — dev team to implement JSON-LD"],
    ],
    col_widths=[4.5, 12.0],
)

# ── Tone block ────────────────────────────────────────────────────────────────
add_heading(doc, "2.  Tone & Voice", level=1)
bullets = [
    ("Audience: ", "Life science researchers and core facility managers. PhD-level, comfortable with technical terms — not a pathologist or clinician."),
    ("Brand: ",    "Use 'Evident Scientific' throughout. Never write 'Olympus' on this page."),
    ("Framing: ",  "Research only. Do not write 'diagnosis,' 'patient,' or 'pathologist.' Those belong on the clinical page."),
    ("Style: ",    "Confident, specific. Avoid vague superlatives ('industry-leading'). Lead with capabilities and applications."),
    ("Format: ",   "Active voice. Short paragraphs (3–5 sentences max). Use the heading structure below exactly as written."),
]
for bold, rest in bullets:
    add_bullet(doc, rest, bold_prefix=bold)
doc.add_paragraph()

add_divider(doc)

# ── Page Structure ────────────────────────────────────────────────────────────
add_heading(doc, "3.  Page Structure & Writing Guide", level=1)
add_body(doc, "Write each section in order. Word count targets are per-section minimums — you may run slightly over but do not go under. Total target: ~1,400 words.")

# ─── INTRO ───────────────────────────────────────────────────────────────────
add_heading(doc, "Intro / Hero  [~100 words]", level=2)
add_writer_note(doc, "Hook on the research pain point — don't open with a product pitch. Position WSI as the solution, VS200 as the tool.")
add_body(doc, "Cover these points:")
add_bullet(doc, "Manually reviewing slides one field of view at a time slows research, limits reproducibility, and makes sharing results difficult.")
add_bullet(doc, "Whole slide imaging (WSI) digitizes entire slides into high-resolution virtual images — enabling batch scanning, computational analysis, and remote collaboration.")
add_bullet(doc, "The SLIDEVIEW VS200 is Evident Scientific's research slide scanner, built for multi-modal life science workflows from single labs to shared core facilities.")

# ─── H2 1 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: What Is a Research Slide Scanner?  [~150 words]", level=2)
add_writer_note(doc, "Define the technology for researchers who may be new to WSI. Differentiate from both a standard microscope and a clinical scanner.")
add_body(doc, "Cover these points:")
add_bullet(doc, "A research slide scanner (also called a whole slide imaging system) automatically captures high-resolution images across the entire slide and stitches them into a single navigable digital file — a virtual slide.")
add_bullet(doc, "Virtual slides can be viewed, shared, and analyzed without physical access to the glass slide.")
add_bullet(doc, "Research scanners differ from clinical/digital pathology scanners in one key way: they prioritize imaging flexibility — supporting fluorescence, darkfield, phase contrast, and polarization — rather than standardized high-speed brightfield scanning.")
add_bullet(doc, "INTERNAL LINK: at the end of this section, add — 'Looking for a clinical digital pathology scanner? See the SLIDEVIEW DX →' — linking to /en/clinical-microscopes/slide-scanners")

# ─── H2 2 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: Research Applications for Whole Slide Imaging  [~300 words total]", level=2)
add_writer_note(doc, "This is the most important section on the page. Do not write it generically. Each H3 should read like it was written for that specific research community.")
add_body(doc, "Write a 1–2 sentence intro before the H3s: state that the VS200 is configured for applications that demand imaging flexibility, high throughput, and quantitative analysis — from neuroscience to drug discovery.")

add_heading(doc, "H3: Neuroscience & Brain Research  [~60 words]", level=3)
add_bullet(doc, "Large-format brain section digitization — VS200 handles oversized slides and stitches multi-image acquisitions seamlessly.")
add_bullet(doc, "Polarized light imaging reveals amyloid structures in brain tissue.")
add_bullet(doc, "Cleared tissue imaging via the VS-SILA optical sectioning module.")
add_writer_note(doc, "Link to brain/neuroscience application note when URL is confirmed.")

add_heading(doc, "H3: Cancer Research & Tumor Microenvironment  [~60 words]", level=3)
add_bullet(doc, "Multiplexed fluorescence imaging for biomarker co-localization in tissue sections.")
add_bullet(doc, "IHC/FISH quantification at scale — scan hundreds of slides unattended.")
add_bullet(doc, "TruAI deep-learning tools for automated cell detection and morphological measurements in tumor tissue.")
add_writer_note(doc, "Link to cancer/oncology application note when URL is confirmed.")

add_heading(doc, "H3: Spatial Biology  [~60 words]", level=3)
add_bullet(doc, "Image multiple fluorescence-labeled biomarkers simultaneously to map cell types and interactions within tissue.")
add_bullet(doc, "VS200 supports up to seven fluorescence channels for multiplex immunofluorescence workflows.")
add_bullet(doc, "Compatible with spatial transcriptomics and proteomics sample preparations.")
add_writer_note(doc, "INTERNAL LINK: 'spatial biology' → /en/insights/how-spatial-biology-is-transforming-and-innovating-research")

add_heading(doc, "H3: Drug Discovery & Preclinical Research  [~60 words]", level=3)
add_bullet(doc, "High-throughput IHC and FISH quantification across large compound libraries.")
add_bullet(doc, "Phenotypic screening of tissue-level compound effects.")
add_bullet(doc, "Consistent, reproducible image acquisition across scan sessions and instruments — critical for cross-site preclinical studies.")
add_bullet(doc, "Integration with LIMS and data management systems via open file formats.")

add_heading(doc, "H3: Stem Cell & Plant Science  [~60 words]", level=3)
add_bullet(doc, "Automated batch scanning for colony counting and morphological phenotyping in stem cell research.")
add_bullet(doc, "Large-area plant tissue section imaging with brightfield and fluorescence.")
add_bullet(doc, "VS200's flexible slide loading accepts non-standard specimen preparations.")

# ─── H2 3 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: Imaging Modes — Brightfield, Fluorescence & Beyond  [~150 words]", level=2)
add_writer_note(doc, "Write a 1–2 sentence intro, then hand off to the table below. Ask the web team to render this as a styled table in the CMS.")
add_body(doc, "Intro direction: The SLIDEVIEW VS200 supports a wider range of imaging modes than any other slide scanner in its class — your scanner won't limit your research methods.")
doc.add_paragraph()
make_table(doc,
    headers=["Imaging Mode", "Research Use Case"],
    rows=[
        ["Brightfield",                       "H&E staining, IHC, routine histology"],
        ["Fluorescence (up to 7 channels)",   "Multiplexed biomarker imaging, FISH"],
        ["Darkfield",                         "Unlabeled sample contrast, nanoparticle detection"],
        ["Phase Contrast",                    "Label-free transparent sample imaging"],
        ["Polarization",                      "Amyloid detection, collagen fiber orientation"],
        ["Z-stacking",                        "Cytology, thick specimens, 3D reconstruction"],
        ["Optical Sectioning (VS-SILA)",      "Cleared tissue, >100 µm thick fluorescent samples"],
    ],
    col_widths=[7.0, 9.5],
)
add_body(doc, "Closing line: All modes are available on a single instrument — switch between them without moving your sample.")

# ─── H2 4 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: High-Throughput Scanning for Core Facilities  [~120 words]", level=2)
add_writer_note(doc, "Speak to core facility managers and high-volume research labs. Lead with the numbers.")
add_body(doc, "Cover these points:")
add_bullet(doc, "The VS200 loads up to 210 slides in a single automated session — with hot-swap functionality so you can add slides without interrupting an ongoing scan.")
add_bullet(doc, "Mixed slide formats (standard 25×75 mm and non-standard sizes) are detected and handled automatically in the same run.")
add_bullet(doc, "Overnight unattended scanning means the instrument works while your team doesn't.")
add_bullet(doc, "Images are immediately available for review via the NIS-SQL image management system — no post-processing queue.")
add_bullet(doc, "The VS200 supports multi-user queuing and remote image access, making it well suited for shared core facility environments.")

# ─── H2 5 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: AI-Guided Image Analysis with TruAI  [~120 words]", level=2)
add_writer_note(doc, "Position TruAI as built into the workflow — not an add-on. Emphasise that it reduces manual work at the scanning stage, not just at analysis.")
add_body(doc, "Cover these points:")
add_bullet(doc, "TruAI is Evident's deep-learning image analysis technology, integrated directly into the VS200 scanning workflow.")
add_bullet(doc, "Automatic sample detection: TruAI identifies tissue regions and restricts scanning to those areas — saving scan time and storage without sacrificing coverage.")
add_bullet(doc, "Cell detection and segmentation: identify and measure cells, nuclei, or morphological features without manual annotation.")
add_bullet(doc, "TruAI models can be trained on your own data for custom assays.")
add_bullet(doc, "Supports label-free observation — detect structural features without fluorescent markers.")
add_writer_note(doc, "INTERNAL LINK: 'TruAI' → link to TruAI product/technology page (confirm URL with web team).")

# ─── H2 6 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: How to Choose a Research Slide Scanner  [~200 words]", level=2)
add_writer_note(doc, "Write this as a practical self-qualification guide — plain, direct language. No marketing copy. This section captures high-intent 'how to choose' searches and is a major SEO opportunity for this page.")
add_body(doc, "Intro line: Every lab has different imaging needs. Use these questions to narrow down the right configuration for your research.")
doc.add_paragraph()
qa_rows = [
    ("Do you need fluorescence imaging?",
     "The VS200 starts as a brightfield system and can be field-upgraded to add fluorescence — no need to buy everything upfront."),
    ("How many slides do you scan per session?",
     "For low-volume labs, a smaller loader configuration may suffice. Core facilities running overnight batches should plan for the 210-slide loader with hot-swap."),
    ("Do you work with thick specimens or cleared tissue?",
     "Standard slide scanners focus on a single plane. For specimens thicker than ~15 µm, Z-stacking or the VS-SILA optical sectioning module is essential."),
    ("What magnification does your application require?",
     "20x is standard for histology and IHC. 40x or 60x (with oil immersion) is needed for cytology and subcellular-resolution work."),
    ("Do you need LIMS or data management integration?",
     "The VS200 exports DICOM and other standard open formats, compatible with NIS-SQL and third-party platforms."),
    ("Will multiple users share the system?",
     "The VS200 supports multi-user queuing and remote image access — well suited for core facility deployment."),
]
make_table(doc,
    headers=["Question", "Guidance for Researcher"],
    rows=qa_rows,
    col_widths=[7.5, 9.0],
)
add_body(doc, "Closing CTA line: Not sure which configuration fits your lab? Contact an Evident specialist — we'll help you spec the right system.")

add_divider(doc)

# ─── FAQ ─────────────────────────────────────────────────────────────────────
add_heading(doc, "4.  Frequently Asked Questions  [~250 words]", level=1)
add_writer_note(doc, "Format as an accordion FAQ in the CMS. Ask the dev team to add FAQ JSON-LD schema markup — this section is eligible for Google rich results.")

add_faq(doc,
    "What is a research slide scanner?",
    "A research slide scanner — also called a whole slide imaging (WSI) system — automatically digitizes entire microscope slides into high-resolution navigable virtual images. Unlike a standard microscope, it captures the full slide in one automated session, enabling batch processing, computational analysis, and remote image sharing."
)
add_faq(doc,
    "What is the difference between a research slide scanner and a digital pathology scanner?",
    "Research slide scanners are built for imaging flexibility — supporting fluorescence, darkfield, phase contrast, polarization, and Z-stacking alongside brightfield. Digital pathology scanners are optimized for high-speed, standardized brightfield scanning of H&E slides for clinical review. Evident offers both: the SLIDEVIEW VS200 for research and the SLIDEVIEW DX for digital pathology.",
    link_text="See the SLIDEVIEW DX for digital pathology → /en/clinical-microscopes/slide-scanners"
)
add_faq(doc,
    "Can the VS200 handle fluorescence and thick samples?",
    "Yes. The VS200 supports up to seven fluorescence channels. For thick specimens (100 µm or greater), the optional VS-SILA optical sectioning module eliminates out-of-focus light — delivering high-contrast fluorescence images of cleared tissue, thick brain sections, and 3D organoids without computational deconvolution."
)
add_faq(doc,
    "How many slides can the VS200 scan in one session?",
    "Up to 210 slides with the automated loader. Hot-swap functionality allows slides to be added mid-session without interrupting scanning. The system accepts mixed slide formats in a single run and detects format type automatically."
)
add_faq(doc,
    "What file formats does the VS200 produce?",
    "The VS200 outputs in standard open formats including DICOM and VSI, and is compatible with third-party image analysis platforms and laboratory information management systems (LIMS). Images are managed through Evident's NIS-SQL database or exported for use in external workflows."
)

add_divider(doc)

# ─── Internal Links ───────────────────────────────────────────────────────────
add_heading(doc, "5.  Internal Links — Pass to Web Team", level=1)
add_writer_note(doc, "These links should be embedded in the copy. Anchor text is specified — do not use 'click here' or the raw URL as the anchor.")
make_table(doc,
    headers=["Anchor Text (use exactly this)", "Destination URL"],
    rows=[
        ["digital pathology slide scanners",             "/en/clinical-microscopes/slide-scanners"],
        ["SLIDEVIEW VS200",                              "/en/products/slide-scanners/vs200"],
        ["spatial biology",                              "/en/insights/how-spatial-biology-is-transforming-and-innovating-research"],
        ["whole slide imaging FAQ",                      "/en/insights/whole-slide-imaging-faq"],
        ["VS200 for research applications",              "/en/applications/vs200-for-research"],
        ["virtual slide microscopy",                     "/en/solutions/virtual-slide-microscopy"],
        ["SLIDEVIEW DX for digital pathology",           "/en/products/slide-scanners/vs-m1"],
    ],
    col_widths=[8.5, 8.0],
)

# ─── Word Count ───────────────────────────────────────────────────────────────
add_heading(doc, "6.  Word Count Targets by Section", level=1)
make_table(doc,
    headers=["Section", "Target Words"],
    rows=[
        ["Intro / Hero",                          "100"],
        ["What Is a Research Slide Scanner?",     "150"],
        ["Research Applications (all H3s)",       "300"],
        ["Imaging Modes",                         "150"],
        ["High-Throughput / Core Facilities",     "120"],
        ["AI / TruAI",                            "120"],
        ["How to Choose",                         "200"],
        ["FAQ",                                   "250"],
        ["**Total**",                             "**~1,390**"],
    ],
    col_widths=[11.0, 5.5],
)

# ─── Save ─────────────────────────────────────────────────────────────────────
out = "/home/user/content-brief-generator/Evident_Research_Slide_Scanners_Content_Brief.docx"
doc.save(out)
print(f"Saved: {out}")
