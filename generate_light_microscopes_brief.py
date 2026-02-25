from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ── Colour palette ────────────────────────────────────────────────────────────
EVIDENT_BLUE  = RGBColor(0x00, 0x5B, 0xAA)
SECTION_BLUE  = RGBColor(0xE8, 0xF1, 0xFA)
HEADER_DARK   = RGBColor(0x00, 0x3A, 0x70)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT     = RGBColor(0x44, 0x44, 0x44)

# ── Style helpers ─────────────────────────────────────────────────────────────
def set_cell_bg(cell, rgb: RGBColor):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "{:02X}{:02X}{:02X}".format(*rgb))
    tcPr.append(shd)

def para_space(para, before=0, after=0):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)

def add_heading(doc, text, level=1, color=EVIDENT_BLUE):
    p   = doc.add_heading(text, level=level)
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.color.rgb = color
    run.font.bold      = True
    run.font.size      = Pt({1: 18, 2: 14, 3: 12}.get(level, 11))
    para_space(p, before=14 if level == 1 else 10, after=4)
    return p

def add_body(doc, text, italic=False, color=GREY_TEXT):
    p   = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size      = Pt(10.5)
    run.font.color.rgb = color
    run.font.italic    = italic
    para_space(p, before=2, after=4)
    return p

def add_bullet(doc, text, level=0, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = GREY_TEXT
        r2 = p.add_run(text)
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = GREY_TEXT
    else:
        r = p.add_run(text)
        r.font.size = Pt(10.5)
        r.font.color.rgb = GREY_TEXT
    p.paragraph_format.left_indent  = Cm(0.5 + level * 0.5)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    return p

def add_writer_note(doc, text):
    p   = doc.add_paragraph()
    run = p.add_run("✏  Writer note: " + text)
    run.font.size      = Pt(9.5)
    run.font.italic    = True
    run.font.color.rgb = RGBColor(0x7A, 0x5C, 0x00)
    p.paragraph_format.left_indent  = Cm(0.4)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "FFF8E1")
    pPr.append(shd)
    return p

def make_table(doc, headers, rows, col_widths=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style     = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
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
    for ri, row in enumerate(rows):
        tr = t.rows[ri + 1]
        bg = SECTION_BLUE if ri % 2 == 0 else WHITE
        for ci, val in enumerate(row):
            cell = tr.cells[ci]
            set_cell_bg(cell, bg)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            if "**" in str(val):
                parts = str(val).split("**")
                for idx, part in enumerate(parts):
                    if part:
                        run = p.add_run(part)
                        run.font.bold = (idx % 2 == 1)
                        run.font.size = Pt(9.5)
                        run.font.color.rgb = GREY_TEXT
            else:
                run = p.add_run(str(val))
                run.font.size      = Pt(9.5)
                run.font.color.rgb = GREY_TEXT
    if col_widths:
        for ci, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[ci].width = Cm(w)
    doc.add_paragraph()
    return t

def add_faq(doc, q, a, link_text=None):
    p = doc.add_paragraph()
    r = p.add_run("Q: " + q)
    r.font.bold      = True
    r.font.size      = Pt(10.5)
    r.font.color.rgb = EVIDENT_BLUE
    para_space(p, before=8, after=2)
    p2 = doc.add_paragraph()
    r2 = p2.add_run(a)
    r2.font.size      = Pt(10.5)
    r2.font.color.rgb = GREY_TEXT
    para_space(p2, before=0, after=2)
    if link_text:
        p3 = doc.add_paragraph()
        r3 = p3.add_run("→ " + link_text)
        r3.font.italic    = True
        r3.font.size      = Pt(10)
        r3.font.color.rgb = EVIDENT_BLUE
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
title = doc.add_heading("Content Brief: Light Microscopes for Material Science", 0)
for run in title.runs:
    run.font.color.rgb = EVIDENT_BLUE
    run.font.size      = Pt(22)
title.alignment = WD_ALIGN_PARAGRAPH.LEFT
para_space(title, before=0, after=4)

sub = doc.add_paragraph()
r   = sub.add_run("Material Science Category Page  ·  Evident Scientific")
r.font.size      = Pt(11)
r.font.italic    = True
r.font.color.rgb = GREY_TEXT
para_space(sub, before=0, after=12)

add_divider(doc)

# ── SEO Tags ──────────────────────────────────────────────────────────────────
add_heading(doc, "1.  SEO Tags", level=1)
add_writer_note(doc, "Pass these tags to your web/CMS team once copy is approved. Confirm current title tag — it may still contain 'Olympus' branding that needs to be replaced with 'Evident Scientific'.")

make_table(doc,
    headers=["Tag", "Value"],
    rows=[
        ["Title Tag",        "Light Microscopes for Material Science & Industrial Inspection | Evident Scientific"],
        ["Meta Description", "Explore Evident Scientific's upright and inverted light microscopes for metallurgy, failure analysis, and industrial QC. BX53M, GX53, and PRECiV software — built for precision material inspection.  (178 chars)"],
        ["H1",               "Light Microscopes for Material Science and Industrial Inspection"],
        ["URL Slug",         "/en/material-science-microscopes/light  (keep as-is)"],
        ["Schema",           "FAQ schema on FAQ section — dev team to implement JSON-LD"],
    ],
    col_widths=[4.5, 12.0],
)

# ── Tone block ────────────────────────────────────────────────────────────────
add_heading(doc, "2.  Tone & Voice", level=1)
bullets = [
    ("Audience: ",  "Materials scientists, metallurgists, QC engineers, failure analysis engineers, and R&D lab managers in manufacturing, automotive, steel, electronics, and academic research."),
    ("Brand: ",     "Use 'Evident Scientific' throughout. Never use 'Olympus' — the legacy brand name is retired on this page."),
    ("Framing: ",   "Industrial and materials science use cases only. This is not a life science or clinical page. Do not reference biological samples, cells, or clinical diagnosis."),
    ("Style: ",     "Precise, application-led. Lead with what the microscope enables the user to see or measure. Avoid vague phrases like 'state-of-the-art' or 'world-class.' Use specific numbers and standard references (ASTM, ISO, JIS) wherever possible."),
    ("Products: ",  "Three primary systems: BX53M (upright, reflected/transmitted), GX53 (inverted, large/heavy samples), DSX2000 (fully motorized all-in-one digital). Do not invent other model names or specs — reference only confirmed information."),
    ("Format: ",    "Active voice. Short paragraphs (3–5 sentences max). Use the heading structure below exactly as written."),
]
for bold, rest in bullets:
    add_bullet(doc, rest, bold_prefix=bold)
doc.add_paragraph()

add_divider(doc)

# ── Page Structure ────────────────────────────────────────────────────────────
add_heading(doc, "3.  Page Structure & Writing Guide", level=1)
add_body(doc, "Write each section in order. Word count targets are per-section minimums. Total target: ~1,500 words.")

# ─── INTRO ───────────────────────────────────────────────────────────────────
add_heading(doc, "Intro / Hero  [~100 words]", level=2)
add_writer_note(doc, "Open on the inspection challenge — not the product. Frame light microscopy as the standard tool for seeing what's inside a material, before recommending any specific system.")
add_body(doc, "Cover these points:")
add_bullet(doc, "Understanding material microstructure — grain boundaries, surface defects, coating integrity, phase distribution — is fundamental to quality control, failure analysis, and new material development.")
add_bullet(doc, "Light microscopy gives materials scientists and QC engineers a fast, non-destructive window into the surface and near-surface structure of metals, ceramics, semiconductors, minerals, and polymers.")
add_bullet(doc, "Evident Scientific's light microscopes for material science combine precision optics, versatile contrast techniques, and AI-powered software — from routine inspection to advanced quantitative analysis.")

# ─── H2 1 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: What Is a Material Science Light Microscope?  [~150 words]", level=2)
add_writer_note(doc, "Distinguish from biological/life science microscopes — the key difference is reflected (epi) illumination for opaque samples. Keep accessible; some readers may be setting up a new QC lab and comparing options.")
add_body(doc, "Cover these points:")
add_bullet(doc, "Material science light microscopes — also called metallurgical or industrial microscopes — use reflected (epi-illumination) light rather than transmitted light, because most material samples are opaque.")
add_bullet(doc, "This allows direct surface examination of polished metals, cross-sections, electronic components, minerals, and coatings without the need to make the sample transparent.")
add_bullet(doc, "They support multiple contrast methods — brightfield, darkfield, polarized light, DIC, and fluorescence — to reveal different types of surface features on the same sample.")
add_bullet(doc, "Evident offers both upright and inverted configurations to accommodate different sample sizes, weights, and workflow requirements.")
add_bullet(doc, "INTERNAL LINK: At the end of this section — 'Compare all Evident material science microscopes →' linking to /en/material-science-microscopes")

# ─── H2 2 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: Applications — What You Can Inspect and Measure  [~350 words total]", level=2)
add_writer_note(doc, "Most important section. Write each H3 as if addressing that specific user community. Use industry-specific vocabulary. Do not make this generic.")
add_body(doc, "Write a 1–2 sentence intro: Evident's light microscopes are used wherever precision surface and microstructure analysis matters — across metals, electronics, ceramics, minerals, and polymers.")

add_heading(doc, "H3: Metallurgy & Metals Inspection  [~70 words]", level=3)
add_bullet(doc, "Grain size analysis: measure and classify grain size in metals and alloys in compliance with ASTM E112, ISO 643, JIS G 0551, DIN 50601, and other international standards.")
add_bullet(doc, "Inclusion and porosity detection: identify non-metallic inclusions, voids, and porosity in steel and aluminum sections.")
add_bullet(doc, "Phase identification: distinguish different metallurgical phases using polarized light and DIC.")
add_bullet(doc, "Cast iron nodularity evaluation: classify graphite form, size, and distribution to ISO, ASTM, JIS, and GB/T standards using PRECiV Material Solutions.")
add_writer_note(doc, "INTERNAL LINK: 'grain size analysis' → /en/applications/grain-size-analysis")

add_heading(doc, "H3: Failure Analysis  [~70 words]", level=3)
add_bullet(doc, "Fracture surface examination: image fracture surfaces, fatigue striations, and crack propagation paths using darkfield and MIX observation for enhanced edge contrast.")
add_bullet(doc, "Corrosion and oxidation assessment: document and measure corrosion depth, pitting, and surface degradation.")
add_bullet(doc, "Weld inspection: evaluate weld quality, heat-affected zone microstructure, and weld fusion line integrity.")
add_bullet(doc, "Root cause analysis: the GX53 inverted design accepts large, heavy samples — entire failed components can be placed directly on the stage without sectioning.")

add_heading(doc, "H3: Quality Control & Industrial Inspection  [~70 words]", level=3)
add_bullet(doc, "Coating thickness and uniformity: measure and evaluate electroplated, anodized, and thermal spray coatings using cross-section analysis and PRECiV Count and Measure tools.")
add_bullet(doc, "Surface defect detection: brightfield reveals machining marks, scratches, and surface topography; darkfield highlights pits, burrs, and fine particles that are invisible under direct illumination.")
add_bullet(doc, "Reproducible inspection workflows: coded hardware settings log the observation method, illumination intensity, and magnification — so different operators run the same inspection identically.")
add_bullet(doc, "Standards-compliant reporting: generate reports that reference specific ASTM, ISO, JIS, or DIN criteria directly from PRECiV software.")

add_heading(doc, "H3: Electronics & Semiconductor Inspection  [~70 words]", level=3)
add_bullet(doc, "PCB and solder joint inspection: high-magnification brightfield and darkfield imaging of solder joints, bonding wires, microcracks, and pad defects.")
add_bullet(doc, "Silicon wafer imaging: infrared (IR) objectives on the BX53M allow imaging through silicon — essential for semiconductor defect inspection and process control.")
add_bullet(doc, "ESD-safe design: the BX53M frame is outfitted with electrostatic discharge (ESD) protection to safeguard sensitive electronic components during inspection.")
add_bullet(doc, "Fine particle detection: TruAI deep-learning enables detection and measurement of sub-micron particles on semiconductor surfaces that conventional threshold methods miss.")

add_heading(doc, "H3: Materials R&D & New Material Development  [~70 words]", level=3)
add_bullet(doc, "Alloy and composite development: track microstructural changes resulting from heat treatment, deformation, or processing — and correlate structure with mechanical properties.")
add_bullet(doc, "Ceramics and minerals: polarized light imaging distinguishes mineral phases, reveals stress patterns, and assesses crystal anisotropy in ceramics and geological samples.")
add_bullet(doc, "Additive manufacturing: evaluate layer bonding, porosity, and microstructure in 3D-printed metal and polymer parts.")
add_bullet(doc, "Panoramic imaging: stitch multiple fields of view into a single overview image using PRECiV — essential for carburizing depth and metal flow assessment without a motorized stage.")

# ─── H2 3 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: Contrast Techniques — Seeing More in Every Sample  [~150 words]", level=2)
add_writer_note(doc, "Write a 1–2 sentence intro, then hand off to the table. Ask the web team to render this as a styled comparison table in the CMS. Do not write the table as prose.")
add_body(doc, "Intro direction: Different surface features require different illumination strategies. Evident light microscopes support the full range of contrast techniques — so you choose the method that reveals the structure you need to see.")
doc.add_paragraph()
make_table(doc,
    headers=["Contrast Method", "Best For", "Key Use Cases"],
    rows=[
        ["Brightfield",                 "General surface review",             "Machining marks, scratches, H&E-stained cross-sections, grain boundaries after etching"],
        ["Darkfield",                   "Edge and defect contrast",           "Pits, burrs, fine particles, cracks — features invisible under direct illumination"],
        ["MIX (Brightfield + Darkfield)", "Simultaneous surface + edge detail", "Complex surface textures, polished metal cross-sections, weld inspection"],
        ["Polarized Light",             "Phase and crystal analysis",         "Grain orientation, anisotropic coatings, stress in ceramics, cast iron graphite"],
        ["DIC (Differential Interference Contrast)", "Subtle height differences", "Fine surface texture on polished samples, step heights, coating interfaces"],
        ["Fluorescence",                "Fluorescent features or markers",    "Polymer blends, fluorescent-labeled inclusions, some coating systems"],
        ["Near-Infrared (NIR)",         "Imaging through opaque materials",   "Silicon wafer inspection, semiconductor through-substrate imaging"],
    ],
    col_widths=[5.0, 4.5, 7.0],
)
add_body(doc, "Closing line: All contrast modes are available on a single BX53M instrument frame — switch between methods without changing the sample or objective.")

# ─── H2 4 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: Upright vs. Inverted — Choosing the Right Configuration  [~180 words]", level=2)
add_writer_note(doc, "This is a practical buying-decision section. Write it as a direct, no-fluff comparison. Many readers will be choosing between BX53M and GX53 for the first time. Avoid marketing copy.")
add_body(doc, "Intro line: The right microscope frame depends on your sample type and workflow — here is what to consider.")
doc.add_paragraph()
make_table(doc,
    headers=["",  "BX53M — Upright", "GX53 — Inverted"],
    rows=[
        ["**Illumination**",        "Reflected and transmitted light",         "Reflected light only"],
        ["**Sample orientation**",  "Sample placed face-up on stage",          "Sample placed face-down — no leveling needed"],
        ["**Best for**",            "Standard metallographic mounts, polished cross-sections, geological thin sections", "Large, heavy, or irregularly shaped components; unmounted samples"],
        ["**Sample size limit**",   "Standard microscope slide and mounts",    "Large workpieces; handles thick, bulky components"],
        ["**Contrast methods**",    "Brightfield, darkfield, MIX, polarized, DIC, fluorescence, NIR", "Brightfield, darkfield, MIX, polarized, directional darkfield"],
        ["**ESD protection**",      "Yes — optional ESD frame for electronics","Yes — for electronics and semiconductor work"],
        ["**Motorization**",        "Manual, coded, or fully motorized",       "Manual, coded, or fully motorized"],
        ["**Software**",            "PRECiV™, TruAI, Live AI",                "PRECiV™, TruAI, Live AI"],
    ],
    col_widths=[4.0, 7.0, 5.5],
)
add_writer_note(doc, "After the table, add: 'Not sure which frame is right for your application? Contact an Evident specialist for a configuration recommendation.'")

# ─── H2 5 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: AI-Powered Analysis — TruAI and Live AI  [~150 words]", level=2)
add_writer_note(doc, "Position AI as part of the everyday workflow — not a future add-on. Emphasise that it reduces operator subjectivity, which is a real pain point in QC labs. No coding required is an important differentiator for industrial users.")
add_body(doc, "Cover these points:")
add_bullet(doc, "TruAI is Evident's deep-learning image analysis technology, integrated into PRECiV and Stream Enterprise software. It applies trained neural networks to detect, segment, and measure features that conventional threshold methods miss — including challenging defects, particles, and microstructural features.")
add_bullet(doc, "Instance segmentation: TruAI's instance segmentation merges semantic segmentation and object splitting into a single step, removing manual post-processing from complex measurement workflows.")
add_bullet(doc, "No programming required: TruAI models can be trained on your own image data without coding skills — practical for industrial QC labs where operators are not software engineers.")
add_bullet(doc, "Live AI enables real-time feature detection and image enhancement on the live microscope view — before a single image is captured. This supports fast, AI-assisted go/no-go decisions during inspection.")
add_bullet(doc, "Applications: metallography analysis, grain detection, semiconductor particle inspection, mineralogy, and defect classification in ceramics and coatings.")
add_writer_note(doc, "INTERNAL LINK: 'TruAI' → /en/landing/truai-technology")

# ─── H2 6 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: Standards-Compliant Analysis with PRECiV Material Solutions  [~150 words]", level=2)
add_writer_note(doc, "Lead with the standards — this is what QC engineers and quality managers search for. Name the specific standards. Do not be vague about 'international standards compliance.'")
add_body(doc, "Cover these points:")
add_bullet(doc, "PRECiV™ software integrates directly with Evident light microscopes to deliver a complete imaging, measurement, analysis, and reporting platform — without switching between applications.")
add_bullet(doc, "Material Solutions are optional guided workflow modules within PRECiV that walk operators step-by-step through standardized analyses — reducing subjectivity and enabling reproducible results across operators and sessions.")
add_bullet(doc, "Supported standards include ASTM E112, ISO 643, JIS G 0551, JIS G 0552, DIN 50601, GB/T 6394, GOST 5639 (grain size); and applicable ASTM, ISO, JIS, DIN standards for coatings, inclusions, and weld inspection.")
add_bullet(doc, "Coded hardware integration: the microscope's observation method, illumination intensity, and magnification are stored alongside captured images — so repeat inspections are reproducible.")
add_bullet(doc, "Extended Focus Imaging (EFI): PRECiV stacks images from different focal planes into a single all-in-focus image — useful for rough or textured surfaces that fall outside a single depth of field.")
add_bullet(doc, "Panoramic imaging: generate wide-area overview images by moving the manual stage — no motorized stage required for basic stitching tasks.")
add_writer_note(doc, "INTERNAL LINK: 'PRECiV software' → /en/products/software/preciv. Also link 'Grain Size Solutions' → /en/landing/microscope-solutions-for-materialography/grain-size-solutions")

# ─── H2 7 ────────────────────────────────────────────────────────────────────
add_heading(doc, "H2: How to Choose a Material Science Light Microscope  [~200 words]", level=2)
add_writer_note(doc, "Write this as a self-qualification checklist — plain, direct language. Captures high-intent 'how to choose' and 'which microscope' searches. This section should feel like a conversation with a knowledgeable colleague, not a brochure.")
add_body(doc, "Intro line: The best light microscope for material science depends on your sample type, required contrast methods, and throughput needs. Use these questions to narrow down the right system.")
doc.add_paragraph()
make_table(doc,
    headers=["Question", "What It Means for Your Choice"],
    rows=[
        ["Are your samples large, heavy, or irregular?",
         "If you can't mount or level the sample, the GX53 inverted microscope is the right frame — place the sample face-down with no preparation required."],
        ["Do you need transmitted light (thin sections or geological specimens)?",
         "The BX53M upright supports both reflected and transmitted illumination. The GX53 is reflected light only."],
        ["Do you need infrared imaging through silicon?",
         "The BX53M with NIR objectives supports semiconductor inspection through silicon substrates — the GX53 does not offer this configuration."],
        ["How many contrast methods do you use?",
         "If your workflow demands polarized light, DIC, fluorescence, and darkfield on the same sample, the BX53M modular frame is the most flexible platform."],
        ["Do you need standards-compliant reporting?",
         "PRECiV Material Solutions provide guided, standards-locked workflows for grain size, inclusions, coatings, and more — on any Evident light microscope."],
        ["Is operator-to-operator reproducibility a QC requirement?",
         "Coded hardware integration ensures settings are logged and reproducible — critical in multi-shift and multi-operator QC environments."],
        ["Do you need an all-in-one digital system with no separate camera?",
         "The DSX2000 fully motorized digital microscope integrates imaging, measurement, and analysis in one platform — no separate camera or software setup required."],
    ],
    col_widths=[6.0, 10.5],
)
add_body(doc, "Closing CTA: Need help configuring a system for your application? Contact an Evident specialist — we'll recommend the right microscope, objectives, and software for your workflow.")

add_divider(doc)

# ─── FAQ ─────────────────────────────────────────────────────────────────────
add_heading(doc, "4.  Frequently Asked Questions  [~250 words]", level=1)
add_writer_note(doc, "Format as an accordion FAQ in the CMS. Ask the dev team to add FAQ JSON-LD schema markup for Google rich results eligibility.")

add_faq(doc,
    "What is a material science light microscope?",
    "A material science light microscope — also called a metallurgical or industrial microscope — uses reflected (epi) illumination to examine the surface and near-surface structure of opaque materials such as metals, ceramics, semiconductors, and minerals. Unlike biological microscopes, which transmit light through transparent specimens, material science microscopes reflect light off the sample surface, revealing grain boundaries, surface defects, coating layers, and microstructural features."
)
add_faq(doc,
    "What is the difference between an upright and an inverted metallurgical microscope?",
    "In an upright microscope (BX53M), the objective lens is above the sample — the sample sits face-up on the stage. In an inverted microscope (GX53), the objective lens is below the stage — the sample sits face-down. The inverted design is ideal for large, heavy, or unmounted samples that cannot be leveled or prepared as standard metallographic mounts. Upright microscopes support a broader range of illumination methods, including transmitted light for thin sections.",
    link_text="Compare BX53M upright and GX53 inverted microscopes → /en/material-science-microscopes/metallurgical"
)
add_faq(doc,
    "What contrast methods are available on Evident light microscopes?",
    "Evident's BX53M and GX53 support brightfield, darkfield, MIX (simultaneous brightfield and darkfield), polarized light, differential interference contrast (DIC), fluorescence, and near-infrared (NIR) imaging. The specific combination available depends on the microscope frame and objective configuration. The DSX2000 digital microscope adds shaded relief as an additional visualization mode."
)
add_faq(doc,
    "Can Evident light microscopes perform grain size analysis to ASTM or ISO standards?",
    "Yes. PRECiV Material Solutions include guided grain size workflows compliant with ASTM E112, ISO 643, JIS G 0551, JIS G 0552, DIN 50601, GB/T 6394, GOST 5639, and ASTM E1382 — using both the intercept and planimetric methods. The software automates grain counting, calculates ASTM G numbers and mean intercept lengths, and generates standards-referenced reports.",
    link_text="See Grain Size Solutions → /en/landing/microscope-solutions-for-materialography/grain-size-solutions"
)
add_faq(doc,
    "What is TruAI and how does it help materials analysis?",
    "TruAI is Evident's deep-learning image analysis technology, integrated into PRECiV and Stream Enterprise software. It enables detection, segmentation, and measurement of features that are difficult to resolve with conventional threshold methods — including challenging defects, particles, and microstructural boundaries. TruAI models can be trained on your own image data without any programming skills. Live AI extends this capability to the live microscope view, enabling real-time feature detection during inspection.",
    link_text="Learn more about TruAI → /en/landing/truai-technology"
)

add_divider(doc)

# ─── Internal Links ───────────────────────────────────────────────────────────
add_heading(doc, "5.  Internal Links — Pass to Web Team", level=1)
add_writer_note(doc, "Embed these in the copy using the specified anchor text. Do not use 'click here' or the raw URL as the anchor text.")
make_table(doc,
    headers=["Anchor Text (use exactly this)", "Destination URL"],
    rows=[
        ["material science microscopes",                          "/en/material-science-microscopes"],
        ["metallurgical microscopes",                             "/en/material-science-microscopes/metallurgical"],
        ["BX53M upright metallurgical microscope",                "/en/products/upright/bx53m"],
        ["GX53 inverted metallurgical microscope",                "/en/products/inverted/gx53"],
        ["DSX2000 digital microscope",                            "/en/material-science-microscopes/digital"],
        ["PRECiV software",                                       "/en/products/software/preciv"],
        ["TruAI deep-learning technology",                        "/en/landing/truai-technology"],
        ["grain size analysis",                                   "/en/applications/grain-size-analysis"],
        ["Grain Size Solutions",                                  "/en/landing/microscope-solutions-for-materialography/grain-size-solutions"],
        ["digital microscopes",                                   "/en/material-science-microscopes/digital"],
        ["image analysis software",                               "/en/material-science-microscopes/image-analysis-software"],
    ],
    col_widths=[8.5, 8.0],
)

# ─── Word Count Targets ───────────────────────────────────────────────────────
add_heading(doc, "6.  Word Count Targets by Section", level=1)
make_table(doc,
    headers=["Section", "Target Words"],
    rows=[
        ["Intro / Hero",                                          "100"],
        ["What Is a Material Science Light Microscope?",          "150"],
        ["Applications (all H3s combined)",                       "350"],
        ["Contrast Techniques",                                   "150"],
        ["Upright vs. Inverted",                                  "180"],
        ["AI — TruAI & Live AI",                                  "150"],
        ["Standards / PRECiV Material Solutions",                 "150"],
        ["How to Choose",                                         "200"],
        ["FAQ",                                                   "250"],
        ["**Total**",                                             "**~1,480**"],
    ],
    col_widths=[11.0, 5.5],
)

# ─── Save ─────────────────────────────────────────────────────────────────────
out = "/home/user/content-brief-generator/Evident_Light_Microscopes_Material_Science_Content_Brief.docx"
doc.save(out)
print(f"Saved: {out}")
