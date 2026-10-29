from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pathlib import Path

OUT = Path('/home/user/uploads/MVRLA_AI_v5_6_Complete_Project_Presentation.pptx')

prs = Presentation()
prs.slide_width = Inches(13.333333)
prs.slide_height = Inches(7.5)
prs.core_properties.title = 'MVRLA-AI Engineering Stores — Complete Project Presentation v5.6'
prs.core_properties.subject = 'Inventory search, live valuation, admin photo management, Material Lens, GitHub Pages and PWA deployment'
prs.core_properties.author = 'SOMU SS74'
prs.core_properties.keywords = 'MVRLA-AI, inventory, Material Lens, Apps Script, GitHub Pages, PWA'
prs.core_properties.comments = '© 2026 SOMU SS74 · somu.ss74@gmail.com · All Rights Reserved'
prs.core_properties.language = 'en-US'

# Palette
NAVY = '0A1628'
NAVY2 = '10243D'
NAVY3 = '16334F'
CYAN = '36D2FF'
CYAN2 = 'DFF7FF'
GREEN = '1EA56C'
GREEN2 = 'E4F6EE'
AMBER = 'D88A1D'
AMBER2 = 'FFF3DF'
RED = 'D95555'
RED2 = 'FDECEC'
INK = '102038'
INK2 = '42526A'
MUTED = '6C7890'
LIGHT = 'F5F7FB'
WHITE = 'FFFFFF'
LINE = 'DDE4EE'
PALE = 'EAF0F7'
FONT = 'Aptos'
FONT_DISPLAY = 'Aptos Display'
MONO = 'Aptos Mono'

SW = 13.333333
SH = 7.5


def rgb(h):
    return RGBColor.from_string(h)


def rect(slide, x, y, w, h, fill, line=None, radius=True, transparency=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = rgb(fill)
    if transparency:
        shp.fill.transparency = transparency
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    return shp


def oval(slide, x, y, w, h, fill, line_color=None, line_width=1):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
    if line_color:
        shp.line.color.rgb = rgb(line_color); shp.line.width = Pt(line_width)
    else:
        shp.line.fill.background()
    return shp


def line(slide, x1, y1, x2, y2, color=LINE, width=1.2, dash=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x1), Inches(y1),
                                 Inches(max(0.001, x2-x1)), Inches(max(0.001, y2-y1)))
    # Use a true connector for horizontal/vertical lines.
    slide.shapes._spTree.remove(shp._element)
    ln = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    if dash:
        ln.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    return ln


def text(slide, x, y, w, h, value, size=18, color=INK, bold=False,
         font=FONT, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
         margin=0, italic=False, spacing=1.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = spacing
    r = p.add_run()
    r.text = value
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = rgb(color)
    return box


def rich_text(slide, x, y, w, h, runs, size=18, color=INK,
              align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]; p.alignment = align
    for item in runs:
        # item: text, bold, color, optional size
        r = p.add_run(); r.text = item[0]
        r.font.name = FONT; r.font.size = Pt(item[3] if len(item) > 3 else size)
        r.font.bold = bool(item[1]); r.font.color.rgb = rgb(item[2] if len(item) > 2 else color)
    return box


def bullets(slide, x, y, w, h, items, size=16, color=INK2, bullet_color=None,
            gap=7, leading=1.05, prefix='•'):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for idx, item in enumerate(items):
        if isinstance(item, tuple):
            value, lvl = item
        else:
            value, lvl = item, 0
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = (prefix if lvl == 0 else '–') + ' ' + value
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = leading
        p.space_after = Pt(gap)
        p.font.name = FONT
        p.font.size = Pt(size - (1 if lvl else 0))
        p.font.color.rgb = rgb(color)
        p.margin_left = Inches(0.20 * lvl)
    return box


def pill(slide, x, y, w, label, fill=CYAN2, color=NAVY, size=10.5, line_color=None):
    rect(slide, x, y, w, 0.34, fill, line_color, True)
    text(slide, x+0.06, y+0.01, w-0.12, 0.29, label, size, color, True,
         align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def card(slide, x, y, w, h, title, body=None, accent=CYAN, dark=False,
         title_size=17, body_size=13.5, number=None):
    fill = NAVY2 if dark else WHITE
    border = NAVY3 if dark else LINE
    rect(slide, x, y, w, h, fill, border, True)
    rect(slide, x, y, 0.055, h, accent, None, False)
    compact = h < 1.55
    top = 0.14 if compact else 0.22
    bubble = 0.38 if compact else 0.45
    body_offset = 0.58 if compact else 0.83
    bottom_pad = 0.10 if compact else 0.15
    tx = x + 0.25
    if number is not None:
        rect(slide, tx, y+top, bubble, bubble, accent, None, True)
        text(slide, tx, y+top, bubble, bubble-0.02, str(number), 10.5 if compact else 12,
             NAVY if accent == CYAN else WHITE, True,
             align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        tx += bubble + 0.15
    text(slide, tx, y+top, w-(tx-x)-0.20, 0.38 if compact else 0.46, title, title_size,
         WHITE if dark else INK, True)
    if body:
        body_h = h-body_offset-bottom_pad
        if isinstance(body, list):
            bullets(slide, x+0.25, y+body_offset, w-0.48, body_h, body, body_size,
                    PALE if dark else INK2, gap=4 if compact or body_size <= 13 else 5)
        else:
            text(slide, x+0.25, y+body_offset, w-0.48, body_h, body, body_size,
                 PALE if dark else INK2, spacing=1.02 if compact else 1.06)


def footer(slide, num, dark=False):
    c = '9AA9BC' if dark else MUTED
    line(slide, 0.65, 7.06, 12.68, 7.06, NAVY3 if dark else LINE, 0.8)
    text(slide, 0.66, 7.12, 10.9, 0.20,
         'MVRLA-AI  |  © 2026 SOMU SS74  |  somu.ss74@gmail.com  |  All Rights Reserved',
         8.5, c)
    text(slide, 12.0, 7.10, 0.66, 0.22, f'{num:02d}', 9, c, True, align=PP_ALIGN.RIGHT)


def base_slide(title_value, kicker, num, dark=False, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background.fill; bg.solid(); bg.fore_color.rgb = rgb(NAVY if dark else LIGHT)
    rect(slide, 0.65, 0.28, 0.46, 0.055, CYAN, None, False)
    text(slide, 1.22, 0.18, 10.8, 0.28, kicker.upper(), 9.5,
         CYAN if dark else MUTED, True)
    text(slide, 0.65, 0.58, 12.0, 0.56, title_value, 28,
         WHITE if dark else INK, True, FONT_DISPLAY)
    if subtitle:
        text(slide, 0.67, 1.13, 11.95, 0.34, subtitle, 12.5,
             'AFC0D2' if dark else MUTED)
    footer(slide, num, dark)
    return slide


def notes(slide, value):
    slide.notes_slide.notes_text_frame.text = value


def chevron(slide, x, y, w=0.28, h=0.44, color=CYAN):
    shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = rgb(color); shp.line.fill.background()
    return shp


def label(slide, x, y, w, value, dark=False, size=10.5):
    text(slide, x, y, w, 0.26, value.upper(), size, CYAN if dark else MUTED, True)


def table(slide, x, y, widths, rows, row_h=0.48, header=True, dark=False,
          font_size=12.2, first_bold=True):
    yy = y
    for ri, row in enumerate(rows):
        is_head = header and ri == 0
        fill = NAVY3 if dark and is_head else (NAVY2 if dark else (NAVY2 if is_head else WHITE))
        if not is_head and not dark and ri % 2 == 0:
            fill = 'F0F4F9'
        rect(slide, x, yy, sum(widths), row_h, fill, NAVY3 if dark else LINE, False)
        xx = x
        for ci, value in enumerate(row):
            if ci:
                line(slide, xx, yy, xx, yy+row_h, NAVY3 if dark else LINE, 0.6)
            text(slide, xx+0.12, yy+0.04, widths[ci]-0.24, row_h-0.07, str(value),
                 font_size if not is_head else font_size-0.2,
                 WHITE if is_head or dark else INK,
                 is_head or (first_bold and ci == 0),
                 valign=MSO_ANCHOR.MIDDLE)
            xx += widths[ci]
        yy += row_h
    return yy


# 01 — Cover
s = prs.slides.add_slide(prs.slide_layouts[6])
bg = s.background.fill; bg.solid(); bg.fore_color.rgb = rgb(NAVY)
rect(s, 0, 0, SW, 0.09, CYAN, None, False)
text(s, 0.72, 0.55, 6.8, 0.30, 'COMPLETE PROJECT PRESENTATION', 11, CYAN, True)
text(s, 0.72, 1.18, 7.0, 1.55, 'MVRLA-AI\nEngineering Stores', 38, WHITE, True, FONT_DISPLAY, spacing=0.90)
text(s, 0.75, 3.02, 6.5, 0.85,
     'Inventory search, live valuation, admin-managed photos\nand Material Lens visual matching',
     18, 'C4D3E3', spacing=1.10)
pill(s, 0.75, 4.20, 1.55, 'SECURE ACCESS', NAVY3, CYAN, 9.2)
pill(s, 2.43, 4.20, 1.25, 'LIVE DATA', NAVY3, CYAN, 9.2)
pill(s, 3.81, 4.20, 1.70, 'LENS v5.6', NAVY3, CYAN, 9.2)
pill(s, 5.64, 4.20, 1.25, 'PWA READY', NAVY3, CYAN, 9.2)
text(s, 0.75, 5.30, 6.4, 0.45, 'Implementation • Deployment • Testing • Demonstration', 13, '91A5BB')
text(s, 0.75, 6.27, 6.5, 0.30, '29 October 2026', 11, '91A5BB')
text(s, 0.75, 6.69, 8.0, 0.28, '© 2026 SOMU SS74  |  somu.ss74@gmail.com  |  All Rights Reserved', 10, 'C4D3E3')
# Phone illustration
rect(s, 8.45, 0.62, 3.85, 6.20, NAVY2, '31506B', True)
rect(s, 8.70, 0.94, 3.35, 5.48, '07111F', None, True)
text(s, 9.02, 1.18, 2.7, 0.28, 'MATERIAL LENS', 12, WHITE, True, align=PP_ALIGN.CENTER)
rect(s, 9.06, 1.76, 2.65, 2.34, '0D2138', '25455F', True)
# Viewfinder corners
for x, y, w, h in [(9.28,2.00,0.42,0.035),(9.28,2.00,0.035,0.42),(11.28,2.00,0.42,0.035),(11.67,2.00,0.035,0.42),
                   (9.28,3.66,0.42,0.035),(9.28,3.28,0.035,0.42),(11.28,3.66,0.42,0.035),(11.67,3.28,0.035,0.42)]:
    rect(s, x, y, w, h, CYAN, None, False)
rect(s, 9.42, 2.82, 1.92, 0.035, CYAN, None, False)
text(s, 9.25, 4.33, 2.3, 0.27, 'PHOTO → MATCH → CONFIRM', 9.5, CYAN, True, align=PP_ALIGN.CENTER)
for i, (label_, col) in enumerate([('REFERENCE MATCH 01', GREEN), ('REFERENCE MATCH 02', AMBER), ('REFERENCE MATCH 03', MUTED)]):
    yy = 4.82 + i*0.43
    rect(s, 9.06, yy, 2.65, 0.33, NAVY3, None, True)
    rect(s, 9.18, yy+0.10, 0.12, 0.12, col, None, True)
    text(s, 9.42, yy+0.035, 2.1, 0.24, label_, 8.3, 'D8E3EF', True)
text(s, 11.82, 7.08, 0.50, 0.22, '01', 9, '91A5BB', True, align=PP_ALIGN.RIGHT)
notes(s, 'Open with the business purpose: MVRLA-AI helps engineering stores find parts, see live stock value, preserve admin photos and receive visual part suggestions. Explain that this presentation covers the complete implementation and the recommended GitHub Pages plus PWA deployment path. Do not reveal real credentials, admin codes or OTP values during the presentation.')

# 02 — Overview
s = base_slide('Project at a Glance', '01 · Executive overview', 2,
               subtitle='A single web application connects inventory data, user access, photo records and assisted identification.')
card(s, 0.65, 1.65, 2.90, 2.12, 'Find parts faster', 'Search by part number, item description or machine. Open a complete detail drawer for the selected material.', CYAN)
card(s, 3.70, 1.65, 2.90, 2.12, 'Keep values live', 'Read stock and unit price from Google Sheets. Calculate each line value and the total inventory valuation.', GREEN)
card(s, 6.75, 1.65, 2.90, 2.12, 'Preserve photos', 'Store admin images in Drive and maintain a separate Photos sheet keyed by the exact part number.', AMBER)
card(s, 9.80, 1.65, 2.88, 2.12, 'Identify by photo', 'Show the user photo beside each admin reference, with five ranked candidates and a confirmation step.', CYAN)
label(s, 0.67, 4.08, 3.0, 'Key operating principles')
principles = [
    ('2', 'Roles', 'User and administrator'),
    ('5', 'Suggestions', 'Top matches per scan'),
    ('5', 'Visual signals', 'Five weighted features'),
    ('0', 'Auto-deletes', 'Photos require admin removal')
]
for i, (n, t, b) in enumerate(principles):
    x = 0.65 + i*3.05
    rect(s, x, 4.47, 2.90, 1.66, WHITE, LINE, True)
    text(s, x+0.20, 4.65, 0.70, 0.60, n, 28, CYAN if i != 3 else GREEN, True, FONT_DISPLAY)
    text(s, x+0.94, 4.69, 1.72, 0.30, t, 15, INK, True)
    text(s, x+0.20, 5.44, 2.48, 0.42, b, 12.5, INK2)
notes(s, 'Use this slide to summarize the four main capabilities. The numbers are design choices, not measured accuracy claims: two roles, five suggestions, five visual signals and no automatic photo deletion. Emphasize that Material Lens assists the user; it does not make the final identification decision.')

# 03 — Roles and boundaries
s = base_slide('Users, Administrator and Project Boundaries', '02 · Roles and rules', 3)
card(s, 0.65, 1.50, 5.88, 2.60, 'Standard user', [
    'Sign in, reset a password with email OTP and search inventory.',
    'View prices, stock, valuation, photos and full part details.',
    'Capture or choose a photo for Material Lens.',
    'Review suggestions and confirm the correct part.'
], CYAN, title_size=19, body_size=14)
card(s, 6.78, 1.50, 5.90, 2.60, 'Administrator', [
    'Upload or delete material reference photos.',
    'Maintain the photo location during image upload.',
    'Review users and block, unblock or delete accounts.',
    'Manage price updates through the protected backend action.'
], GREEN, title_size=19, body_size=14)
label(s, 0.67, 4.28, 4.0, 'Non-negotiable project rules')
rect(s, 0.65, 4.65, 12.03, 1.73, NAVY2, None, True)
bullets(s, 0.95, 4.91, 11.45, 1.28, [
    'Photo upload and deletion are admin-only; photos must never be removed automatically.',
    'The Photos sheet uses the exact trimmed Part_Number as its persistence key.',
    'Lens compares only against admin reference photos; the user must confirm the final selection.',
    'Existing materials, billing and valuation behavior must remain unchanged unless explicitly requested.'
], 13.5, PALE, gap=5)
notes(s, 'Clarify the separation of duties. A normal user can consume photos for identification but cannot create or delete reference photos. The administrator controls the reference library and user administration. The persistence and confirmation rules are intentional requirements, not optional enhancements.')

# 04 — Architecture
s = base_slide('System Architecture', '03 · End-to-end design', 4, dark=True,
               subtitle='The frontend is static; Google Apps Script orchestrates sheets, Drive images and email OTP.')
label(s, 0.67, 1.63, 2.0, 'Client', True)
card(s, 0.65, 2.00, 2.45, 1.42, 'Chrome / PWA', 'Recommended full browser path', CYAN, True, 14, 12.2)
card(s, 0.65, 3.65, 2.45, 1.42, 'Kodular WebView', 'Existing wrapper; camera limited', AMBER, True, 13.5, 12.2)
chevron(s, 3.30, 3.10, 0.36, 0.60)
label(s, 3.92, 1.63, 2.0, 'Frontend', True)
card(s, 3.86, 2.00, 2.42, 3.07, 'index.html', [
    'HTML and CSS interface',
    'Authentication and search',
    'v5.6 multi-feature engine',
    'Local descriptor cache',
    'PWA integration'
], CYAN, True, 18, 12.5)
chevron(s, 6.50, 3.10, 0.36, 0.60)
label(s, 7.03, 1.63, 2.2, 'Backend', True)
card(s, 6.98, 2.00, 2.42, 3.07, 'Apps Script', [
    'GET and POST router',
    'Business validation',
    'Photo persistence',
    'Image proxy: getimg',
    'OTP email service'
], GREEN, True, 18, 12.5)
chevron(s, 9.62, 3.10, 0.36, 0.60)
label(s, 10.13, 1.63, 2.2, 'Google services', True)
card(s, 10.08, 1.98, 2.60, 1.05, 'Google Sheets', 'Materials, Photos, Users, OTP and price history', CYAN, True, 15, 11.5)
card(s, 10.08, 3.16, 2.60, 0.90, 'Google Drive', 'Admin reference images', AMBER, True, 15, 11.5)
card(s, 10.08, 4.19, 2.60, 0.90, 'MailApp', 'Signup and reset OTP', GREEN, True, 15, 11.5)
rect(s, 0.65, 5.57, 12.03, 0.72, NAVY3, None, True)
text(s, 0.92, 5.76, 11.50, 0.33,
     'Lens image path: Drive file ID → getimg → base64 data URL → canvas-safe descriptor processing in the browser.',
     13.2, PALE, align=PP_ALIGN.CENTER)
notes(s, 'Trace one request from the client to the backend and Google services. GitHub Pages hosts only the static frontend; it does not run the Apps Script file. The getimg action is important because direct Drive thumbnails are not reliable for canvas processing. Lens descriptors are calculated on the device after the reference image is returned as a data URL.')

# 05 — Package and data
s = base_slide('Project Package and Data Model', '04 · Files and storage', 5)
label(s, 0.67, 1.49, 3.2, 'Deployable project files')
rect(s, 0.65, 1.85, 4.15, 4.55, NAVY2, None, True)
files = [
    ('index.html', 'Complete frontend and Lens v5.6'),
    ('MVRLA_Backend_LivePrice.gs', 'Apps Script backend source'),
    ('manifest.webmanifest', 'PWA name, icons and display'),
    ('sw.js', 'App-shell cache; live API excluded'),
    ('icon-192.png', 'Small application icon'),
    ('icon-512.png', 'Large application icon')
]
for i, (f, desc) in enumerate(files):
    yy = 2.08 + i*0.67
    text(s, 0.92, yy, 3.60, 0.25, f, 12.2, CYAN, True, MONO)
    text(s, 0.92, yy+0.27, 3.55, 0.24, desc, 10.8, 'B5C5D6')
label(s, 5.12, 1.49, 3.2, 'Google Sheet tabs')
rows = [
    ['Sheet', 'Primary purpose'],
    ['Materials', 'Part details, machine, location, stock, price and image reference'],
    ['Photos', 'Persistent image URL, location and update date by part number'],
    ['Users', 'Account identity, salted hash fields, role and status'],
    ['OTP_Log', 'OTP purpose, expiry and verification state'],
    ['Price_History', 'Old/new price, editor, date and resulting stock value']
]
table(s, 5.10, 1.85, [1.62, 5.93], rows, 0.65, True, False, 12.0)
rect(s, 5.10, 5.94, 7.58, 0.46, CYAN2, None, True)
text(s, 5.28, 6.05, 7.20, 0.25,
     'GitHub hosts the frontend; the deployed Apps Script endpoint serves live data.',
     11.5, NAVY, True, align=PP_ALIGN.CENTER)
notes(s, 'Explain which files must be uploaded to the static host and which file belongs in Apps Script. The Photos sheet is separate from Materials so a material row can be edited or recreated without intentionally deleting the photo record. The backend source can be kept with the project for backup, but GitHub Pages does not execute it.')

# 06 — Auth
s = base_slide('Authentication and Account Management', '05 · Secure access flow', 6)
label(s, 0.67, 1.48, 3.0, 'Signup and login')
flow = [('1', 'Enter account\ndetails'), ('2', 'Send email\nOTP'), ('3', 'Verify and\nregister'), ('4', 'Login and\nopen app')]
for i, (n, t) in enumerate(flow):
    x = 0.65 + i*3.06
    rect(s, x, 1.89, 2.62, 1.15, WHITE, LINE, True)
    rect(s, x+0.18, 2.12, 0.46, 0.46, CYAN if i < 3 else GREEN, None, True)
    text(s, x+0.18, 2.12, 0.46, 0.43, n, 12, NAVY if i < 3 else WHITE, True,
         align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    text(s, x+0.80, 2.00, 1.58, 0.78, t, 14.5, INK, True, valign=MSO_ANCHOR.MIDDLE)
    if i < 3: chevron(s, x+2.72, 2.25, 0.23, 0.35, MUTED)
card(s, 0.65, 3.43, 5.88, 2.63, 'OTP and password behavior', [
    'Six-digit OTP sent with MailApp; no paid SMS gateway.',
    'OTP expires after five minutes and has a verification state.',
    'Password storage uses salt and password-hash fields.',
    'Reset password uses the same email OTP pattern.',
    'Username is remembered for faster repeat login.'
], CYAN, body_size=13.5)
card(s, 6.78, 3.43, 5.90, 2.63, 'Administrator controls', [
    'List registered users and review account status.',
    'Block or unblock an account when required.',
    'Delete an account through the protected action.',
    'Keep the admin credential out of presentations and public source.',
    'Run authorizeEmail_() once before testing OTP delivery.'
], GREEN, body_size=13.5)
notes(s, 'Demonstrate signup or login with a test account, never a real user account. The OTP is valid for five minutes. The MailApp scope must be authorized from the Apps Script editor by running authorizeEmail_ once. For production, review the current password policy and never expose the admin credential in slides or public frontend source.')

# 07 — Inventory
s = base_slide('Inventory Search and Live Valuation', '06 · Core inventory workflow', 7)
label(s, 0.67, 1.48, 3.0, 'User workflow')
steps = ['Sync material data', 'Search part, name or machine', 'Review result cards', 'Open full details']
for i, t in enumerate(steps):
    y = 1.91 + i*0.87
    rect(s, 0.65, y, 5.48, 0.67, WHITE, LINE, True)
    rect(s, 0.84, y+0.13, 0.40, 0.40, CYAN, None, True)
    text(s, 0.84, y+0.13, 0.40, 0.37, str(i+1), 11, NAVY, True,
         align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    text(s, 1.46, y+0.15, 4.36, 0.36, t, 15, INK, True, valign=MSO_ANCHOR.MIDDLE)
label(s, 6.55, 1.48, 3.0, 'Valuation logic')
rect(s, 6.52, 1.91, 6.16, 2.20, NAVY2, None, True)
text(s, 6.90, 2.24, 5.40, 0.35, 'LINE VALUE', 11, CYAN, True, align=PP_ALIGN.CENTER)
text(s, 6.90, 2.70, 5.40, 0.46, 'Stock quantity × Unit price', 23, WHITE, True, FONT_DISPLAY, align=PP_ALIGN.CENTER)
line(s, 7.42, 3.39, 11.78, 3.39, NAVY3, 1.0)
text(s, 6.90, 3.57, 5.40, 0.30, 'TOTAL = sum of all material line values', 13, 'C4D3E3', True, align=PP_ALIGN.CENTER)
card(s, 6.52, 4.38, 6.16, 2.00, 'Refresh and offline behavior', [
    'Automatic data synchronization runs every five minutes.',
    'The PWA caches the app shell for faster repeat opening.',
    'Apps Script calls are not cached; live data requires network access.'
], GREEN, body_size=12.8)
notes(s, 'Show a part-number search and then a description or machine search. Open the detail drawer to show all available fields. Explain the valuation formula with one example from the test data, but do not claim that the values are current unless the sheet has just been verified. Offline mode can display the cached shell; it cannot provide live stock or new reference images without connectivity.')

# 08 — Photo persistence
s = base_slide('Admin Photo Workflow and Persistence', '07 · Reference image management', 8)
label(s, 0.67, 1.48, 4.0, 'Upload sequence')
seq = [('01', 'Find part'), ('02', 'Choose image'), ('03', 'Set location'), ('04', 'Save'), ('05', 'Users see photo')]
for i, (n, t) in enumerate(seq):
    x = 0.65 + i*2.46
    rect(s, x, 1.87, 2.12, 1.02, WHITE, LINE, True)
    text(s, x+0.16, 2.02, 0.42, 0.26, n, 10, CYAN, True, MONO)
    text(s, x+0.16, 2.36, 1.78, 0.30, t, 13.5, INK, True)
    if i < 4: chevron(s, x+2.17, 2.20, 0.20, 0.33, MUTED)
card(s, 0.65, 3.30, 5.88, 2.90, 'Storage path', [
    'The original image is stored in the MVRLA Drive image folder.',
    'The Photos sheet stores Part_Number, Image_URL, Location and Updated_Date.',
    'Matching material rows and the persistent Photos record are updated.',
    'Only an authorized administrator can upload or delete the reference image.'
], CYAN, body_size=12.7)
card(s, 6.78, 3.30, 5.90, 2.90, 'Why the photo survives row changes', [
    'The exact trimmed part number is the persistence key.',
    'Materials image is used first; Photos provides the fallback image.',
    'A blank material location can fall back to the saved Photos location.',
    'Deleting and re-adding a material row does not require deleting its photo.',
    'Admin deletion is the only intentional photo-removal path.'
], GREEN, body_size=12.7)
notes(s, 'This slide addresses the earlier photo-loss problem. The important design decision is the separate Photos sheet keyed by the exact part number. Demonstrate with a test part: upload an image and location, then explain how the fallback restores them if a Materials row is recreated. Do not test deletion on a production photo.')

# 09 — Lens comparison design and journey
s = base_slide('Material Lens v5.6: Compare Before Confirmation', '08 · Side-by-side comparison design', 9, dark=True,
               subtitle='Every candidate shows the user photo, the admin reference, the overall score and a deliberate confirmation action.')
# Comparison mockup
rect(s, 0.65, 1.55, 7.60, 4.92, NAVY2, NAVY3, True)
text(s, 0.95, 1.76, 3.70, 0.26, 'CANDIDATE #1 OF 5', 10, CYAN, True, MONO)
text(s, 6.47, 1.69, 1.25, 0.48, '82%', 29, GREEN, True, FONT_DISPLAY, align=PP_ALIGN.RIGHT)
text(s, 6.17, 2.13, 1.55, 0.18, 'ILLUSTRATIVE SCORE', 7.2, '91A5BB', True, align=PP_ALIGN.RIGHT)
text(s, 0.95, 2.13, 3.15, 0.20, 'YOUR PHOTO', 8.5, 'AFC0D2', True, MONO, align=PP_ALIGN.CENTER)
text(s, 4.75, 2.13, 3.15, 0.20, 'ADMIN REFERENCE', 8.5, 'AFC0D2', True, MONO, align=PP_ALIGN.CENTER)
rect(s, 0.95, 2.39, 3.15, 1.54, '07111F', '29445E', True)
rect(s, 4.75, 2.39, 3.15, 1.54, '07111F', '29445E', True)
# Same bearing with different scale and illumination
oval(s, 2.00, 2.60, 1.05, 1.05, '9AA8B5', 'D4DEE7', 1.2)
oval(s, 2.30, 2.90, 0.45, 0.45, '07111F', '5D7185', 1.0)
oval(s, 5.85, 2.66, 0.95, 0.95, '718391', 'B9C6D0', 1.2)
oval(s, 6.125, 2.935, 0.40, 0.40, '07111F', '4B6173', 1.0)
oval(s, 4.26, 2.96, 0.38, 0.38, CYAN, None)
text(s, 4.26, 2.99, 0.38, 0.28, 'VS', 9, NAVY, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
text(s, 0.95, 4.13, 5.55, 0.29, 'Cam Follower Bearing KR22', 16, WHITE, True)
text(s, 0.95, 4.47, 5.70, 0.22, '262-030-00010  ·  Bay 1  ·  BM Stacking', 9.5, 'AFC0D2', False, MONO)
for j, (lab, val) in enumerate([('Luma structure',84),('Edge geometry',79),('Perceptual pattern',76)]):
    yy = 4.84 + j*0.29
    text(s, 0.95, yy, 1.55, 0.19, lab, 8.7, 'C4D3E3')
    rect(s, 2.60, yy+0.045, 3.90, 0.075, NAVY3, None, True)
    rect(s, 2.60, yy+0.045, 3.90*val/100, 0.075, CYAN if j != 1 else GREEN, None, True)
    text(s, 6.66, yy-0.01, 0.55, 0.20, str(val)+'%', 8.7, WHITE, True, MONO, align=PP_ALIGN.RIGHT)
rect(s, 0.95, 5.80, 6.95, 0.43, CYAN, None, True)
text(s, 0.95, 5.87, 6.95, 0.27, 'COMPARE  →  CONFIRM & OPEN PART', 11, NAVY, True, align=PP_ALIGN.CENTER)
# User flow
rect(s, 8.50, 1.55, 4.18, 4.92, NAVY2, NAVY3, True)
text(s, 8.78, 1.78, 3.55, 0.28, 'USER FLOW', 11, CYAN, True, MONO)
flow56 = [
    ('1','Open Lens and capture or select a photo.'),
    ('2','Optionally enter a part-start filter.'),
    ('3','Analyze the component with rotation variants.'),
    ('4','Review five side-by-side candidates.'),
    ('5','Open the full similarity breakdown.'),
    ('6','Confirm the correct part to open details.')
]
for i,(n,t) in enumerate(flow56):
    yy = 2.23 + i*0.65
    oval(s, 8.78, yy, 0.33, 0.33, CYAN if i < 3 else GREEN, None)
    text(s, 8.78, yy+0.01, 0.33, 0.27, n, 9, NAVY if i < 3 else WHITE, True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    text(s, 9.28, yy-0.01, 3.05, 0.47, t, 11.2, 'D5E1EC')
notes(s, 'This is the new v5.6 interaction design. The user no longer taps a text-only result and immediately opens a part. Each candidate presents the user photo and admin reference together. The Compare and Confirm screen adds the overall score, five feature bars and the correct material index before the existing detail drawer opens. The score and bearing graphics on this slide are illustrative, not measured production results.')

# 10 — Matching engine v5.6
s = base_slide('Visual Matching Engine — Version 5.6', '09 · Multi-feature comparison', 10,
               subtitle='Foreground-focused normalization and five signals replace the earlier brightness-heavy comparison.')
features56 = [
    ('30% · Luma', '32 × 32 normalized crop with zero-mean, unit-norm correlation.', CYAN),
    ('35% · HOG edges', 'Gradient orientation cells with block normalization for shape and edges.', GREEN),
    ('20% · pHash', 'A 63-bit DCT perceptual hash for robust visual-pattern comparison.', CYAN),
    ('10% · Colour', 'A 64-bin RGB histogram compares colour distribution overlap.', AMBER),
    ('5% · Aspect', 'Rotation-invariant crop aspect ratio adds a small geometric check.', GREEN)
]
for i,(t,b,col) in enumerate(features56):
    card(s, 0.65+i*2.44, 1.70, 2.27, 2.70, t, b, col, False, 14.5, 11.4)
rect(s, 0.65, 4.69, 12.03, 0.66, NAVY2, None, True)
text(s, 0.85, 4.86, 11.63, 0.30,
     'Score = 0.30 luma + 0.35 HOG + 0.20 pHash + 0.10 colour + 0.05 aspect',
     14.2, WHITE, True, MONO, align=PP_ALIGN.CENTER)
label(s, 0.67, 5.59, 2.0, 'Result badges')
pill(s, 0.65, 5.94, 1.63, 'GREEN  ≥ 75', GREEN2, GREEN, 11)
pill(s, 2.48, 5.94, 1.83, 'AMBER  50–74', AMBER2, AMBER, 11)
pill(s, 4.51, 5.94, 1.63, 'GREY  < 50', PALE, INK2, 11)
text(s, 6.58, 5.70, 6.10, 0.65,
     'Query variants at 0°, 90°, 180° and 270° improve orientation tolerance. Displayed scores are capped at 99% and remain similarity scores, not probabilities.',
     11.8, INK2, True, valign=MSO_ANCHOR.MIDDLE)
notes(s, 'Version 5.6 first estimates the foreground from border contrast and a central connected component, then normalizes a 32 by 32 crop. It compares normalized luma, HOG edge geometry, a DCT perceptual hash, colour overlap and aspect ratio. Four query orientations are evaluated with a small penalty for rotated alternatives. Reference descriptors and their thumbnails use the new v56 cache prefix, so stale v2 descriptors are not reused.')

# 11 — Performance
s = base_slide('Performance and Caching Design', '10 · Reducing scan waiting time', 11)
label(s, 0.67, 1.48, 4.0, 'Background descriptor warm-up')
warm = [('Sync', 'Material data arrives'), ('Wait', 'Six-second delay'), ('Prepare', 'Up to 80 references'), ('Pause', '450 ms between items'), ('Reuse', 'Descriptor stored locally')]
for i, (t, b) in enumerate(warm):
    x = 0.65 + i*2.46
    rect(s, x, 1.89, 2.12, 1.36, WHITE, LINE, True)
    text(s, x+0.16, 2.08, 1.80, 0.30, t, 15, INK, True)
    text(s, x+0.16, 2.52, 1.78, 0.48, b, 11.8, INK2)
    if i < 4: chevron(s, x+2.17, 2.35, 0.20, 0.33, MUTED)
card(s, 0.65, 3.68, 5.88, 2.45, 'Device-side optimizations', [
    'Video capture is limited to 720 px; gallery photos to 900 px wide.',
    'The v56 cache stores descriptors plus 144 px reference thumbnails.',
    'Three references load in parallel; warm-up stops when the app is hidden.'
], CYAN, body_size=13.0)
card(s, 6.78, 3.68, 5.90, 2.45, 'Data freshness safeguards', [
    'The service worker caches the app shell, not Apps Script responses.',
    'Stock, prices and reference data are requested live when online.',
    'Four query orientations are compared locally; live data still bypasses the shell cache.'
], GREEN, body_size=13.0)
notes(s, 'Describe the warm-up as a best-effort background optimization, not a guarantee. Version 5.6 caches both the descriptor and a small reference thumbnail, loads up to three missing references in parallel and evaluates four query orientations locally. The v5.6 service worker removes older MVRLA shell caches and uses network-first navigation so an online user receives the updated index.html.')

# 12 — Accuracy
s = base_slide('Reading v5.6 Scores Correctly', '11 · Interpretation and improvement', 12)
rect(s, 0.65, 1.51, 5.88, 2.06, AMBER2, None, True)
label(s, 0.92, 1.76, 2.5, 'Illustrative example')
rich_text(s, 0.92, 2.19, 5.30, 0.48, [
    ('Candidate A: ', True, INK, 14), ('82', True, GREEN, 20), ('   |   Candidate B: ', True, INK, 14), ('79', True, AMBER, 20)
])
text(s, 0.92, 2.87, 5.26, 0.42,
     'Close scores still require side-by-side visual and part-number confirmation.', 12.8, INK2)
card(s, 6.78, 1.51, 5.90, 2.06, 'What v5.6 adds', 'Foreground normalization, five signals and rotation variants improve ranking. A single-reference system still cannot guarantee the correct component.', CYAN, title_size=17, body_size=13.2)
card(s, 0.65, 3.91, 5.88, 2.35, 'Improve results with the current system', [
    'Center the complete part with sharp focus, good light and a plain background.',
    'Use a known part-start prefix; retake from another angle when needed.',
    'Confirm the reference image and part number before opening details.'
], GREEN, body_size=12.3)
card(s, 6.78, 3.91, 5.90, 2.35, 'Future accuracy upgrades', [
    'Store three to five reference views per part; this requires a data-model upgrade.',
    'Evaluate a labeled vision model or paid AI service after cost and privacy review.',
    'Keep human confirmation even with a more advanced model.'
], CYAN, body_size=12.3)
notes(s, 'Be direct and honest: version 5.6 improves the comparison design and ranking features, but no responsible presentation should promise 100 percent identification from one reference photo. Angle, scale, lighting, dirt, background and similar geometry still matter. When the first two scores are close, the interface explicitly tells the user to compare carefully rather than auto-selecting the top row.')

# 13 — Camera issue
s = base_slide('Camera Compatibility: Root Cause and Decision', '12 · Browser versus WebView', 13, dark=True)
rows = [
    ['Capability', 'Chrome / installed PWA', 'Kodular built-in WebView'],
    ['Secure camera preview', 'Supported with HTTPS and permission', 'Observed blocked or unavailable'],
    ['Camera file capture', 'Uses the Android camera flow', 'May open gallery or not respond'],
    ['Gallery image return', 'Standard browser file input', 'Can be unreliable in the wrapper'],
    ['Recommended role', 'Primary deployment path', 'Not recommended for Lens camera']
]
table(s, 0.65, 1.61, [2.60, 4.55, 4.88], rows, 0.63, True, True, 12.0)
card(s, 0.65, 5.16, 5.88, 1.10, 'Why a Camera component alone is insufficient', 'It can declare or request permission, but it does not implement the WebView file-chooser and image-return bridge.', AMBER, True, 14.2, 11.5)
card(s, 6.78, 5.16, 5.90, 1.10, 'Recommended decision', 'Use GitHub Pages with Chrome/PWA. Treat a native Kodular Lens as a separate project.', CYAN, True, 14.2, 11.5)
notes(s, 'Separate the hosting issue from the wrapper issue. GitHub Pages supplies HTTPS, but the camera behavior still depends on the browser engine. If the same URL is placed inside the limited WebView, the limitation remains. The reliable presentation path is Chrome or the installed PWA. Do not spend demonstration time trying unsupported WebView workarounds.')

# 14 — GitHub Pages
s = base_slide('GitHub Pages Deployment — Step by Step', '13 · Static frontend hosting', 14)
steps = [
    ('1', 'Create a repository', 'Use private source hosting where the plan and company policy allow.'),
    ('2', 'Upload frontend files', 'Place index.html, manifest, service worker and both icons in the repository root.'),
    ('3', 'Enable Pages', 'Settings → Pages → Deploy from a branch → main → / (root) → Save.'),
    ('4', 'Open the HTTPS URL', 'Wait for the deployment and open the published Pages address in Chrome.'),
    ('5', 'Verify the backend', 'Keep the deployed Apps Script /exec URL in APP_SCRIPT_URL and test live data.'),
    ('6', 'Test the complete flow', 'Check login, inventory, photo records, Lens camera, gallery and match confirmation.')
]
for i, (n, t, b) in enumerate(steps):
    row = i//2; col = i%2
    x = 0.65 + col*6.13; y = 1.49 + row*1.54
    card(s, x, y, 5.90, 1.29, t, b, CYAN if i < 3 else GREEN, False, 15.5, 11.5, n)
rect(s, 0.65, 6.18, 12.03, 0.30, AMBER2, None, True)
text(s, 0.80, 6.20, 11.73, 0.24,
     'Uploading files alone does not publish the app. Public repositories expose source code; never include secrets in frontend code.',
     10.8, AMBER, True, align=PP_ALIGN.CENTER)
notes(s, 'Show the repository file list and the Pages settings if a live deployment demonstration is required. Use a test repository rather than exposing proprietary source. The Pages URL will use the GitHub username and repository name. The backend remains in Apps Script and must already be deployed as a working web application.')

# 15 — PWA
s = base_slide('Install as a PWA and Enable the Camera', '14 · Recommended user setup', 15)
label(s, 0.67, 1.48, 4.0, 'Installation sequence')
install = [('01', 'Open the Pages URL\nin Android Chrome'), ('02', 'Chrome menu →\nAdd to Home screen'), ('03', 'Install and open\nthe MVRLA-AI icon'), ('04', 'Lens → Camera →\nAllow permission')]
for i, (n, t) in enumerate(install):
    x = 0.65 + i*3.06
    rect(s, x, 1.89, 2.72, 1.35, WHITE, LINE, True)
    text(s, x+0.18, 2.06, 0.42, 0.24, n, 10, CYAN, True, MONO)
    text(s, x+0.18, 2.43, 2.34, 0.60, t, 14.2, INK, True)
    if i < 3: chevron(s, x+2.77, 2.37, 0.20, 0.33, MUTED)
card(s, 0.65, 3.67, 5.88, 2.45, 'If the camera remains blocked', [
    'Confirm that the page uses HTTPS, not a local or insecure address.',
    'Check Android Chrome permission and the site camera permission.',
    'Close and reopen the installed app after changing permissions.'
], AMBER, body_size=13.0)
card(s, 6.78, 3.67, 5.90, 2.45, 'What the PWA provides', [
    'Home-screen icon, portrait orientation and full-screen display.',
    'A cached shell plus standard Chrome camera and gallery behavior.',
    'Live inventory and new reference photos still require internet access.'
], GREEN, body_size=13.0)
notes(s, 'Install the PWA before the presentation and grant camera permission in advance. Keep Chrome updated. Demonstrate both camera and gallery selection. If a company device manages browser permissions, coordinate with IT rather than attempting to bypass device policy.')

# 16 — Backend deployment
s = base_slide('Apps Script Backend Deployment', '15 · Publishing live services', 16)
steps = [
    ('1', 'Open the correct project', 'Use the Apps Script project connected to the intended Google Sheet.'),
    ('2', 'Save the final backend', 'Paste MVRLA_Backend_LivePrice.gs and resolve any save conflict first.'),
    ('3', 'Run setupSheet() once', 'Create or verify Materials, Price_History, Users, OTP_Log and Photos.'),
    ('4', 'Run authorizeEmail_() once', 'Grant the MailApp scope and confirm the test email.'),
    ('5', 'Deploy a New version', 'Deploy → Manage deployments → Edit → New version → Deploy.'),
    ('6', 'Configure and test', 'Use the /exec URL in the frontend; test ping, getAll and getimg.')
]
for i, (n, t, b) in enumerate(steps):
    row = i//2; col = i%2
    x = 0.65 + col*6.13; y = 1.49 + row*1.42
    card(s, x, y, 5.90, 1.18, t, b, CYAN if i < 4 else GREEN, False, 15, 11.2, n)
rect(s, 0.65, 5.92, 12.03, 0.53, RED2, None, True)
text(s, 0.87, 6.04, 11.59, 0.29,
     'Saving code does not update the live web app. “Unknown POST action: sendotp” commonly indicates an older deployed version.',
     11.8, RED, True, align=PP_ALIGN.CENTER)
notes(s, 'Use a test deployment when possible. The web app must execute with an account authorized to access the Sheet, Drive folder and MailApp. Choose access settings according to company policy while ensuring the frontend can call the endpoint. Every backend change needs a new deployment version; otherwise users may continue reaching older code.')

# 17 — Tests
s = base_slide('End-to-End Acceptance Tests', '16 · Verification checklist', 17)
rows = [
    ['Test area', 'Action', 'Expected result'],
    ['Signup and reset', 'Request OTP; verify valid and expired cases', 'Registration/reset works; expiry is enforced'],
    ['Login and status', 'Test valid, invalid and blocked accounts', 'Only eligible accounts enter the app'],
    ['Inventory', 'Search and compare values with the sheet', 'Details and calculated valuation agree'],
    ['Admin image', 'Upload image and location as authorized admin', 'Drive, Photos and material view are updated'],
    ['Persistence', 'Recreate a test Materials row', 'Photo and location fall back from Photos'],
    ['Lens input', 'Test camera, gallery, cancel and invalid file', 'Preview appears; failures are handled clearly'],
    ['Lens results', 'Run synthetic tests; compare five candidates', 'Correct test ranking, both photos and breakdown appear'],
    ['PWA', 'Install; test permission and offline behavior', 'Icon works; live-data limits are understood']
]
table(s, 0.65, 1.48, [2.08, 4.46, 5.49], rows, 0.55, True, False, 11.6)
rect(s, 0.65, 6.53, 12.03, 0.25, GREEN2, None, True)
text(s, 0.80, 6.53, 11.73, 0.22, 'Use test accounts and test parts. Record the device, browser, date and result for each acceptance run.', 10.3, GREEN, True, align=PP_ALIGN.CENTER)
notes(s, 'A complete demonstration should be repeatable. Prepare a checklist and test the exact device that will be used for the presentation. Include negative tests such as an expired OTP, an invalid file and a blocked account. Avoid destructive tests on production photos or real user accounts.')

# 18 — Troubleshooting
s = base_slide('Troubleshooting Guide', '17 · Common symptoms and fixes', 18)
issues = [
    ('Camera does not open', 'Use Chrome/PWA over HTTPS. Check Chrome and site camera permissions; do not rely on the limited WebView.'),
    ('OTP action is unknown', 'Deploy the backend as a New version and confirm that the frontend uses the current /exec URL.'),
    ('OTP email does not arrive', 'Run authorizeEmail_ once, verify the email address, spam folder and MailApp quota.'),
    ('Lens returns no matches', 'Confirm admin photos exist, Drive IDs are valid, getimg is deployed and the device is online.'),
    ('A photo seems lost', 'Check the exact Part_Number and Photos record. Verify the Materials and location fallback data.'),
    ('An update is not visible', 'Confirm the correct Pages branch/root, publish again, then reload and verify service-worker updates.')
]
for i, (t, b) in enumerate(issues):
    row = i//2; col = i%2
    x = 0.65 + col*6.13; y = 1.48 + row*1.65
    card(s, x, y, 5.90, 1.40, t, b, AMBER if i in (0,2,4) else CYAN, False, 15.2, 11.6, i+1)
notes(s, 'Start troubleshooting with the simplest distinction: frontend hosting, backend deployment, device permission or source data. Check the browser console only on an authorized test device and avoid sharing logs that contain account or OTP information. The camera issue is expected in the limited WebView and should not be treated as a Sheets or backend failure.')

# 19 — Security
s = base_slide('Security and Corporate Readiness', '18 · Production recommendations', 19,
               subtitle='The current build is a working prototype. Apply organizational controls before a public or production release.')
card(s, 0.65, 1.66, 5.88, 4.35, 'Application and source controls', [
    'Use an approved private repository or private hosting path where possible.',
    'Move administrator secrets out of client-side JavaScript and enforce them only on the server.',
    'Strengthen the current four-character minimum password policy for production.',
    'Use HTTPS everywhere and review every third-party extension before inclusion.',
    'Do not place real admin codes, passwords, OTP values or user data in slides.'
], CYAN, title_size=18, body_size=13.5)
card(s, 6.78, 1.66, 5.90, 4.35, 'Data and operational controls', [
    'Restrict Google Sheet and Drive access to authorized accounts.',
    'Protect Users and OTP_Log; define retention and cleanup procedures.',
    'Review Apps Script execute-as and web-app access settings with IT.',
    'Maintain backups and monitor price, photo and account administration.',
    'Test a restore procedure and document the production support owner.'
], GREEN, title_size=18, body_size=13.5)
notes(s, 'This is an important corporate presentation slide. The prototype contains client-side configuration and a short minimum password requirement, so production hardening is recommended. OTP logs and user records are sensitive. Agree on ownership, retention, backups and access control with IT before rolling the app out to a wider audience.')

# 20 — Demo and close
s = base_slide('Ten-Minute Live Demonstration Plan', '19 · Presentation runbook', 20, dark=True,
               subtitle='A focused sequence that proves the workflow without exposing credentials or production data.')
demo = [
    ('00:00–01:00', 'Access', 'Login with a prepared test account.'),
    ('01:00–02:30', 'Inventory', 'Search, open details and explain valuation.'),
    ('02:30–04:00', 'Admin photo', 'Upload a test image and location; explain persistence.'),
    ('04:00–06:30', 'Lens', 'Compare user/reference photos, inspect breakdown and confirm.'),
    ('06:30–08:00', 'Architecture', 'Explain Sheets, Drive, MailApp and getimg.'),
    ('08:00–09:00', 'Deployment', 'Show GitHub Pages and the installed PWA.'),
    ('09:00–10:00', 'Close', 'State accuracy limits, security next steps and take questions.')
]
for i, (tm, t, b) in enumerate(demo):
    y = 1.55 + i*0.62
    text(s, 0.67, y+0.07, 1.35, 0.27, tm, 10.5, CYAN, True, MONO)
    text(s, 2.15, y+0.03, 1.43, 0.30, t, 13.5, WHITE, True)
    text(s, 3.62, y+0.04, 5.23, 0.38, b, 11.8, 'C4D3E3')
    if i < 6: line(s, 0.67, y+0.52, 8.87, y+0.52, NAVY3, 0.7)
card(s, 9.18, 1.55, 3.50, 3.84, 'Demo preparation', [
    'Use one test account and three known test parts.',
    'Prepare clear admin reference photos.',
    'Pre-install the PWA and allow camera access.',
    'Verify internet, backend and getimg before starting.',
    'Keep backup screenshots and hide all secrets.'
], CYAN, True, 16.5, 12.0)
rect(s, 9.18, 5.64, 3.50, 0.78, NAVY3, None, True)
text(s, 9.38, 5.72, 3.10, 0.60, 'FAST SEARCH\nPERSISTENT PHOTOS\nASSISTED IDENTIFICATION', 10.8, CYAN, True, align=PP_ALIGN.CENTER, spacing=0.9)
notes(s, 'Rehearse this sequence and keep to the timing. The Lens demonstration should use a known part and a known optional prefix. End by repeating three messages: inventory is easy to search, admin photos are persistent, and Lens provides suggestions that require confirmation. Close with the copyright and contact email already shown in the footer.')

# Safety: remove any accidental Telugu characters from notes or slide text is checked externally.
OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f'Created {OUT} with {len(prs.slides)} slides')
