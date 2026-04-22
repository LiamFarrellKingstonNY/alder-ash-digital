from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# Colors
FOREST_GREEN = HexColor('#254D3B')
LIGHT_GREEN = HexColor('#7FB89A')
CREAM = HexColor('#F7F2E8')
AMBER = HexColor('#C4892A')
DARK_TEXT = HexColor('#1A1A1A')
MID_TEXT = HexColor('#3D3D3D')
LIGHT_AMBER = HexColor('#F5E6CC')
BOX_GREEN = HexColor('#1C3B2C')

W, H = letter  # 612 x 792


def draw_rounded_rect(c, x, y, w, h, radius=6, fill_color=None, stroke_color=None, stroke_width=0):
    c.saveState()
    if fill_color:
        c.setFillColor(fill_color)
    if stroke_color:
        c.setStrokeColor(stroke_color)
        c.setLineWidth(stroke_width)
    else:
        c.setLineWidth(0)
    c.roundRect(x, y, w, h, radius, fill=1 if fill_color else 0, stroke=1 if stroke_color else 0)
    c.restoreState()


def draw_text(c, text, x, y, font, size, color, align='left', max_width=None):
    c.saveState()
    c.setFont(font, size)
    c.setFillColor(color)
    if align == 'center':
        c.drawCentredString(x, y, text)
    elif align == 'right':
        c.drawRightString(x, y, text)
    else:
        c.drawString(x, y, text)
    c.restoreState()


def draw_wrapped_text(c, text, x, y, font, size, color, max_width, line_height=None):
    """Simple word-wrap for canvas text."""
    if line_height is None:
        line_height = size * 1.45
    c.saveState()
    c.setFont(font, size)
    c.setFillColor(color)
    words = text.split()
    line = ''
    for word in words:
        test = (line + ' ' + word).strip()
        if c.stringWidth(test, font, size) <= max_width:
            line = test
        else:
            c.drawString(x, y, line)
            y -= line_height
            line = word
    if line:
        c.drawString(x, y, line)
        y -= line_height
    c.restoreState()
    return y  # return final y position


def generate(output_path):
    c = canvas.Canvas(output_path, pagesize=letter)
    margin = 36  # 0.5 inch margins

    # ── HEADER BAND ──────────────────────────────────────────────────────────
    header_h = 108
    header_y = H - header_h
    c.setFillColor(FOREST_GREEN)
    c.rect(0, header_y, W, header_h, fill=1, stroke=0)

    # Thin amber accent line at bottom of header
    c.setFillColor(AMBER)
    c.rect(0, header_y, W, 3, fill=1, stroke=0)

    # Company name
    draw_text(c, 'Alder & Ash Digital', margin, H - 42, 'Helvetica-Bold', 26, white)

    # Tagline
    draw_text(c, 'Web Design & AI Automation for Local Businesses',
              margin, H - 64, 'Helvetica', 11.5, LIGHT_GREEN)

    # Location — right-aligned with a subtle leaf marker
    draw_text(c, 'Hudson Valley, NY', W - margin, H - 42,
              'Helvetica-Oblique', 10, LIGHT_GREEN, align='right')

    # Thin decorative vertical separator in header
    c.setStrokeColor(LIGHT_GREEN)
    c.setLineWidth(0.5)
    c.setDash([3, 5])
    c.line(W - margin - 110, H - 28, W - margin - 110, H - 82)
    c.setDash([])

    # ── BODY BACKGROUND ──────────────────────────────────────────────────────
    body_top = header_y - 3  # just below amber accent
    footer_h = 48
    body_h = body_top - footer_h
    c.setFillColor(CREAM)
    c.rect(0, footer_h, W, body_h, fill=1, stroke=0)

    # ── COLUMN LAYOUT ─────────────────────────────────────────────────────────
    left_col_w = W * 0.57
    right_col_x = left_col_w + margin * 0.6
    right_col_w = W - right_col_x - margin

    content_top = body_top - 22

    # ── LEFT COLUMN ───────────────────────────────────────────────────────────
    lx = margin

    # Headline
    draw_text(c, 'Is Your Website', lx, content_top, 'Helvetica-Bold', 19, FOREST_GREEN)
    draw_text(c, 'Working For You?', lx, content_top - 22, 'Helvetica-Bold', 19, AMBER)

    # Amber underline accent
    c.setFillColor(AMBER)
    c.rect(lx, content_top - 27, 90, 2.5, fill=1, stroke=0)

    # Body paragraph
    para_y = content_top - 52
    body_copy = (
        'Most local businesses are losing customers every day to competitors '
        'with faster, smarter websites. Your site should be your best salesperson — '
        'capturing leads, answering questions, and booking appointments even while you sleep.'
    )
    para_end_y = draw_wrapped_text(c, body_copy, lx, para_y,
                                   'Helvetica', 9.5, MID_TEXT,
                                   left_col_w - margin - 8, line_height=15)

    # "What We Build:" section
    section_y = para_end_y - 20

    # Small amber pill label
    draw_rounded_rect(c, lx, section_y - 2, 88, 16, radius=3, fill_color=AMBER)
    draw_text(c, 'WHAT WE BUILD', lx + 7, section_y + 3, 'Helvetica-Bold', 7.5, white)

    bullet_y = section_y - 22
    bullets = [
        ('AI-powered websites', 'that generate leads 24/7'),
        ('Automated follow-up systems', '(email + text)'),
        ('Smart chatbots', 'trained on your business'),
    ]

    for bold_part, rest in bullets:
        # Amber bullet diamond
        c.saveState()
        c.setFillColor(AMBER)
        cx2, cy2 = lx + 5, bullet_y + 4
        c.translate(cx2, cy2)
        c.rotate(45)
        c.rect(-3.5, -3.5, 7, 7, fill=1, stroke=0)
        c.restoreState()

        # Bold + normal text
        bold_w = c.stringWidth(bold_part + ' ', 'Helvetica-Bold', 9.5)
        draw_text(c, bold_part + ' ', lx + 16, bullet_y, 'Helvetica-Bold', 9.5, FOREST_GREEN)
        draw_text(c, rest, lx + 16 + bold_w, bullet_y, 'Helvetica', 9.5, MID_TEXT)

        bullet_y -= 20

    # ── RIGHT COLUMN — Quick Stats Box ────────────────────────────────────────
    rx = right_col_x
    box_top = content_top + 6
    box_h = 178
    box_y = box_top - box_h

    # Shadow effect
    draw_rounded_rect(c, rx + 3, box_y - 3, right_col_w, box_h,
                      radius=8, fill_color=HexColor('#D4C9B0'))
    # Main box
    draw_rounded_rect(c, rx, box_y, right_col_w, box_h,
                      radius=8, fill_color=FOREST_GREEN)

    # "Quick Stats" header in box
    draw_text(c, 'QUICK STATS', rx + right_col_w / 2,
              box_top - 20, 'Helvetica-Bold', 9, LIGHT_GREEN, align='center')

    # Thin separator line
    c.setStrokeColor(LIGHT_GREEN)
    c.setLineWidth(0.5)
    c.line(rx + 10, box_top - 26, rx + right_col_w - 10, box_top - 26)

    stats = [
        ('5 days', 'Average build time'),
        ('3x', 'More leads for clients'),
        ('98%', 'Mobile performance score'),
        ('$250/mo', 'Retainer starts at'),
    ]

    stat_y = box_top - 44
    for number, label in stats:
        draw_text(c, number, rx + right_col_w / 2, stat_y,
                  'Helvetica-Bold', 18, AMBER, align='center')
        draw_text(c, label, rx + right_col_w / 2, stat_y - 14,
                  'Helvetica', 7.5, LIGHT_GREEN, align='center')
        # Divider (except after last)
        if (number, label) != stats[-1]:
            c.setStrokeColor(HexColor('#3A6B54'))
            c.setLineWidth(0.5)
            c.line(rx + 18, stat_y - 22, rx + right_col_w - 18, stat_y - 22)
        stat_y -= 38

    # ── PRICING SECTION ───────────────────────────────────────────────────────
    pricing_top = min(bullet_y - 28, box_y - 22)

    # Section label
    draw_rounded_rect(c, margin, pricing_top - 2, 104, 16, radius=3, fill_color=FOREST_GREEN)
    draw_text(c, 'PRICING & PACKAGES', margin + 7, pricing_top + 3,
              'Helvetica-Bold', 7.5, white)

    # Thin full-width rule
    c.setStrokeColor(HexColor('#D4C9B0'))
    c.setLineWidth(0.75)
    c.line(margin, pricing_top - 6, W - margin, pricing_top - 6)

    pricing_card_top = pricing_top - 14
    usable_w = W - 2 * margin
    gap = 10
    card_w = (usable_w - 2 * gap) / 3
    card_h = 96

    packages = [
        {
            'name': 'Launch Package',
            'price': '$3,500',
            'period': 'one-time',
            'features': ['5-page website', 'On-page SEO', 'AI chatbot included'],
            'highlight': False,
        },
        {
            'name': 'Growth System',
            'price': '$5,500+',
            'period': 'one-time',
            'features': ['Marketing automation', 'CRM integration', 'Online booking'],
            'highlight': True,
        },
        {
            'name': 'Growth Retainer',
            'price': '$250',
            'period': 'per month',
            'features': ['Hosting & security', 'Monthly updates', 'Performance tuning'],
            'highlight': False,
        },
    ]

    for i, pkg in enumerate(packages):
        cx = margin + i * (card_w + gap)
        cy = pricing_card_top - card_h

        if pkg['highlight']:
            # Shadow
            draw_rounded_rect(c, cx + 3, cy - 3, card_w, card_h,
                               radius=7, fill_color=HexColor('#A07020'))
            draw_rounded_rect(c, cx, cy, card_w, card_h,
                               radius=7, fill_color=AMBER)
            name_color = white
            price_color = white
            period_color = HexColor('#F5E6CC')
            feat_color = white
            bullet_color = white
            # "POPULAR" badge
            draw_rounded_rect(c, cx + card_w / 2 - 24, pricing_card_top - 1, 48, 13,
                               radius=6, fill_color=FOREST_GREEN)
            draw_text(c, 'POPULAR', cx + card_w / 2, pricing_card_top + 4,
                      'Helvetica-Bold', 6.5, white, align='center')
        else:
            draw_rounded_rect(c, cx + 2, cy - 2, card_w, card_h,
                               radius=7, fill_color=HexColor('#D4C9B0'))
            draw_rounded_rect(c, cx, cy, card_w, card_h,
                               radius=7, fill_color=white)
            name_color = FOREST_GREEN
            price_color = AMBER
            period_color = MID_TEXT
            feat_color = MID_TEXT
            bullet_color = AMBER

        text_x = cx + 10
        ty = pricing_card_top - 18

        draw_text(c, pkg['name'], text_x, ty, 'Helvetica-Bold', 9, name_color)
        ty -= 18
        draw_text(c, pkg['price'], text_x, ty, 'Helvetica-Bold', 17, price_color)
        pw = c.stringWidth(pkg['price'], 'Helvetica-Bold', 17)
        draw_text(c, ' ' + pkg['period'], text_x + pw, ty + 4,
                  'Helvetica', 7, period_color)

        # Thin separator
        sep_color = HexColor('#E8E0D0') if not pkg['highlight'] else HexColor('#D4922E')
        c.setStrokeColor(sep_color)
        c.setLineWidth(0.5)
        c.line(text_x, ty - 7, cx + card_w - 10, ty - 7)
        ty -= 18

        for feat in pkg['features']:
            # Bullet
            c.setFillColor(bullet_color)
            c.circle(text_x + 3, ty + 3, 2, fill=1, stroke=0)
            draw_text(c, feat, text_x + 10, ty, 'Helvetica', 8, feat_color)
            ty -= 14

    # ── FOOTER BAND ───────────────────────────────────────────────────────────
    c.setFillColor(FOREST_GREEN)
    c.rect(0, 0, W, footer_h, fill=1, stroke=0)

    # Amber accent top of footer
    c.setFillColor(AMBER)
    c.rect(0, footer_h - 3, W, 3, fill=1, stroke=0)

    # CTA text
    footer_mid = footer_h / 2
    cta = 'Get Your Free Website Audit'
    cta_w = c.stringWidth(cta, 'Helvetica-Bold', 11)
    draw_text(c, cta, W / 2, footer_mid + 9, 'Helvetica-Bold', 11, white, align='center')

    # Contact line with separators
    sep = '  ·  '
    contact = 'hello@alderandash.com' + sep + 'alderandash.com'
    draw_text(c, contact, W / 2, footer_mid - 8,
              'Helvetica', 9, LIGHT_GREEN, align='center')

    c.save()
    print(f'Saved: {output_path}')


if __name__ == '__main__':
    import os
    out = os.path.expanduser('~/Desktop/alder-ash-digital/Alder-Ash-One-Pager.pdf')
    generate(out)
