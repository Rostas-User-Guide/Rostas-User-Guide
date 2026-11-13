"""Draws setup-map.svg, the "Getting started with Rostas" map in Part B.

Edit the steps below, then run:  python3 make_setup_map.py
and re-export setup-map.png (see the export note inside the SVG).
Sized so the smallest text is 34px at 1800 wide (about 16px in the guide).
"""
from html import escape

W = 1800
NAVY, GREY, GOLD, ARROW, BG = "#1a2e4a", "#5a6a7e", "#b8932a", "#2b4570", "#f7f9fb"
COLS = [50, 650, 1250]
BW, BH = 520, 190
ROW_GAP = 70

out = []
add = out.append


def text(x, y, s, size, weight=400, fill=NAVY, anchor="start"):
    add(f'<text x="{x}" y="{y}" font-size="{size}" font-weight="{weight}" fill="{fill}" text-anchor="{anchor}">{escape(s)}</text>')


def block(col, y, title, lines, style="normal", badge=None):
    x = COLS[col]
    fill, stroke, sw, dash = "#ffffff", NAVY, 3, ""
    if style == "critical":
        fill, stroke, sw = "#fdf3dc", GOLD, 5
    if style == "optional":
        dash = ' stroke-dasharray="14 9"'
    add("<g>")
    add(f'<rect x="{x}" y="{y}" width="{BW}" height="{BH}" rx="18" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{dash}/>')
    total = 48 + 42 * len(lines)
    base = y + (BH - total) / 2 + 38
    cx = x + BW / 2
    text(cx, round(base), title, 40, 700, NAVY, "middle")
    for i, ln in enumerate(lines):
        text(cx, round(base + 50 + 42 * i), ln, 34, 400, GREY, "middle")
    if badge:
        add(f'<rect x="{x + 22}" y="{y - 23}" width="86" height="46" rx="23" fill="{GOLD}" stroke="{BG}" stroke-width="4"/>')
        text(x + 65, y + 10, badge, 30, 700, "#ffffff", "middle")
    add("</g>")


def row_arrows(y):
    cy = y + BH / 2
    for c in range(2):
        add(f'<line x1="{COLS[c] + BW + 3}" y1="{cy}" x2="{COLS[c + 1] - 4}" y2="{cy}"/>')


def wrap_arrow(r1):
    """From the end of row 1 down and back to the start of row 2."""
    top = r1 + BH + 3
    mid = r1 + BH + ROW_GAP / 2
    x3, x1 = COLS[2] + BW / 2, COLS[0] + BW / 2
    add(f'<path d="M{x3},{top} L{x3},{mid} L{x1},{mid} L{x1},{r1 + BH + ROW_GAP - 4}"/>')


def heading(y, bold, rest):
    add(f'<text x="{COLS[0]}" y="{y}" font-size="40" font-weight="700" fill="{GOLD}">{escape(bold)}'
        f'<tspan font-weight="400" font-size="34" fill="{GREY}">{escape(rest)}</tspan></text>')


def section(hy, title, rest, steps, top_pad=30):
    heading(hy, title, rest)
    r1 = hy + top_pad
    r2 = r1 + BH + ROW_GAP
    for i, s in enumerate(steps):
        if s is None:
            continue
        block(i % 3, r1 if i < 3 else r2, *s)
    add(f'<g stroke="{ARROW}" stroke-width="3" fill="none" marker-end="url(#arr)">')
    row_arrows(r1)
    wrap_arrow(r1)
    n2 = sum(1 for s in steps[3:] if s is not None)
    if n2 == 3:
        row_arrows(r2)
    add("</g>")
    return r1, r2


# ---- page -------------------------------------------------------------
H = 2720
add(f'<rect width="{W}" height="{H}" fill="{BG}"/>')
text(W / 2, 82, "Getting started with Rostas", 60, 700, NAVY, "middle")
text(W / 2, 142, "Copy it, set it up, build your first roster, then send it.", 36, 400, GREY, "middle")

add(f'<rect x="{COLS[0]}" y="178" width="{COLS[2] + BW - COLS[0]}" height="80" rx="18" fill="{NAVY}"/>')
add(f'<text x="{W / 2}" y="230" font-size="34" text-anchor="middle" fill="#ffffff"><tspan font-weight="700">Before you start: </tspan>'
    f'{escape("use a computer, have your licence key, and keep one Google account per browser")}</text>')

# legend
ly = 300
add(f'<rect x="{COLS[0]}" y="{ly}" width="70" height="40" rx="10" fill="#fdf3dc" stroke="{GOLD}" stroke-width="5"/>')
text(COLS[0] + 88, ly + 32, "Don't skip this", 34, 400, NAVY)
add(f'<rect x="600" y="{ly}" width="70" height="40" rx="10" fill="#ffffff" stroke="{NAVY}" stroke-width="3" stroke-dasharray="14 9"/>')
text(688, ly + 32, "Optional", 34, 400, NAVY)
add(f'<rect x="1000" y="{ly}" width="70" height="40" rx="10" fill="#fff1c2"/>')
text(1088, ly + 32, "A decision to make", 34, 400, NAVY)

section(420, "1 · Make your copy of Rostas", "  (about 10 minutes, once)", [
    ("Make your copy", ["From your welcome page"]),
    ("Share for updates", ["rostasapp@gmail.com", "as Editor"], "optional"),
    ("Deploy as Web app", ["Extensions → Apps Script", "→ Deploy"]),
    ("Google's warning", ["Advanced → Go to… (unsafe)"]),
    ("Tick Select all", ["then Continue"], "critical"),
    ("Open Rostas", ["Open the Web app URL", "and bookmark it"]),
])

section(980, "2 · Setup Wizard", "  (about 5 minutes, once)", [
    ("1  Licence", ["Key from your welcome email"]),
    ("2  Organisation", ["Name, timezone, country"]),
    ("3  Your details", ["Name, email, and mobile", "with country code"]),
    ("4  SMS test", ["Or Continue without SMS"]),
    ("5  Done", ["Go to Build"]),
    ("Coordinator link", ["Emailed to you:", "bookmark it"]),
])

r1, r2 = section(1540, "3 · Set up your first roster", "  (20–30 minutes, once per roster)", [
    ("Name, roles, season", ["What the roster is", "and when it runs"], "normal", "1A"),
    ("Get volunteers in", ["Connect your database,", "or import a list"], "normal", "1B"),
    ("Set up each volunteer", ["Roles, how often,", "who they serve with"], "critical", "1C"),
    ("Review & create", ["Check, then Create Roster"], "normal", "1D"),
    None, None,
], top_pad=40)
# decision note beside 1D
nx, nw = COLS[1], COLS[2] + BW - COLS[1]
add(f'<rect x="{nx}" y="{r2}" width="{nw}" height="{BH}" rx="18" fill="#fff1c2"/>')
text(nx + 36, r2 + 46, "Decide at 1C", 40, 700, NAVY)
for i, ln in enumerate([
    "•  Set roles yourself, or ask volunteers at 2C",
    "•  Ask availability at 2D? Yes, the first time",
    "•  A volunteer with no roles is left out of the build",
]):
    text(nx + 36, r2 + 90 + 38 * i, ln, 34, 400, NAVY)

r1, r2 = section(2100, "4 · Roster cycle: build, check, send", "  (about 10 minutes each time)", [
    ("Confirm Dates", ["Which dates to build"], "normal", "2A"),
    ("Roles", ["How many each role needs"], "normal", "2B"),
    ("Role Prefs", ["Ask which roles suit them"], "optional", "2C"),
    ("Availability", ["Poll the dates", "people can't do"], "optional", "2D"),
    ("Build & Check", ["Check the Board.", "Nothing is sent yet"], "normal", "2E"),
    ("Notify", ["Test modes off,", "then send"], "critical", "2F"),
], top_pad=40)
# next-time loop: 2F -> down, left, up the gutter, into 2A's left side
fx = COLS[2] + BW / 2
low = r2 + BH + 36
a_mid = r1 + BH / 2
add(f'<path d="M{fx},{r2 + BH + 3} L{fx},{low} L20,{low} L20,{a_mid} L{COLS[0] - 5},{a_mid}" '
    f'stroke="{GOLD}" stroke-width="4" stroke-dasharray="14 9" fill="none" marker-end="url(#arrGold)"/>')
text(COLS[0], low + 58, "↺ Next time, start at 2A. Only volunteers new since the last build need setting up.", 34, 700, GOLD)

svg = (
    f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}" '
    f'font-family="\'DM Sans\', Arial, Helvetica, sans-serif">\n'
    "  <!-- 13 Nov 2026 — Getting started map for the top of Part B, from the\n"
    "       Rostas Getting Started whiteboard. Generated by make_setup_map.py:\n"
    "       edit that, not this file. Three steps per row so the text stays\n"
    "       readable in the guide's 840px column on a 14-inch screen.\n"
    "       Export: wrap in HTML that loads DM Sans from Google Fonts and screenshot\n"
    f"       with headless Chrome at {W}x{H}. -->\n"
    "  <defs>\n"
    f'    <marker id="arr" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{ARROW}"/></marker>\n'
    f'    <marker id="arrGold" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="{GOLD}"/></marker>\n'
    "  </defs>\n  "
    + "\n  ".join(out)
    + "\n</svg>\n"
)
open("setup-map.svg", "w").write(svg)
print("wrote setup-map.svg", W, H)
