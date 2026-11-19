"""Build the ILLXL shortcut reference PDF from the live sources.

    python -I tools/make-pdf.py <preset name> <output.pdf>

Commands, groups and descriptions come from the catalog in src/modBindings.bas
(the same text the ribbon tooltips show). Keys, colors and styles come from
presets/<preset>.ini. So the PDF can never drift from the add-in again.
The page is HTML printed to PDF by Microsoft Edge (built into Windows).
"""
import base64
import html
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
GROUP_TITLES = {
    "Core": ("Core", "Settings, performance mode, navigation and formula references."),
    "Formats": ("Number Formats", "Cycle formats - press again for the next step. Restarts when the selection changes."),
    "Formulas": ("Formulas", "Financial formulas in one keystroke. Picked cells can be on any sheet."),
    "Styles": ("Styles & Layout", "Colors, fonts, input and header styles, indent, headers and checks."),
    "Borders": ("Borders", "Edge cycles restart at step 1 when the selection changes."),
    "Unit Tags": ("Unit Tags", "FAST-standard tags. Typed text only - formulas are never touched."),
    "Model Tools": ("Model Tools", "Audit a model, wrap formulas, build an index, and get a model ready to send."),
}


def catalog():
    src = (ROOT / "src" / "modBindings.bas").read_text(encoding="utf-8")
    out = []
    for m in re.finditer(r'Cmd "(\w+)", "([^"]+)", "([^"]+)", _\s*"((?:[^"]|"")*)"', src):
        out.append(dict(macro=m[1], group=m[2], label=m[3], desc=m[4].replace('""', '"')))
    return out


def read_ini(name):
    data, sec = {}, None
    for line in (ROOT / "presets" / f"{name}.ini").read_text(encoding="utf-8-sig").splitlines():
        s = line.strip()
        if not s or s.startswith(";"):
            continue
        if s.startswith("[") and "]" in s:
            sec = s[1:s.index("]")].strip()
            data[sec] = {}
        elif sec and "=" in s:
            k, v = s.split("=", 1)
            data[sec][k.strip()] = v.strip()
    return data


def items(ini, sec):
    d = ini.get(sec, {})
    return [d[k] for k in sorted((k for k in d if k.isdigit()), key=int)]


def key_html(k):
    if not k:
        return '<span class="none">-</span>'
    return "<span class='plus'>+</span>".join(f"<kbd>{html.escape(p)}</kbd>" for p in k.split("+"))


def swatch(color, label=""):
    c = color.strip().lower()
    if c == "none":
        return '<span class="sw none-sw"></span>'
    if c == "stripe":
        return '<span class="sw stripe"></span>'
    return f'<span class="sw" style="background:{html.escape(color)}">{html.escape(label)}</span>'


def build_html(preset):
    ini = read_ini(preset)
    keys = ini.get("Keys", {})
    logo = base64.b64encode((ROOT / "assets" / "logo-256.png").read_bytes()).decode()
    cmds = catalog()
    groups = {}
    for g in GROUP_TITLES:
        rows = "".join(
            f"<tr><td class='k'>{key_html(keys.get(c['macro'], ''))}</td>"
            f"<td><b>{html.escape(c['label'])}</b><div class='d'>{html.escape(c['desc'])}</div></td></tr>"
            for c in cmds if c["group"] == g)
        title, sub = GROUP_TITLES[g]
        groups[g] = f"<section><h2>{title}</h2><p class='sub'>{sub}</p><table>{rows}</table></section>"
    # one printed page each, so every page gets the running header
    command_pages = [["Core", "Formulas"], ["Formats"], ["Styles"], ["Borders", "Unit Tags"], ["Model Tools"]]

    def cycle_table(sec, kind):
        rows = []
        for i, it in enumerate(items(ini, sec), 1):
            p = [x.strip() for x in it.split("|")]
            name, cols = p[0], p[1:] or [p[0]]
            if kind == "input":
                sample = f'<span class="cell" style="background:{cols[0]};color:{cols[1]};border:1px dotted {cols[2] if len(cols) > 2 else "#000"}">1,234</span>'
            elif kind == "header":
                sample = f'<span class="cell" style="background:{cols[0]};color:{cols[1]};font-weight:700">FY27</span>'
            else:
                sample = swatch(cols[0])
            rows.append(f"<tr><td class='n'>{i}</td><td>{html.escape(name)}</td><td>{sample}</td>"
                        f"<td class='hex'>{' / '.join(html.escape(c) for c in cols)}</td></tr>")
        return "<table class='ref'>" + "".join(rows) + "</table>"

    def fmt_list(sec):
        return "<ol class='fmt'>" + "".join(f"<li><code>{html.escape(f)}</code></li>" for f in items(ini, sec)) + "</ol>"

    ac = ini.get("AutoColor", {})
    autocolor = "".join(f"<li>{swatch(ac.get(k, '#000'))} {lbl}</li>" for k, lbl in [
        ("Hardcode", "Hardcodes (typed numbers)"), ("Formula", "Formulas"), ("SameSheetLink", "Links on the same sheet"),
        ("CrossSheetLink", "Links to other sheets"), ("ExternalLink", "Links to other files")])

    css = """
    @page { size: Letter; margin: 0.45in 0.5in; }
    * { box-sizing: border-box; }
    body { font-family: 'Segoe UI', Arial, sans-serif; color: #111827; font-size: 9.2pt; margin: 0; }
    .cover { height: 9.9in; background: linear-gradient(135deg, #0F172A, #1E3A8A); color: #fff; border-radius: 14px;
             padding: 0.9in 0.7in; display: flex; flex-direction: column; justify-content: center; }
    .cover img { width: 120px; margin-bottom: 26px; }
    .cover h1 { font-size: 46pt; margin: 0; letter-spacing: 1px; }
    .cover .tag { color: #38BDF8; font-size: 16pt; margin: 6px 0 26px; }
    .cover p { font-size: 11pt; max-width: 5.6in; line-height: 1.5; color: #CBD5E1; }
    .stats { display: flex; gap: 40px; margin-top: 30px; }
    .stats b { display: block; font-size: 30pt; color: #fff; } .stats span { color: #93C5FD; font-size: 9pt; letter-spacing: 1px; }
    .meta { margin-top: auto; color: #94A3B8; font-size: 9pt; }
    header.run { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #2563EB;
                 padding-bottom: 6px; margin-bottom: 10px; }
    header.run img { width: 26px; vertical-align: middle; margin-right: 8px; } header.run b { font-size: 12pt; }
    header.run span { color: #6B7280; }
    h2 { font-size: 13pt; margin: 14px 0 2px; color: #0F172A; } .sub { color: #6B7280; margin: 0 0 6px; font-size: 8.6pt; }
    section { page-break-inside: avoid; }
    table { width: 100%; border-collapse: collapse; }
    td { padding: 4px 6px; border-bottom: 1px solid #E5E7EB; vertical-align: top; }
    td.k { width: 2.15in; white-space: nowrap; }
    .d { color: #4B5563; font-size: 8.4pt; margin-top: 1px; }
    kbd { display: inline-block; background: #F3F4F6; border: 1px solid #D1D5DB; border-bottom-width: 2px; border-radius: 4px;
          padding: 0 5px; font: 600 8.2pt Consolas, monospace; color: #111827; }
    .plus { color: #9CA3AF; margin: 0 2px; } .none { color: #9CA3AF; }
    .cols { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 26px; }
    .ref td { padding: 3px 6px; } .ref td.n { width: 18px; color: #9CA3AF; } .hex { font: 8pt Consolas, monospace; color: #6B7280; }
    .sw { display: inline-block; width: 30px; height: 14px; border: 1px solid #D1D5DB; border-radius: 3px; vertical-align: middle; }
    .none-sw { background: #fff; position: relative; } .stripe { background: repeating-linear-gradient(45deg, #fff 0 3px, #A6A6A6 3px 5px); }
    .cell { display: inline-block; padding: 1px 10px; font-size: 8.5pt; }
    ol.fmt { margin: 0 0 6px 18px; padding: 0; } ol.fmt code { font: 8pt Consolas, monospace; }
    ul.ac { list-style: none; padding: 0; margin: 0; } ul.ac li { margin: 3px 0; }
    .tips li { margin: 4px 0; }
    """
    run = (f"<header class='run'><div><img src='data:image/png;base64,{logo}'><b>ILLXL</b> &nbsp;Shortcut Reference</div>"
           f"<span>Preset: {html.escape(preset)}</span></header>")
    pages = f"""
    <div class="cover">
      <img src="data:image/png;base64,{logo}">
      <h1>ILLXL</h1><div class="tag">The Illest Excel Add-In for Finance Pros</div>
      <p>{len(cmds)} commands on the keyboard and the ILLXL ribbon tab - number formats, styles, borders, formulas,
      unit tags and model tools. Every list, color and key comes from a preset you can edit in the Settings window
      (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Shift</kbd>+<kbd>S</kbd>).</p>
      <div class="stats"><div><b>{len(cmds)}</b><span>COMMANDS</span></div><div><b>3</b><span>PRESETS</span></div>
      <div><b>FREE</b><span>ALWAYS</span></div></div>
      <div class="meta">Preset shown: {html.escape(preset)} &nbsp;|&nbsp; github.com/bzilles015 &nbsp;|&nbsp; linkedin.com/in/brucezacharyilles</div>
    </div>
    {''.join(f'<div style="page-break-before: always"></div>{run}' + ''.join(groups[g] for g in pg) for pg in command_pages)}
    <div style="page-break-before: always"></div>{run}
    <h2>Your preset: {html.escape(preset)}</h2><p class="sub">{html.escape(ini.get('Info', {}).get('Description', ''))}</p>
    <div class="cols">
      <section><h2>Input styles</h2>{cycle_table('InputStyles', 'input')}</section>
      <section><h2>Header styles</h2>{cycle_table('HeaderStyles', 'header')}</section>
      <section><h2>Fill cycle</h2>{cycle_table('Fills', 'fill')}</section>
      <section><h2>Font colors</h2>{cycle_table('FontColors', 'fill')}</section>
      <section><h2>Tab colors</h2>{cycle_table('TabColors', 'fill')}</section>
      <section><h2>Auto-Color</h2><ul class="ac">{autocolor}</ul></section>
    </div>
    <div style="page-break-before: always"></div>{run}
    <div class="cols">
      <section><h2>Number formats</h2>{fmt_list('NumberFormats')}</section>
      <section><h2>Currency formats</h2>{fmt_list('CurrencyFormats')}</section>
      <section><h2>Percent formats</h2>{fmt_list('PercentFormats')}</section>
      <section><h2>Date formats</h2>{fmt_list('DateFormats')}</section>
      <section><h2>Other formats</h2>{fmt_list('OtherFormats')}</section>
      <section><h2>Unit tags</h2><p>{html.escape(ini.get('UnitTags', {}).get('Value', ''))}<br>
        {html.escape(ini.get('UnitTags', {}).get('Duration', ''))}<br>{html.escape(ini.get('UnitTags', {}).get('Rate', ''))}</p></section>
    </div>
    <section><h2>How it works</h2><ul class="tips">
      <li><b>Cycles</b> - press a cycle key again for the next step; it starts over when you select different cells. The ribbon dropdowns jump straight to any step.</li>
      <li><b>Undo</b> - Ctrl+Z undoes any ILLXL command (values, formulas and formatting) for selections up to 5,000 cells.</li>
      <li><b>Whole columns</b> - selecting a whole column works on the rows you actually use, so it is instant.</li>
      <li><b>Performance Mode</b> - manual calc, events and conditional formatting off for heavy formatting. Undo is off while it is on.</li>
      <li><b>Presets</b> - switch everything (formats, colors, keys) in one click from the ribbon or Settings. Share one by exporting its .ini file.</li>
      <li><b>Keys switched off</b> - {html.escape(ini.get('General', {}).get('DisabledKeys', ''))} (hit by accident far more often than on purpose).</li>
    </ul></section>
    """
    return f"<!doctype html><html><head><meta charset='utf-8'><title>ILLXL Shortcut Reference</title><style>{css}</style></head><body>{pages}</body></html>"


def main(preset, out):
    out = Path(out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as d:
        page = Path(d) / "ref.html"
        page.write_text(build_html(preset), encoding="utf-8")
        subprocess.run([EDGE, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                        f"--user-data-dir={Path(d) / 'edge'}", f"--print-to-pdf={out}", page.as_uri()],
                       check=True, capture_output=True, timeout=120)
    print(f"PDF: {out}  ({out.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
