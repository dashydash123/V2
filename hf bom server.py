#!/usr/bin/env python3
"""
HuggingFace Model BOM - local backend.

Why this exists
---------------
The tool needs data from huggingface.co, api.github.com, pypi.org, wikidata and
arxiv. When a browser sits behind an isolation proxy (e.g. Zscaler Browser
Isolation), fetches made *from the page* are blocked. Python's own HTTPS calls
are not subject to browser isolation, so this script fetches the external URLs
server-side; the page only ever talks to http://localhost.

All the BOM logic still lives in project.html. This file only:
  1. serves project.html, and
  2. exposes /proxy?url=... which fetches the upstream URL in Python and
     returns the response (status + body) to the page.

Standard library only - nothing to pip install.
"""

import http.server
import socketserver
import socket
import urllib.request
import urllib.parse
import urllib.error
import ssl
import json
import os
import re
import sys
import zipfile
import io
import threading
import webbrowser
from xml.sax.saxutils import escape as _xesc

HERE = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(HERE, "project.html")
DEFAULT_UA = "HF-BOM/1.0"
TIMEOUT = 30  # seconds per upstream request


def build_ssl_context():
    """
    Verify TLS the same way the machine already does, so the corporate
    SSL-inspection root (installed in the OS trust store) is trusted with no
    manual cert wrangling.

    Preference order:
      1. `truststore` package  -> uses the OS trust store natively (if present)
      2. REQUESTS_CA_BUNDLE / SSL_CERT_FILE env var pointing at a CA bundle
      3. Python default context -> on Windows this loads the Windows cert
         store, which already contains the corporate root.
    """
    try:
        import truststore  # optional; not required
        return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except Exception:
        pass

    ctx = ssl.create_default_context()
    ca = os.environ.get("REQUESTS_CA_BUNDLE") or os.environ.get("SSL_CERT_FILE")
    if ca and os.path.exists(ca):
        try:
            ctx.load_verify_locations(ca)
        except Exception:
            pass
    return ctx


SSL_CTX = build_ssl_context()


# ---------------------------------------------------------------------------
# Excel (.xlsx) generation - pure standard library (an xlsx is a zip of XML).
# Mirrors the confidence colour coding of the original in-browser export.
# ---------------------------------------------------------------------------

CONF_STYLE = {"HIGH": 2, "MEDIUM": 3, "LOW": 4, "UNRESOLVED": 5}  # -> cellXfs index

_URL_RE = re.compile(r"https?://[^\s)\]]+")


def _first_url(text):
    if not text:
        return ""
    m = _URL_RE.search(text)
    return m.group(0).rstrip(".,;)") if m else ""


def _col_letter(idx):  # 0 -> A
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell(col, row, text, style):
    ref = "%s%d" % (_col_letter(col), row)
    txt = _xesc("" if text is None else str(text))
    return ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, style, txt))


def _sheet_xml(headers, rows, conf_col=None, widths=None):
    """rows: list of lists of strings. conf_col: index whose value picks the row colour."""
    ncols = len(headers)
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    last = "%s%d" % (_col_letter(ncols - 1), len(rows) + 1)
    out.append('<dimension ref="A1:%s"/>' % last)
    if widths:
        out.append("<cols>")
        for i, w in enumerate(widths):
            out.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, w))
        out.append("</cols>")
    out.append("<sheetData>")
    # header row (style 1)
    out.append('<row r="1">')
    for c, h in enumerate(headers):
        out.append(_cell(c, 1, h, 1))
    out.append("</row>")
    # data rows
    for ri, row in enumerate(rows, start=2):
        style = 6  # plain wrap
        if conf_col is not None and conf_col < len(row):
            style = CONF_STYLE.get(str(row[conf_col]).upper(), 6)
        out.append('<row r="%d">' % ri)
        for c in range(ncols):
            val = row[c] if c < len(row) else ""
            out.append(_cell(c, ri, val, style))
        out.append("</row>")
    out.append("</sheetData></worksheet>")
    return "".join(out)


def _summary_sheet_xml(model_id, model_license):
    rows = [
        ("A1", "Model BOM \u2014 " + model_id, 7),
        ("A3", "Model License", 8), ("B3", model_license, 6),
        ("A4", "Confidence legend", 8),
        ("B4", "GREEN = HIGH  |  YELLOW = MEDIUM  |  ORANGE = LOW  |  RED = UNRESOLVED", 6),
    ]
    body = []
    for ref, text, style in rows:
        col = ord(ref[0]) - 65
        r = int(ref[1:])
        body.append('<row r="%d">%s</row>' % (r, _cell(col, r, text, style)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<dimension ref="A1:B6"/>'
            '<cols><col min="1" max="1" width="24" customWidth="1"/>'
            '<col min="2" max="2" width="70" customWidth="1"/></cols>'
            '<sheetData>' + "".join(body) + '</sheetData></worksheet>')


_STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="4">'
    '<font><sz val="11"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    '<font><b/><sz val="13"/><name val="Calibri"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    '</fonts>'
    '<fills count="7">'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFCC99"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF2F5597"/></patternFill></fill>'
    '</fills>'
    '<borders count="1"><border/></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="9">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '</styleSheet>'
)


def build_xlsx(bom):
    model_id = bom.get("modelId", "model")
    model_license = bom.get("modelLicense", "")

    sheets = []  # (name, xml)
    sheets.append(("Summary", _summary_sheet_xml(model_id, model_license)))

    # Dataset licenses
    ds = bom.get("datasetLicenses", {}) or {}
    ds_rows = []
    for ds_id, info in ds.items():
        info = info or {}
        src = info.get("source", "")
        ds_rows.append([ds_id, info.get("license", ""), info.get("confidence", "UNRESOLVED"),
                        info.get("via", ""), src, _first_url(src)])
    sheets.append(("Dataset Licenses", _sheet_xml(
        ["Dataset", "License", "Confidence", "Resolved Via", "Source / Notes", "Link"],
        ds_rows, conf_col=2, widths=[40, 40, 12, 40, 40, 40])))

    # arXiv-derived datasets (only if present)
    ax = bom.get("arxivDatasets", {}) or {}
    if ax:
        ax_rows = []
        for name, info in ax.items():
            info = info or {}
            src = info.get("source", "")
            ax_rows.append([name, info.get("license", ""), info.get("confidence", "UNRESOLVED"),
                            info.get("via", ""), src, _first_url(src)])
        sheets.append(("arXiv Paper Datasets", _sheet_xml(
            ["Dataset (text-mined name)", "License", "Confidence", "Resolved Via", "Source / Notes", "Link"],
            ax_rows, conf_col=2, widths=[40, 40, 12, 40, 40, 40])))

    # Dependencies
    deps = bom.get("dependencies", []) or []
    dep_rows = []
    for d in deps:
        d = d or {}
        src = d.get("source", "")
        dep_rows.append([d.get("library", ""), d.get("version", ""), d.get("license", ""),
                         d.get("confidence", "UNRESOLVED"), src, d.get("depSource", ""), _first_url(src)])
    sheets.append(("Dependencies", _sheet_xml(
        ["Library", "Version", "License", "Confidence", "License Source", "Dependency Source", "Link"],
        dep_rows, conf_col=3, widths=[30, 22, 30, 12, 40, 30, 40])))

    n = len(sheets)
    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                     '<Default Extension="xml" ContentType="application/xml"/>',
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(1, n + 1):
        content_types.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
    content_types.append('</Types>')

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')

    wb_sheets, wb_rels = [], []
    wb_rels.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for i, (name, _) in enumerate(sheets, start=1):
        safe = _xesc(name[:31])
        wb_sheets.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (safe, i, i))
        wb_rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i))
    wb_rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n + 1))
    wb_rels.append('</Relationships>')

    workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets>' + "".join(wb_sheets) + '</sheets></workbook>')

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(content_types))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
        z.writestr("xl/styles.xml", _STYLES_XML)
        for i, (_, xml) in enumerate(sheets, start=1):
            z.writestr("xl/worksheets/sheet%d.xml" % i, xml)
    return buf.getvalue()


class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        sys.stderr.write("  " + (fmt % args) + "\n")

    def _send(self, code, body, content_type="text/plain; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        try:
            self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/", "/index.html", "/project.html"):
            return self._serve_html()
        if parsed.path == "/proxy":
            return self._do_proxy(parsed)
        self._send(404, "Not found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/export":
            return self._send(404, "Not found")
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length) if length else b"{}"
            bom = json.loads(raw.decode("utf-8"))
            data = build_xlsx(bom)
        except Exception as e:
            return self._send(500, "export error: " + str(e))
        self.send_response(200)
        self.send_header("Content-Type",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="bom.xlsx"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionResetError):
            pass

    def _serve_html(self):
        try:
            with open(HTML_FILE, "rb") as f:
                data = f.read()
        except FileNotFoundError:
            return self._send(500, "project.html not found next to this script.")
        self._send(200, data, "text/html; charset=utf-8")

    def _do_proxy(self, parsed):
        qs = urllib.parse.parse_qs(parsed.query)
        target = (qs.get("url") or [""])[0]
        if not target.startswith(("http://", "https://")):
            return self._send(400, "bad or missing url")

        headers = {"User-Agent": DEFAULT_UA}
        raw = self.headers.get("X-Upstream-Headers")
        if raw:
            try:
                for k, v in json.loads(raw).items():
                    if k and v:
                        headers[k] = v
            except Exception:
                pass
        headers.setdefault("User-Agent", DEFAULT_UA)  # GitHub API rejects requests with no UA

        req = urllib.request.Request(target, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT, context=SSL_CTX) as resp:
                body = resp.read()
                ctype = resp.headers.get("Content-Type", "application/octet-stream")
                return self._send(resp.status, body, ctype)
        except urllib.error.HTTPError as e:
            # Pass the upstream status + body straight through, so the page's
            # `response.ok` check behaves exactly as it did before.
            body = b""
            try:
                body = e.read()
            except Exception:
                pass
            ctype = "text/plain"
            if e.headers:
                ctype = e.headers.get("Content-Type", "text/plain")
            return self._send(e.code, body, ctype)
        except Exception as e:
            return self._send(502, "proxy error: " + str(e))


class ThreadingHTTPServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def pick_port(start=8000, tries=25):
    for p in range(start, start + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start


def main():
    if not os.path.exists(HTML_FILE):
        print("ERROR: project.html must be in the same folder as this script.")
        try:
            input("Press Enter to exit...")
        except EOFError:
            pass
        return

    port = pick_port()
    url = "http://localhost:%d/" % port
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)

    line = "=" * 58
    print(line)
    print("  HuggingFace Model BOM  -  local server running")
    print("  " + url)
    print("  Keep this window open. Close it to stop the tool.")
    print(line)

    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
