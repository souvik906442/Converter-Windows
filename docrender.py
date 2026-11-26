"""Offline DOCX / XLSX -> page images renderer (Python port of Android DocRender.java).

Needs only Pillow. Handles text, headings, bold/italic, lists, tables and embedded
pictures of a .docx, and every sheet of a .xlsx (as tables). It is a simple renderer,
not Microsoft Word: complex layout (text boxes, headers/footers, charts) is ignored.
"""
import io
import os
import re
import zipfile
import xml.parsers.expat
from decimal import Context, InvalidOperation

from PIL import Image, ImageDraw, ImageFont

W, H, M = 1190, 1684, 100
AVAIL = W - 2 * M
BASE = 22.0
HM = [1.0, 1.9, 1.6, 1.35, 1.15, 1.1]
MAX_ENTRY = 25 * 1024 * 1024


# ---------------- fonts ----------------

def _font_dirs():
    d = []
    win = os.environ.get("WINDIR") or os.environ.get("SystemRoot")
    if win:
        d.append(os.path.join(win, "Fonts"))
    d += ["/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/dejavu",
          "/usr/share/fonts/truetype/freefont", "/usr/share/fonts/truetype/noto",
          "/usr/share/fonts/opentype/noto", "/Library/Fonts", "/System/Library/Fonts/Supplemental"]
    return d


# (regular, bold, italic, bold-italic) candidates; first existing file wins
_LATIN = [
    ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
    ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
    ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
]
# Bengali / Hindi / other scripts that Segoe UI does not have
_COMPLEX = [
    ("nirmala.ttf", "nirmalab.ttf", "nirmala.ttf", "nirmalab.ttf"),
    ("Nirmala.ttc", "Nirmala.ttc", "Nirmala.ttc", "Nirmala.ttc"),
    ("NotoSansBengali-Regular.ttf", "NotoSansBengali-Bold.ttf", "NotoSansBengali-Regular.ttf", "NotoSansBengali-Bold.ttf"),
    ("FreeSans.ttf", "FreeSansBold.ttf", "FreeSansOblique.ttf", "FreeSansBoldOblique.ttf"),
    ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
]


def _find(name):
    for d in _font_dirs():
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
    return None


def _pick(cands):
    for fam in cands:
        paths = [_find(n) for n in fam]
        if paths[0]:
            return [p or paths[0] for p in paths]
    return None


_FILES = {}
_CACHE = {}


def _is_complex(s):
    for ch in s:
        o = ord(ch)
        if 0x0590 <= o <= 0x1FFF or 0x3000 <= o <= 0x9FFF:
            return True
    return False


def get_font(sz, bold, italic, cx):
    key = (round(sz, 1), bool(bold), bool(italic), bool(cx))
    f = _CACHE.get(key)
    if f is not None:
        return f
    kind = "c" if cx else "l"
    if kind not in _FILES:
        _FILES[kind] = _pick(_COMPLEX if cx else _LATIN)
    files = _FILES[kind]
    if files is None and cx:
        # no complex-script font found: fall back to the Latin one
        if "l" not in _FILES:
            _FILES["l"] = _pick(_LATIN)
        files = _FILES["l"]
    if files:
        try:
            f = ImageFont.truetype(files[(1 if bold else 0) + (2 if italic else 0)], size=max(4.0, key[0]))
        except Exception:
            f = None
    if f is None:
        try:
            f = ImageFont.load_default(size=max(4.0, key[0]))
        except TypeError:
            f = ImageFont.load_default()
    _CACHE[key] = f
    return f


# ---------------- layout engine ----------------

class Run:
    __slots__ = ("t", "b", "i", "sz")

    def __init__(self, t, b=False, i=False, sz=0.0):
        self.t, self.b, self.i, self.sz = t, b, i, sz


class Tok:
    __slots__ = ("s", "b", "i", "nl", "tab", "glue", "sz", "w", "gap", "cx")

    def __init__(self):
        self.s = ""
        self.b = self.i = self.nl = self.tab = self.glue = self.cx = False
        self.sz = BASE
        self.w = self.gap = 0.0


class Layout:
    def __init__(self, sink, dry):
        self.sink, self.dry = sink, dry
        self.cur = None
        self.dr = None
        self.y = 0.0
        self.started = False
        self.pages = 0

    def new_page(self):
        if not self.dry:
            self.cur = Image.new("RGB", (W, H), "white")
            self.dr = ImageDraw.Draw(self.cur)
        self.y = float(M)
        self.started = True

    def flush(self):
        if self.started:
            self.pages += 1
            self.started = False
            if not self.dry:
                b, self.cur, self.dr = self.cur, None, None
                self.sink(b)

    def ensure(self, h):
        if not self.started:
            self.new_page()
        elif self.y + h > H - M:
            self.flush()
            self.new_page()

    @staticmethod
    def font(t):
        return get_font(t.sz, t.b, t.i, t.cx)

    def measure(self, t):
        return self.font(t).getlength(t.s)

    def space_w(self, t):
        return self.font(t).getlength(" ")

    def add_word(self, toks, w, r, avail, glue):
        t = Tok()
        t.b, t.i = r.b, r.i
        t.sz = r.sz if r.sz > 0 else BASE
        t.s, t.glue, t.cx = w, glue, _is_complex(w)
        f = self.font(t)
        tw = f.getlength(w)
        if tw <= avail:
            t.w = tw
            toks.append(t)
            return
        rest, first = w, True
        while rest:
            lo, hi, n = 1, len(rest), 1
            while lo <= hi:                      # longest prefix that fits
                mid = (lo + hi) // 2
                if f.getlength(rest[:mid]) <= avail:
                    n, lo = mid, mid + 1
                else:
                    hi = mid - 1
            c = Tok()
            c.b, c.i, c.sz, c.cx = t.b, t.i, t.sz, t.cx
            c.s = rest[:n]
            c.glue = first and glue
            c.w = f.getlength(c.s)
            toks.append(c)
            rest, first = rest[n:], False

    def wrap(self, runs, avail):
        avail = max(avail, 20.0)
        toks = []
        prev_non_space = False
        for r in runs:
            if r.t is None:
                continue
            if r.t == "\n":
                t = Tok()
                t.nl = True
                t.sz = r.sz if r.sz > 0 else BASE
                toks.append(t)
                prev_non_space = False
                continue
            if r.t == "\t":
                t = Tok()
                t.tab, t.s, t.w = True, "", 60.0
                t.sz = r.sz if r.sz > 0 else BASE
                toks.append(t)
                prev_non_space = False
                continue
            text = r.t
            if not text:
                continue
            starts_space, ends_space = text[0].isspace(), text[-1].isspace()
            first_word = True
            for w in text.split():
                glue = first_word and prev_non_space and not starts_space
                self.add_word(toks, w, r, avail, glue)
                first_word = False
            prev_non_space = not ends_space
        lines, cur, cur_w = [], [], 0.0
        for t in toks:
            if t.nl:
                lines.append(cur)
                cur, cur_w = [], 0.0
                continue
            gap = 0.0 if (not cur or t.glue) else self.space_w(t)
            if cur and not t.glue and cur_w + gap + t.w > avail:
                lines.append(cur)
                cur, cur_w, gap = [], 0.0, 0.0
            t.gap = gap
            cur.append(t)
            cur_w += gap + t.w
        lines.append(cur)
        return lines

    @staticmethod
    def max_sz(line):
        m = max((t.sz for t in line), default=0.0)
        return m if m else BASE

    def line_height(self, line):
        return self.max_sz(line) * 1.4

    @staticmethod
    def line_width(line):
        return sum(t.gap + t.w for t in line)

    def draw_line(self, line, x0, y_top, width, align):
        if self.dry:
            return
        lw = self.line_width(line)
        x = x0 + (width - lw) / 2 if align == 1 else (x0 + width - lw if align == 2 else x0)
        m = self.max_sz(line)
        for t in line:
            x += t.gap
            if not t.tab and t.s:
                self.dr.text((x, y_top + m * 1.05), t.s, font=self.font(t), fill=(0, 0, 0), anchor="ls")
            x += t.w

    def paragraph(self, runs, align, after):
        for line in self.wrap(runs, AVAIL):
            h = self.line_height(line)
            self.ensure(h)
            self.draw_line(line, M, self.y, AVAIL, align)
            self.y += h
        self.y += after

    def gap(self, g):
        if self.started:
            self.y += g

    def image(self, data):
        try:
            im = Image.open(io.BytesIO(data))
            ow, oh = im.size
        except Exception:
            return
        if ow <= 0 or oh <= 0:
            return
        w = min(float(AVAIL), ow * 1.5)
        h = w * oh / ow
        max_h = H - 2 * M
        if h > max_h:
            h = float(max_h)
            w = h * ow / oh
        self.ensure(h + 10)
        if not self.dry:
            try:
                im.load()
                im = im.convert("RGBA").resize((max(1, round(w)), max(1, round(h))), Image.LANCZOS)
                self.cur.paste(im, (M, round(self.y)), im)
            except Exception:
                pass
        self.y += h + 10

    def table_row(self, cells, col_w, pad):
        n = len(cells)
        if n == 0:
            return
        cw = col_w if (col_w is not None and len(col_w) == n) else [AVAIL / n] * n
        allc, row_h = [], 0.0
        for i in range(n):
            ls = self.wrap(cells[i], cw[i] - 2 * pad)
            if len(ls) > 40:
                ls = ls[:40]
            allc.append(ls)
            row_h = max(row_h, 2 * pad + sum(self.line_height(l) for l in ls))
        self.ensure(row_h)
        if not self.dry:
            x = float(M)
            for i in range(n):
                self.dr.rectangle([x, self.y, x + cw[i], self.y + row_h], outline=(136, 136, 136), width=2)
                yy = self.y + pad
                for l in allc[i]:
                    self.draw_line(l, x + pad, yy, cw[i] - 2 * pad, 0)
                    yy += self.line_height(l)
                x += cw[i]
        self.y += row_h


# ---------------- XML helper ----------------

def _parse(data, start, end, text=None):
    p = xml.parsers.expat.ParserCreate()
    p.buffer_text = True
    p.StartElementHandler = start
    p.EndElementHandler = end
    if text:
        p.CharacterDataHandler = text
    p.Parse(data, True)


def _off(v):
    return v is not None and v in ("0", "false")


def parse_rels(data):
    m = {}
    if not data:
        return m

    def st(name, a):
        if name == "Relationship" and "Id" in a and "Target" in a:
            m[a["Id"]] = a["Target"]
    _parse(data, st, lambda n: None)
    return m


# ---------------- DOCX ----------------

def render_docx(z, L):
    doc = z.get("word/document.xml")
    if doc is None:
        raise IOError("এটা সঠিক DOCX ফাইল না")
    rels = parse_rels(z.get("word/_rels/document.xml.rels"))
    S = dict(runs=None, imgs=[], align=0, head=0, in_ppr=False, in_rpr=False, in_t=False,
             list_item=False, page_break=False, rb=False, ri=False, rsz=0.0, tbl=0,
             grid=[], row_cells=None, cell_runs=None)

    def start(n, a):
        if n == "w:p":
            S.update(runs=[], imgs=[], align=0, head=0, list_item=False, page_break=False)
        elif n == "w:pPr":
            S["in_ppr"] = True
        elif n == "w:rPr":
            S["in_rpr"] = True
        elif n == "w:pStyle":
            v = a.get("w:val")
            if v:
                s = v.lower()
                if "title" in s:
                    S["head"] = 1
                elif "heading1" in s or s == "1":
                    S["head"] = 2
                elif "heading2" in s or s == "2":
                    S["head"] = 3
                elif "heading3" in s or s == "3":
                    S["head"] = 4
                elif "heading" in s:
                    S["head"] = 5
        elif n == "w:jc":
            if S["in_ppr"]:
                v = a.get("w:val")
                S["align"] = 1 if v == "center" else (2 if v in ("right", "end") else 0)
        elif n == "w:numPr":
            if S["in_ppr"]:
                S["list_item"] = True
        elif n == "w:r":
            S.update(rb=False, ri=False, rsz=0.0)
        elif n == "w:b":
            if S["in_rpr"] and not S["in_ppr"]:
                S["rb"] = not _off(a.get("w:val"))
        elif n == "w:i":
            if S["in_rpr"] and not S["in_ppr"]:
                S["ri"] = not _off(a.get("w:val"))
        elif n == "w:sz":
            if S["in_rpr"] and not S["in_ppr"]:
                try:
                    S["rsz"] = float(a.get("w:val"))
                except (TypeError, ValueError):
                    pass
        elif n == "w:t":
            S["in_t"] = True
        elif n == "w:tab":
            if not S["in_ppr"] and S["runs"] is not None:
                S["runs"].append(Run("\t"))
        elif n == "w:br":
            if S["runs"] is not None:
                if a.get("w:type") == "page":
                    S["page_break"] = True
                else:
                    S["runs"].append(Run("\n"))
        elif n == "a:blip":
            tg = rels.get(a.get("r:embed"))
            if tg:
                path = tg[1:] if tg.startswith("/") else "word/" + tg
                d = z.get(os.path.normpath(path).replace("\\", "/"))
                if d is not None:
                    S["imgs"].append(d)
        elif n == "w:tbl":
            S["tbl"] += 1
            if S["tbl"] == 1:
                S["grid"] = []
        elif n == "w:gridCol":
            if S["tbl"] == 1:
                try:
                    S["grid"].append(float(a.get("w:w")))
                except (TypeError, ValueError):
                    pass
        elif n == "w:tr":
            if S["tbl"] == 1:
                S["row_cells"] = []
        elif n == "w:tc":
            if S["tbl"] == 1:
                S["cell_runs"] = []

    def text(s):
        if S["in_t"] and S["runs"] is not None:
            h = S["head"]
            sz = S["rsz"] if S["rsz"] > 0 else (BASE * HM[h] if h > 0 else BASE)
            S["runs"].append(Run(s, S["rb"] or h > 0, S["ri"], sz))

    def end(n):
        if n == "w:t":
            S["in_t"] = False
        elif n == "w:pPr":
            S["in_ppr"] = False
        elif n == "w:rPr":
            S["in_rpr"] = False
        elif n == "w:p":
            runs = S["runs"]
            if runs is not None:
                if S["list_item"]:
                    runs.insert(0, Run("•  "))
                if S["tbl"] > 0:
                    cr = S["cell_runs"]
                    if cr is not None:
                        if cr:
                            cr.append(Run("\n"))
                        cr.extend(runs)
                else:
                    L.paragraph(runs, S["align"], 14 if S["head"] > 0 else 10)
                    for img in S["imgs"]:
                        L.image(img)
                    if S["page_break"]:
                        L.flush()
            S["runs"] = None
        elif n == "w:tc":
            if S["tbl"] == 1:
                if S["row_cells"] is not None and S["cell_runs"] is not None:
                    S["row_cells"].append(S["cell_runs"])
                S["cell_runs"] = None
        elif n == "w:tr":
            if S["tbl"] == 1 and S["row_cells"] is not None:
                g, rc = S["grid"], S["row_cells"]
                cw = None
                if g and len(g) == len(rc):
                    tot = sum(g) or 1.0
                    cw = [x * AVAIL / tot for x in g]
                L.table_row(rc, cw, 8)
                S["row_cells"] = None
        elif n == "w:tbl":
            S["tbl"] -= 1
            if S["tbl"] == 0:
                L.gap(12)

    _parse(doc, start, end, text)


# ---------------- XLSX ----------------

def _local(n):
    return n.rsplit(":", 1)[-1]


def parse_shared(data):
    out = []
    if not data:
        return out
    S = dict(sb=None, in_t=False, in_ph=False)

    def start(n, a):
        n = _local(n)
        if n == "si":
            S["sb"] = []
        elif n == "rPh":
            S["in_ph"] = True
        elif n == "t":
            S["in_t"] = True

    def text(s):
        if S["in_t"] and not S["in_ph"] and S["sb"] is not None:
            S["sb"].append(s)

    def end(n):
        n = _local(n)
        if n == "t":
            S["in_t"] = False
        elif n == "rPh":
            S["in_ph"] = False
        elif n == "si":
            out.append("".join(S["sb"] or []))
            S["sb"] = None
    _parse(data, start, end, text)
    return out


def col_index(ref):
    n = 0
    for ch in ref:
        if "A" <= ch <= "Z":
            n = n * 26 + ord(ch) - 64
        elif "a" <= ch <= "z":
            n = n * 26 + ord(ch) - 96
        else:
            break
    return max(0, n - 1)


_CTX = Context(prec=12)


def num(raw):
    try:
        d = _CTX.create_decimal(raw.strip())
        if not d.is_finite():
            return raw
        return format(d.normalize(), "f")
    except (InvalidOperation, ValueError):
        return raw


def parse_sheet(data, ss):
    rows = []
    S = dict(cur=None, col=0, type=None, sb=[], in_v=False, in_t=False, in_is=False)

    def start(n, a):
        n = _local(n)
        if n == "row":
            S["cur"] = []
        elif n == "c":
            ref = a.get("r")
            S["col"] = col_index(ref) if ref else (len(S["cur"]) if S["cur"] is not None else 0)
            S["type"] = a.get("t")
            S["sb"] = []
        elif n == "v":
            S["in_v"] = True
        elif n == "is":
            S["in_is"] = True
        elif n == "t":
            if S["in_is"]:
                S["in_t"] = True

    def text(s):
        if S["in_v"] or S["in_t"]:
            S["sb"].append(s)

    def end(n):
        n = _local(n)
        if n == "v":
            S["in_v"] = False
        elif n == "t":
            S["in_t"] = False
        elif n == "is":
            S["in_is"] = False
        elif n == "c":
            cur = S["cur"]
            if cur is None:
                return
            val, tp = "".join(S["sb"]), S["type"]
            if tp == "s":
                try:
                    val = ss[int(val.strip())]
                except (ValueError, IndexError):
                    val = ""
            elif tp == "b":
                val = "TRUE" if val.strip() == "1" else "FALSE"
            elif tp is None or tp == "n":
                if val:
                    val = num(val)
            col = min(S["col"], 16384)
            while len(cur) < col:
                cur.append("")
            if len(cur) == col:
                cur.append(val)
            else:
                cur[col] = val
        elif n == "row":
            if S["cur"] is not None:
                rows.append(S["cur"])
            S["cur"] = None
    _parse(data, start, end, text)
    return rows


def render_xlsx(z, L):
    ss = parse_shared(z.get("xl/sharedStrings.xml"))
    names = []
    wb = z.get("xl/workbook.xml")
    if wb:
        def st(n, a):
            if _local(n) == "sheet":
                names.append(a.get("name") or "")
        _parse(wb, st, lambda n: None)
    pat = re.compile(r"xl/worksheets/sheet(\d+)\.xml")
    keys = sorted((k for k in z if pat.fullmatch(k)), key=lambda k: int(pat.fullmatch(k).group(1)))
    if not keys:
        raise IOError("এটা সঠিক XLSX ফাইল না")
    for idx, k in enumerate(keys):
        rows = parse_sheet(z[k], ss)
        name = names[idx] if idx < len(names) else "Sheet%d" % (idx + 1)
        draw_sheet(L, name, rows)


def draw_sheet(L, name, rows):
    L.flush()
    L.paragraph([Run(name, True, False, BASE * 1.4)], 0, 10)
    cols = min(max((len(r) for r in rows), default=0), 30)
    if cols == 0:
        L.paragraph([Run("(খালি শীট)", False, True, 0)], 0, 10)
        return
    nat = [0.0] * cols
    for r in rows:
        for c in range(min(cols, len(r))):
            if not r[c]:
                continue
            f = get_font(BASE, False, False, _is_complex(r[c]))
            w = f.getlength(r[c]) + 16
            if w > nat[c]:
                nat[c] = w
    nat = [max(60.0, min(360.0, x)) for x in nat]
    total = sum(nat)
    fs, cw = BASE, nat
    if total > AVAIL:
        scale = AVAIL / total
        cw = [max(24.0, x * scale) for x in nat]
        fs = max(12.0, BASE * max(0.55, scale))
    for ri, r in enumerate(rows):
        cells = []
        for c in range(cols):
            v = r[c] if c < len(r) and r[c] else ""
            cells.append([Run(v, ri == 0 and len(rows) > 1, False, fs)])
        L.table_row(cells, cw, 6)


# ---------------- public API ----------------

def read_zip(path):
    """Read only the parts we need from a .docx / .xlsx into a dict (each part capped at 25 MB)."""
    m = {}
    with zipfile.ZipFile(str(path)) as zf:
        for info in zf.infolist():
            n = info.filename
            want = (n in ("word/document.xml", "word/_rels/document.xml.rels", "xl/sharedStrings.xml",
                          "xl/workbook.xml")
                    or n.startswith("word/media/")
                    or (n.startswith("xl/worksheets/sheet") and n.endswith(".xml")))
            if info.is_dir() or not want:
                continue
            with zf.open(info) as f:
                data = f.read(MAX_ENTRY + 1)
            if len(data) <= MAX_ENTRY:
                m[n] = data
    return m


def render(z, xlsx, sink=None, dry=False):
    """Render parts dict `z`. sink(PIL.Image) is called per page (not needed when dry).
    Returns the number of pages."""
    L = Layout(sink, dry)
    if xlsx:
        render_xlsx(z, L)
    else:
        render_docx(z, L)
    L.flush()
    return L.pages


def render_file(path, sink=None, dry=False):
    p = str(path)
    return render(read_zip(p), p.lower().endswith((".xlsx", ".xlsm")), sink, dry)
