import io, os, re, struct, zlib, zipfile, threading, queue, time, tempfile, shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
from PIL import Image, ImageOps, ImageTk

Image.MAX_IMAGE_PIXELS = None

# ---------------- remove.bg ----------------
import json, uuid, urllib.request, urllib.error
try:
    from _embedded_key import EMBEDDED_KEY   # build-এর সময় GitHub secret থেকে বসে
except Exception:
    EMBEDDED_KEY = ""
CFG = Path(os.environ.get("APPDATA", str(Path.home()))) / "OfflineConverter" / "config.json"


def load_key():
    k = (EMBEDDED_KEY or "").strip() or os.environ.get("REMOVEBG_KEY", "").strip()
    if k:
        return k
    try:
        return json.loads(CFG.read_text(encoding="utf-8")).get("removebg_key", "").strip()
    except Exception:
        return ""


def save_key(k):
    CFG.parent.mkdir(parents=True, exist_ok=True)
    CFG.write_text(json.dumps({"removebg_key": k.strip()}), encoding="utf-8")


def remove_bg(im, key):
    buf = io.BytesIO()
    im.convert("RGBA").save(buf, "PNG")
    b = "----" + uuid.uuid4().hex
    body = (f"--{b}\r\nContent-Disposition: form-data; name=\"size\"\r\n\r\nauto\r\n"
            f"--{b}\r\nContent-Disposition: form-data; name=\"image_file\"; filename=\"a.png\"\r\n"
            f"Content-Type: image/png\r\n\r\n").encode() + buf.getvalue() + f"\r\n--{b}--\r\n".encode()
    req = urllib.request.Request("https://api.remove.bg/v1.0/removebg", data=body,
                                 headers={"X-Api-Key": key, "Content-Type": f"multipart/form-data; boundary={b}"})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            out = Image.open(io.BytesIO(r.read()))
            out.load()
            return out
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"remove.bg error {e.code}: {e.read().decode('utf-8', 'ignore')[:200]}")

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
except Exception:
    pass
try:
    import pymupdf as fitz
except Exception:
    import fitz
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAVE_DND = True
except Exception:
    HAVE_DND = False

JPG, PNG, WEBP, PDF, DOCX = 0, 1, 2, 3, 4
FORMATS = ["JPG", "PNG", "WebP", "PDF (সব মিলিয়ে একটা)", "DOCX (সব মিলিয়ে একটা)"]
MODES = ["ছোট করো (এই সাইজের মধ্যে)", "বড় করো (কমপক্ষে এই সাইজ)", "ঠিক এই সাইজ (ছোট/বড় দুটোই)"]
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif", ".bmp", ".tif", ".tiff", ".gif"}
WORD_EXT = {".doc", ".docx", ".rtf", ".odt"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm", ".csv"}
PPT_EXT = {".ppt", ".pptx"}
OFFICE_EXT = WORD_EXT | EXCEL_EXT | PPT_EXT
ALL_EXT = IMG_EXT | OFFICE_EXT | {".pdf"}


# ---------------- image encoding ----------------

def flatten(img):
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        return bg
    return img.convert("RGB")


def prep(img, fmt):
    if fmt == JPG:
        return flatten(img)
    has_alpha = img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info)
    return img.convert("RGBA" if has_alpha else "RGB")


def comp(img, fmt, q):
    buf = io.BytesIO()
    if fmt == JPG:
        img.save(buf, "JPEG", quality=q)
    elif fmt == PNG:
        img.save(buf, "PNG", optimize=True)
    else:
        img.save(buf, "WEBP", quality=q)
    return buf.getvalue()


def encode_fit(img, fmt, target):
    cur = img
    for _ in range(14):
        if fmt == PNG:
            b = comp(cur, PNG, 100)
            if target <= 0 or len(b) <= target:
                return b
        else:
            b = comp(cur, fmt, 92 if target > 0 else 90)
            if target <= 0 or len(b) <= target:
                return b
            lo, hi, best = 5, 92, None
            while lo <= hi:
                mid = (lo + hi) // 2
                t = comp(cur, fmt, mid)
                if len(t) <= target:
                    best = t
                    lo = mid + 1
                else:
                    hi = mid - 1
            if best is not None:
                return best
        nw, nh = int(cur.width * 0.8), int(cur.height * 0.8)
        if nw < 16 or nh < 16:
            break
        cur = cur.resize((nw, nh), Image.LANCZOS)
    return comp(cur, fmt, 100 if fmt == PNG else 5)


def pad_to(b, fmt, target):
    """Grow file to (at least) target bytes with harmless extra data."""
    if target <= 0 or len(b) >= target:
        return b
    pad = target - len(b)
    if fmt == JPG:
        return b + b"\0" * pad
    if fmt == PNG:
        if pad < 16 or len(b) < 12 or b[-8:-4] != b"IEND":
            return b
        n = pad - 12
        data = b"pad\0" + b"\0" * (n - 4)
        chunk = struct.pack(">I", n) + b"tEXt" + data + struct.pack(">I", zlib.crc32(b"tEXt" + data) & 0xFFFFFFFF)
        return b[:-12] + chunk + b[-12:]
    if fmt == WEBP:
        if pad < 10 or len(b) < 12 or b[:4] != b"RIFF" or b[8:12] != b"WEBP":
            return b
        s = pad - 8
        if s & 1:
            s += 1
        riff = struct.unpack("<I", b[4:8])[0] + 8 + s
        return b[:4] + struct.pack("<I", riff) + b[8:] + b"JUNK" + struct.pack("<I", s) + b"\0" * s
    return b


# ---------------- PDF / DOCX builders ----------------

def build_pdf(jpgs, dims, pad=0):
    n = len(jpgs)
    out = bytearray()
    off = [0] * (3 + 3 * n)
    out += b"%PDF-1.4\n"
    off[1] = len(out)
    out += b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
    off[2] = len(out)
    kids = " ".join(f"{3 + 3 * i} 0 R" for i in range(n))
    out += f"2 0 obj\n<< /Type /Pages /Kids [{kids}] /Count {n} >>\nendobj\n".encode()
    for i, (w, h) in enumerate(dims):
        pw = 595.0
        ph = 595.0 * h / w
        pg, ct, im = 3 + 3 * i, 4 + 3 * i, 5 + 3 * i
        off[pg] = len(out)
        out += (f"{pg} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {pw:.2f} {ph:.2f}] "
                f"/Resources << /XObject << /Im0 {im} 0 R >> >> /Contents {ct} 0 R >>\nendobj\n").encode()
        cs = f"q {pw:.2f} 0 0 {ph:.2f} 0 0 cm /Im0 Do Q"
        off[ct] = len(out)
        out += f"{ct} 0 obj\n<< /Length {len(cs)} >>\nstream\n{cs}\nendstream\nendobj\n".encode()
        off[im] = len(out)
        j = jpgs[i]
        out += (f"{im} 0 obj\n<< /Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace /DeviceRGB "
                f"/BitsPerComponent 8 /Filter /DCTDecode /Length {len(j)} >>\nstream\n").encode()
        out += j
        out += b"\nendstream\nendobj\n"
    if pad >= 2:
        out += b"%" + b"P" * (pad - 2) + b"\n"
    xref = len(out)
    cnt = 3 + 3 * n
    out += f"xref\n0 {cnt}\n0000000000 65535 f \n".encode()
    for i in range(1, cnt):
        out += f"{off[i]:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {cnt} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(out)


def build_docx(jpgs, dims, pad=0):
    hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    bio = io.BytesIO()

    def zi(name, method):
        z = zipfile.ZipInfo(name, date_time=(2020, 1, 1, 0, 0, 0))
        z.compress_type = method
        return z

    with zipfile.ZipFile(bio, "w") as z:
        D, S = zipfile.ZIP_DEFLATED, zipfile.ZIP_STORED
        z.writestr(zi("[Content_Types].xml", D), hdr +
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Default Extension="bin" ContentType="application/octet-stream"/>'
                   '<Default Extension="jpg" ContentType="image/jpeg"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr(zi("_rels/.rels", D), hdr +
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '</Relationships>')
        rels = [hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        body = []
        max_w, max_h = 10466 * 635, 15398 * 635
        for i, (w, h) in enumerate(dims):
            k = i + 1
            rels.append(f'<Relationship Id="rId{k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image{k}.jpg"/>')
            cx = max_w
            cy = cx * h // w
            if cy > max_h:
                cy = max_h
                cx = cy * w // h
            body.append(
                f'<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{k}" name="Picture {k}"/>'
                f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f'<pic:nvPicPr><pic:cNvPr id="{k}" name="img{k}"/><pic:cNvPicPr/></pic:nvPicPr>'
                f'<pic:blipFill><a:blip r:embed="rId{k}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
        rels.append("</Relationships>")
        z.writestr(zi("word/_rels/document.xml.rels", D), "".join(rels))
        z.writestr(zi("word/document.xml", D), hdr +
                   '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                   'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>'
                   + "".join(body) +
                   '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
                   '</w:body></w:document>')
        for i, j in enumerate(jpgs):
            z.writestr(zi(f"word/media/image{i + 1}.jpg", S), j)
        if pad > 0:
            z.writestr(zi("word/pad.bin", S), b"\0" * pad)
    return bio.getvalue()


# ---------------- Office (needs Microsoft Office installed) ----------------

def office_to_pdf(p, tmp):
    try:
        import win32com.client as wc
    except Exception:
        raise RuntimeError("pywin32 পাওয়া যায়নি")
    out = str(Path(tmp) / f"{p.stem}_{abs(hash(str(p))) % 100000}.pdf")
    src = str(p.resolve())
    ext = p.suffix.lower()
    app = None
    try:
        if ext in WORD_EXT:
            app = wc.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            d = app.Documents.Open(src, False, True)
            d.ExportAsFixedFormat(out, 17)
            d.Close(False)
        elif ext in EXCEL_EXT:
            app = wc.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(src, 0, True)
            wb.ExportAsFixedFormat(0, out)
            wb.Close(False)
        else:
            app = wc.DispatchEx("PowerPoint.Application")
            pres = app.Presentations.Open(src, True, False, False)
            pres.SaveAs(out, 32)
            pres.Close()
    except Exception as e:
        raise RuntimeError(f"Microsoft Office দিয়ে খোলা যায়নি ({e})")
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
    if not os.path.exists(out):
        raise RuntimeError("Office PDF বানাতে পারেনি")
    return Path(out)


# ---------------- main conversion ----------------

def unique_path(outdir, name, ext):
    p = outdir / f"{name}.{ext}"
    i = 1
    while p.exists():
        p = outdir / f"{name}_{i}.{ext}"
        i += 1
    return p


def pdf_pages(path):
    with fitz.open(str(path)) as d:
        for pg in d:
            r = pg.rect
            z = min(2.0, 3000.0 / max(r.width, r.height, 1))
            pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            yield Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


_TR = getattr(Image, "Transpose", Image)
# edit ops: 0 = rotate 90 clockwise, 1 = flip left-right, 2 = flip top-bottom
OP_FUNCS = {0: _TR.ROTATE_270, 1: _TR.FLIP_LEFT_RIGHT, 2: _TR.FLIP_TOP_BOTTOM}


def resize_px(im, tw, th, keep):
    w, h = im.size
    if tw > 0 and th > 0:
        if keep:
            f = min(tw / w, th / h)
            nw, nh = round(w * f), round(h * f)
        else:
            nw, nh = tw, th
    elif tw > 0:
        nw, nh = tw, (round(h * tw / w) if keep else h)
    else:
        nh, nw = th, (round(w * th / h) if keep else w)
    nw, nh = max(1, nw), max(1, nh)
    if (nw, nh) == (w, h):
        return im
    return im.resize((nw, nh), Image.LANCZOS)


def img_size(path):
    """Pixel size (w, h) with EXIF rotation applied, or None."""
    try:
        with Image.open(str(path)) as im:
            w, h = im.size
            try:
                o = im.getexif().get(274)
            except Exception:
                o = None
            if o in (5, 6, 7, 8):
                w, h = h, w
            return w, h
    except Exception:
        return None


def image_pages(path, opts=None):
    opts = opts or {}
    im = Image.open(str(path))
    im = ImageOps.exif_transpose(im)
    im.load()
    if opts.get("rmbg") and opts.get("key"):
        im = remove_bg(im, opts["key"])
    for op in opts.get("ops", []):
        im = im.transpose(OP_FUNCS[op])
    cr = opts.get("crop")
    if cr:
        W, H = im.size
        x = max(0, min(int(cr[0]), W - 1))
        y = max(0, min(int(cr[1]), H - 1))
        w = max(1, min(int(cr[2]), W - x))
        h = max(1, min(int(cr[3]), H - y))
        im = im.crop((x, y, x + w, y + h))
    tw, th = opts.get("out_w", 0), opts.get("out_h", 0)
    if tw > 0 or th > 0:
        im = resize_px(im, tw, th, opts.get("keep", True))
    elif max(im.size) > 6000:
        im.thumbnail((6000, 6000))
    yield im


def run(files, fmt, target, mode, outdir, say, prog, opts=None):
    outdir.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="conv_"))
    errors = []
    try:
        sources = []
        for f in files:
            p = Path(f)
            ext = p.suffix.lower()
            base = re.sub(r'[<>:"/\\|?*]', "_", p.stem)
            try:
                if ext == ".pdf":
                    sources.append((base, "pdf", p))
                elif ext in OFFICE_EXT:
                    say(f"Office দিয়ে PDF করা হচ্ছে: {p.name}")
                    sources.append((base, "pdf", office_to_pdf(p, tmp)))
                else:
                    sources.append((base, "img", p))
            except Exception as e:
                errors.append(f"{p.name}: {e}")
        counts = []
        for base, kind, p in sources:
            if kind == "pdf":
                try:
                    with fitz.open(str(p)) as d:
                        counts.append(d.page_count)
                except Exception as e:
                    counts.append(0)
                    errors.append(f"{p.name}: {e}")
            else:
                counts.append(1)
        total = max(1, sum(counts))
        prog(0, total)
        bundle = fmt in (PDF, DOCX)
        if bundle:
            per = 0 if (target <= 0 or mode == 1) else max(2000, (target - (3500 if fmt == DOCX else 1500)) // total)
        else:
            per = target
        jpgs, dims = [], []
        saved = 0
        done = 0
        for (base, kind, p), n in zip(sources, counts):
            if n == 0:
                continue
            try:
                gen = pdf_pages(p) if kind == "pdf" else image_pages(p, opts)
                for i, img in enumerate(gen):
                    say(f"প্রসেস হচ্ছে: {p.name}" + (f" পাতা {i + 1}/{n}" if n > 1 else ""))
                    name = base + (f"_p{i + 1}" if n > 1 else "")
                    if bundle:
                        im = flatten(img)
                        j = encode_fit(im, JPG, per)
                        w, h = Image.open(io.BytesIO(j)).size
                        jpgs.append(j)
                        dims.append((w, h))
                    else:
                        im = prep(img, fmt)
                        out = encode_fit(im, fmt, 0 if mode == 1 else target)
                        if mode > 0:
                            out = pad_to(out, fmt, target)
                        ext = {JPG: "jpg", PNG: "png", WEBP: "webp"}[fmt]
                        unique_path(outdir, name, ext).write_bytes(out)
                        saved += 1
                    done += 1
                    prog(done, total)
            except Exception as e:
                errors.append(f"{p.name}: {e}")
        msg = ""
        if bundle:
            if not jpgs:
                return "কিছু কনভার্ট হয়নি" + ("\n" + "\n".join(errors) if errors else "")
            build = build_pdf if fmt == PDF else build_docx
            out = build(jpgs, dims, 0)
            if mode > 0 and target > 0 and len(out) < target:
                pad = target - len(out)
                for _ in range(5):
                    if pad <= 0:
                        break
                    out = build(jpgs, dims, pad)
                    diff = target - len(out)
                    if diff == 0:
                        break
                    pad += diff
            ext = "pdf" if fmt == PDF else "docx"
            path = unique_path(outdir, "converted_" + time.strftime("%Y%m%d_%H%M%S"), ext)
            path.write_bytes(out)
            msg = f"শেষ! {path.name} ({len(out) // 1024} KB)"
        else:
            msg = f"শেষ! {saved}টি ফাইল সেভ হয়েছে"
        if errors:
            msg += "\nসমস্যা:\n" + "\n".join(errors[:8])
        return msg
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------- colour theme ----------------

BG = "#1E1752"
FIELD = "#3A3280"
FG = "#FFFFFF"
ACCENT = "#FFD166"
BORDER = "#7C6BE0"
FONT = "Nirmala UI"


def apply_theme(root):
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except Exception:
        pass
    root.configure(bg=BG)
    st.configure(".", background=BG, foreground=FG, fieldbackground=FIELD, bordercolor=BORDER,
                 lightcolor=BORDER, darkcolor=BORDER, troughcolor=FIELD, focuscolor=BG)
    st.configure("TFrame", background=BG)
    st.configure("TLabel", background=BG, foreground=FG)
    st.configure("TLabelframe", background=BG, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, relief="solid")
    st.configure("TLabelframe.Label", background=BG, foreground=ACCENT, font=(FONT, 10, "bold"))
    st.configure("TEntry", fieldbackground=FIELD, foreground=FG, insertcolor=FG, bordercolor=BORDER, padding=4)
    st.configure("TCombobox", fieldbackground=FIELD, background=FIELD, foreground=FG, arrowcolor=FG,
                 bordercolor=BORDER, padding=4, selectbackground=FIELD, selectforeground=FG)
    st.map("TCombobox", fieldbackground=[("readonly", FIELD)], foreground=[("readonly", FG)],
           selectbackground=[("readonly", FIELD)], selectforeground=[("readonly", FG)])
    st.configure("TCheckbutton", background=BG, foreground=FG, indicatorcolor=FIELD)
    st.map("TCheckbutton", background=[("active", BG)], indicatorcolor=[("selected", ACCENT)])
    st.configure("TScrollbar", background=FIELD, troughcolor=BG, arrowcolor=FG, bordercolor=BG)
    st.configure("Horizontal.TProgressbar", troughcolor=FIELD, background=ACCENT, bordercolor=BORDER,
                 lightcolor=ACCENT, darkcolor=ACCENT)

    def btn(name, c, hover, font=None):
        st.configure(name, background=c, foreground="white", bordercolor=c, lightcolor=c, darkcolor=c,
                     focuscolor=c, padding=(12, 6), relief="flat", **({"font": font} if font else {}))
        st.map(name, background=[("disabled", "#555070"), ("active", hover)],
               foreground=[("disabled", "#B0AACF")])

    btn("TButton", "#4B41A0", "#5E53C4")
    btn("Blue.TButton", "#0072FF", "#2E8BFF")
    btn("Purple.TButton", "#8E2DE2", "#A24CF0")
    btn("Orange.TButton", "#F0701A", "#FF8A3D")
    btn("Green.TButton", "#11998E", "#1BB8AB")
    btn("Big.TButton", "#F0701A", "#FF8A3D", (FONT, 12, "bold"))
    root.option_add("*TCombobox*Listbox.background", FIELD)
    root.option_add("*TCombobox*Listbox.foreground", FG)
    root.option_add("*TCombobox*Listbox.selectBackground", "#8E2DE2")
    root.option_add("*TCombobox*Listbox.selectForeground", "white")


def _hex(c):
    return tuple(int(c[i:i + 2], 16) for i in (1, 3, 5))


def draw_gradient(cv, w, h, stops=("#FF4D8D", "#8E2DE2", "#0072FF")):
    cv.delete("grad")
    a = [_hex(c) for c in stops]
    for x in range(0, w, 2):
        t = x / max(1, w - 1) * (len(a) - 1)
        i = min(int(t), len(a) - 2)
        f = t - i
        col = "#%02x%02x%02x" % tuple(round(a[i][k] + (a[i + 1][k] - a[i][k]) * f) for k in range(3))
        cv.create_rectangle(x, 0, x + 2, h, fill=col, outline=col, tags="grad")
    cv.tag_lower("grad")


# ---------------- crop dialog ----------------

def _clamp(v, lo, hi):
    return lo if v < lo else (hi if v > hi else v)


class CropDialog:
    CW, CH = 620, 420

    def __init__(self, app, path):
        self.app = app
        im = Image.open(str(path))
        im = ImageOps.exif_transpose(im)
        im.load()
        self.iw, self.ih = im.size
        pv = im.convert("RGB")
        pv.thumbnail((1400, 1400))
        self.cur = pv
        self.cx = self.cy = 0.0
        self.cw, self.ch = float(self.iw), float(self.ih)
        self.aspect = 0.0
        self.dops = list(app.ops)
        for op in self.dops:
            self._apply(op)
        if app.crop:
            self.set_crop(*app.crop)

        self.top = tk.Toplevel(app.root)
        self.top.title("ক্রপ")
        self.top.configure(bg=BG)
        self.top.transient(app.root)
        self.top.resizable(False, False)
        self.info = ttk.Label(self.top, text="")
        self.info.pack(anchor="w", padx=10, pady=(8, 2))
        self.cv = tk.Canvas(self.top, width=self.CW, height=self.CH, bg="#222222", highlightthickness=0)
        self.cv.pack(padx=10)
        self.cv.bind("<ButtonPress-1>", self.on_press)
        self.cv.bind("<B1-Motion>", self.on_move)
        self.cv.bind("<ButtonRelease-1>", self.on_release)

        tb = ttk.Frame(self.top)
        tb.pack(fill="x", padx=10, pady=6)
        ttk.Button(tb, text="Rotate 90°", style="Blue.TButton", command=lambda: self.do_op(0)).pack(side="left", padx=2)
        ttk.Button(tb, text="Flip (বাম-ডান)", style="Purple.TButton", command=lambda: self.do_op(1)).pack(side="left", padx=2)
        ttk.Button(tb, text="Vertical (উপর-নিচ)", style="Orange.TButton", command=lambda: self.do_op(2)).pack(side="left", padx=2)

        gf = ttk.Frame(self.top)
        gf.pack(fill="x", padx=10)
        self.vars = {k: tk.StringVar() for k in ("w", "h", "x", "y")}
        self.ents = {}
        for col, (key, label) in enumerate((("w", "Width"), ("h", "Height"), ("x", "Position (X)"), ("y", "Position (Y)"))):
            f = ttk.Frame(gf)
            f.grid(row=0, column=col, padx=4, sticky="w")
            ttk.Label(f, text=label).pack(anchor="w")
            e = ttk.Entry(f, textvariable=self.vars[key], width=10)
            e.pack()
            e.bind("<KeyRelease>", lambda ev, k=key: self.on_edit(k))
            e.bind("<FocusOut>", lambda ev: self.sync(True))
            self.ents[key] = e
        af = ttk.Frame(self.top)
        af.pack(fill="x", padx=10, pady=6)
        ttk.Label(af, text="Aspect Ratio").pack(side="left")
        self.names = ["FreeForm", "Original", "1:1", "4:3", "3:4", "16:9", "9:16", "3:2", "2:3"]
        self.asp = ttk.Combobox(af, values=self.names, state="readonly", width=14)
        self.asp.current(0)
        self.asp.pack(side="left", padx=8)
        self.asp.bind("<<ComboboxSelected>>", lambda ev: self.on_aspect())

        bf = ttk.Frame(self.top)
        bf.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Button(bf, text="ঠিক আছে", style="Orange.TButton", command=self.ok).pack(side="right", padx=2)
        ttk.Button(bf, text="বাতিল", command=self.top.destroy).pack(side="right", padx=2)
        ttk.Button(bf, text="ক্রপ বাদ", command=self.clear).pack(side="left", padx=2)

        self.mode = 0
        self.photo = None
        self.photo_key = None
        self.update_info()
        self.redraw()
        self.sync(True)
        self.top.grab_set()
        self.top.focus_set()

    # ---- geometry / state
    def ratio(self, idx):
        r = [0.0, self.iw / self.ih, 1.0, 4 / 3, 3 / 4, 16 / 9, 9 / 16, 3 / 2, 2 / 3]
        return r[idx]

    def update_info(self):
        t = f"ছবি: {self.iw} × {self.ih} px"
        if self.dops:
            t += f" (ঘোরানো/ফ্লিপ: {len(self.dops)} ধাপ)"
        self.info.config(text=t + " — ছবির উপর টেনে সরাও, কোণ ধরে মাপ বদলাও")

    def _apply(self, op):
        if op == 0:
            self.cur = self.cur.transpose(OP_FUNCS[0])
            ox, oy, ow, oh = self.cx, self.cy, self.cw, self.ch
            oiw, oih = self.iw, self.ih
            self.iw, self.ih = oih, oiw
            self.cw, self.ch = oh, ow
            self.cx = _clamp(oih - (oy + oh), 0, self.iw - self.cw)
            self.cy = _clamp(ox, 0, self.ih - self.ch)
        elif op == 1:
            self.cur = self.cur.transpose(OP_FUNCS[1])
            self.cx = _clamp(self.iw - (self.cx + self.cw), 0, self.iw - self.cw)
        elif op == 2:
            self.cur = self.cur.transpose(OP_FUNCS[2])
            self.cy = _clamp(self.ih - (self.cy + self.ch), 0, self.ih - self.ch)

    def do_op(self, op):
        self._apply(op)
        self.dops.append(op)
        if op == 0:
            idx = self.asp.current()
            if idx > 0:
                self.set_aspect(self.ratio(idx))
        self.update_info()
        self.changed()

    def changed(self):
        self.redraw()
        self.sync(False)

    def set_crop(self, x, y, w, h):
        w = _clamp(w, 1, self.iw)
        h = _clamp(h, 1, self.ih)
        self.cw, self.ch = float(w), float(h)
        self.cx = _clamp(x, 0, self.iw - self.cw)
        self.cy = _clamp(y, 0, self.ih - self.ch)

    def set_aspect(self, r):
        self.aspect = r
        if r > 0:
            w = self.cw
            h = w / r
            if h > self.ih:
                h = self.ih
                w = h * r
            if w > self.iw:
                w = self.iw
                h = w / r
            ncx = self.cx + self.cw / 2 - w / 2
            ncy = self.cy + self.ch / 2 - h / 2
            self.set_crop(ncx, ncy, w, h)

    def on_aspect(self):
        self.set_aspect(self.ratio(self.asp.current()))
        self.changed()

    def on_edit(self, key):
        try:
            v = int(self.vars[key].get().strip())
        except ValueError:
            return
        a = self.aspect
        if key == "w":
            w = _clamp(v, 1, self.iw)
            h = self.ch
            if a > 0:
                h = w / a
                if h > self.ih:
                    h = self.ih
                    w = h * a
            self.set_crop(self.cx, self.cy, w, h)
        elif key == "h":
            h = _clamp(v, 1, self.ih)
            w = self.cw
            if a > 0:
                w = h * a
                if w > self.iw:
                    w = self.iw
                    h = w / a
            self.set_crop(self.cx, self.cy, w, h)
        elif key == "x":
            self.set_crop(v, self.cy, self.cw, self.ch)
        else:
            self.set_crop(self.cx, v, self.cw, self.ch)
        self.changed()

    def get_crop(self):
        return (round(self.cx), round(self.cy), round(self.cw), round(self.ch))

    def sync(self, force):
        c = self.get_crop()
        vals = {"w": c[2], "h": c[3], "x": c[0], "y": c[1]}
        focus = self.top.focus_get() if hasattr(self, "top") else None
        for k, v in vals.items():
            if force or focus is not self.ents[k]:
                self.vars[k].set(str(v))

    # ---- drawing
    def redraw(self):
        if not hasattr(self, "cv"):
            return
        s = min(self.CW / self.iw, self.CH / self.ih)
        dw, dh = max(1, int(self.iw * s)), max(1, int(self.ih * s))
        ox, oy = (self.CW - dw) / 2, (self.CH - dh) / 2
        self.s, self.ox, self.oy = s, ox, oy
        key = (id(self.cur), dw, dh)
        if key != self.photo_key:
            self.photo = ImageTk.PhotoImage(self.cur.resize((dw, dh), Image.BILINEAR))
            self.photo_key = key
        c = self.cv
        c.delete("all")
        c.create_image(ox, oy, anchor="nw", image=self.photo)
        l, t = ox + self.cx * s, oy + self.cy * s
        r, b = l + self.cw * s, t + self.ch * s
        for box in ((ox, oy, ox + dw, t), (ox, b, ox + dw, oy + dh), (ox, t, l, b), (r, t, ox + dw, b)):
            c.create_rectangle(*box, fill="black", stipple="gray50", outline="")
        w3, h3 = (r - l) / 3, (b - t) / 3
        for i in (1, 2):
            c.create_line(l + w3 * i, t, l + w3 * i, b, fill="#dddddd")
            c.create_line(l, t + h3 * i, r, t + h3 * i, fill="#dddddd")
        c.create_rectangle(l, t, r, b, outline="white", width=2)
        for px, py in ((l, t), (r, t), (l, b), (r, b)):
            c.create_oval(px - 7, py - 7, px + 7, py + 7, fill="white", outline="#888888")

    # ---- mouse
    def on_press(self, e):
        s, ox, oy = self.s, self.ox, self.oy
        l, t = ox + self.cx * s, oy + self.cy * s
        r, b = l + self.cw * s, t + self.ch * s
        self.mode = 0
        best = 20 * 20
        for i, (px, py) in enumerate(((l, t), (r, t), (l, b), (r, b))):
            d = (e.x - px) ** 2 + (e.y - py) ** 2
            if d < best:
                best, self.mode = d, 2 + i
        if self.mode == 0 and l <= e.x <= r and t <= e.y <= b:
            self.mode = 1
        if self.mode >= 2:
            self.ax = self.cx + self.cw if self.mode in (2, 4) else self.cx
            self.ay = self.cy + self.ch if self.mode in (2, 3) else self.cy
        self.lx, self.ly = e.x, e.y

    def on_move(self, e):
        s = self.s
        if self.mode == 1:
            self.cx = _clamp(self.cx + (e.x - self.lx) / s, 0, self.iw - self.cw)
            self.cy = _clamp(self.cy + (e.y - self.ly) / s, 0, self.ih - self.ch)
            self.lx, self.ly = e.x, e.y
            self.changed()
        elif self.mode >= 2:
            px = _clamp((e.x - self.ox) / s, 0, self.iw)
            py = _clamp((e.y - self.oy) / s, 0, self.ih)
            nw = max(8.0, abs(px - self.ax))
            nh = max(8.0, abs(py - self.ay))
            if self.aspect > 0:
                if nw / nh > self.aspect:
                    nw = nh * self.aspect
                else:
                    nh = nw / self.aspect
            nx = self.ax if px >= self.ax else self.ax - nw
            ny = self.ay if py >= self.ay else self.ay - nh
            self.cw, self.ch = min(nw, self.iw), min(nh, self.ih)
            self.cx = _clamp(nx, 0, self.iw - self.cw)
            self.cy = _clamp(ny, 0, self.ih - self.ch)
            self.changed()

    def on_release(self, e):
        self.mode = 0

    # ---- buttons
    def ok(self):
        c = self.get_crop()
        full = c[0] == 0 and c[1] == 0 and c[2] >= self.iw and c[3] >= self.ih
        self.app.crop = None if full else c
        self.app.ops = list(self.dops)
        self.app.update_crop_lbl()
        self.top.destroy()

    def clear(self):
        self.app.crop = None
        self.app.ops = []
        self.app.update_crop_lbl()
        self.top.destroy()


# ---------------- GUI ----------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("Offline Converter")
        sh = root.winfo_screenheight()
        root.geometry(f"740x{min(900, max(560, sh - 90))}")   # টাস্কবারের নিচে ঢাকা পড়বে না
        root.minsize(700, 520)
        try:
            for nm in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
                tkfont.nametofont(nm).configure(family="Nirmala UI", size=10)
        except Exception:
            pass
        apply_theme(root)
        self.files = []
        self.crop = None
        self.ops = []
        self.q = queue.Queue()
        self.outdir = tk.StringVar(value=str(Path.home() / "Downloads" / "Converter"))
        self.size = tk.StringVar()
        P = dict(padx=10, pady=4)

        hd = tk.Canvas(root, height=58, highlightthickness=0, bg=BG)
        hd.pack(fill="x")
        hd.create_text(18, 29, anchor="w", text="Offline Converter", fill="white", font=(FONT, 18, "bold"), tags="title")
        hd.bind("<Configure>", lambda ev: (draw_gradient(hd, ev.width, 58), hd.tag_raise("title")))

        ttk.Label(root, text="ফাইল বাছাই করো (ছবি / PDF / Word / Excel / PowerPoint) — ড্র্যাগ করেও ফেলতে পারো").pack(anchor="w", **P)
        fr = ttk.Frame(root)
        fr.pack(fill="both", expand=True, **P)
        self.lb = tk.Listbox(fr, height=8, selectmode="extended", bg=FIELD, fg=FG, selectbackground="#8E2DE2",
                             selectforeground="white", highlightthickness=1, highlightbackground=BORDER, relief="flat", bd=0)
        sb = ttk.Scrollbar(fr, command=self.lb.yview)
        self.lb.config(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bf = ttk.Frame(root)
        bf.pack(fill="x", **P)
        ttk.Button(bf, text="ফাইল যোগ করো", style="Blue.TButton", command=self.add_dialog).pack(side="left", padx=2)
        ttk.Button(bf, text="ফোল্ডার যোগ করো", style="Blue.TButton", command=self.add_folder).pack(side="left", padx=2)
        ttk.Button(bf, text="বাছাইকৃত মুছো", command=self.remove_sel).pack(side="left", padx=2)
        ttk.Button(bf, text="সব মুছো", command=self.clear).pack(side="left", padx=2)
        self.count = ttk.Label(bf, text="০টি ফাইল")
        self.count.pack(side="right")

        ttk.Label(root, text="আউটপুট ফরম্যাট:").pack(anchor="w", **P)
        self.fmt = ttk.Combobox(root, values=FORMATS, state="readonly")
        self.fmt.current(0)
        self.fmt.pack(fill="x", **P)

        ttk.Label(root, text="সাইজ লিমিট (খালি রাখলে লিমিট নেই):").pack(anchor="w", **P)
        sf = ttk.Frame(root)
        sf.pack(fill="x", **P)
        ttk.Entry(sf, textvariable=self.size, width=14).pack(side="left")
        self.unit = ttk.Combobox(sf, values=["KB", "MB"], state="readonly", width=6)
        self.unit.current(0)
        self.unit.pack(side="left", padx=6)

        ttk.Label(root, text="সাইজ মোড:").pack(anchor="w", **P)
        self.mode = ttk.Combobox(root, values=MODES, state="readonly")
        self.mode.current(0)
        self.mode.pack(fill="x", **P)

        ef = ttk.Frame(root)
        ef.pack(fill="x", **P)
        cf = ttk.LabelFrame(ef, text="ক্রপ / ঘোরানো / ফ্লিপ (শুধু ছবি, pixel হিসাবে)")
        cf.pack(side="left", fill="both", expand=True, padx=(0, 4))
        cr = ttk.Frame(cf)
        cr.pack(fill="x", padx=6, pady=4)
        ttk.Button(cr, text="ক্রপ সেট করো", style="Purple.TButton", command=self.open_crop).pack(side="left", padx=2)
        ttk.Button(cr, text="ক্রপ বাদ দাও", command=self.clear_crop).pack(side="left", padx=2)
        self.crop_lbl = ttk.Label(cf, text="", wraplength=300, justify="left")
        self.crop_lbl.pack(anchor="w", padx=6, pady=(0, 6))

        pf = ttk.LabelFrame(ef, text="আউটপুট ছবির সাইজ (pixel) — খালি = যেমন আছে")
        pf.pack(side="left", fill="both", expand=True, padx=(4, 0))
        self.out_w = tk.StringVar()
        self.out_h = tk.StringVar()
        self.keep = tk.BooleanVar(value=True)
        pr = ttk.Frame(pf)
        pr.pack(fill="x", padx=6, pady=4)
        ttk.Label(pr, text="Width").pack(side="left")
        ttk.Entry(pr, textvariable=self.out_w, width=8).pack(side="left", padx=(4, 10))
        ttk.Label(pr, text="Height").pack(side="left")
        ttk.Entry(pr, textvariable=self.out_h, width=8).pack(side="left", padx=4)
        ttk.Checkbutton(pf, text="Aspect ratio ঠিক রাখো", variable=self.keep).pack(anchor="w", padx=6)
        ttk.Button(pf, text="পিক্সেল বাদ দাও", command=self.clear_px).pack(anchor="w", padx=6, pady=4)
        self.update_crop_lbl()

        self.rmbg = tk.BooleanVar(value=False)
        ttk.Checkbutton(root, text="ব্যাকগ্রাউন্ড মুছো (remove.bg, ইন্টারনেট লাগবে — শুধু ছবি)", variable=self.rmbg).pack(anchor="w", **P)

        out_lbl = ttk.Label(root, text="আউটপুট ফোল্ডার:")
        out_lbl.pack(anchor="w", **P)
        of = ttk.Frame(root)
        of.pack(fill="x", **P)
        ttk.Entry(of, textvariable=self.outdir).pack(side="left", fill="x", expand=True)
        ttk.Button(of, text="বদলাও", command=self.pick_out).pack(side="left", padx=4)
        ttk.Button(of, text="খোলো", style="Green.TButton", command=self.open_out).pack(side="left")

        self.btn = ttk.Button(root, text="কনভার্ট করো", style="Big.TButton", command=self.start)
        self.btn.pack(fill="x", padx=10, pady=8)
        self.bar = ttk.Progressbar(root, mode="determinate")
        self.bar.pack(fill="x", **P)
        self.status = ttk.Label(root, text="প্রস্তুত", wraplength=660, justify="left")
        self.status.pack(anchor="w", **P)

        # জানালা ছোট হলেও "কনভার্ট করো" বাটন, প্রগ্রেস বার ও স্ট্যাটাস যেন কখনো ঢাকা না পড়ে —
        # এগুলো আগে নিচে আটকে দেওয়া হলো (নিচ থেকে উপরে: status, bar, btn, ফোল্ডার সারি, লেবেল)
        for w in (self.status, self.bar, self.btn, of, out_lbl):
            w.pack_configure(side="bottom", before=hd)

        if HAVE_DND and hasattr(root, "drop_target_register"):
            try:
                root.drop_target_register(DND_FILES)
                root.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

    # file list
    def add_paths(self, paths, _top=True):
        added = False
        for p in paths:
            p = Path(p)
            if p.is_dir():
                self.add_paths(sorted(x for x in p.iterdir() if x.is_file()), False)
            elif p.suffix.lower() in ALL_EXT and str(p) not in self.files:
                self.files.append(str(p))
                label = p.name
                if p.suffix.lower() in IMG_EXT:
                    sz = img_size(p)
                    if sz:
                        label += f"  —  {sz[0]} × {sz[1]} px"
                self.lb.insert("end", label)
                added = True
        self.count.config(text=f"{len(self.files)}টি ফাইল")
        if added and _top:
            self.reset_edit()

    def on_drop(self, e):
        self.add_paths(self.root.tk.splitlist(e.data))

    def add_dialog(self):
        pats = " ".join("*" + x for x in sorted(ALL_EXT))
        fs = filedialog.askopenfilenames(filetypes=[("সাপোর্টেড ফাইল", pats), ("সব ফাইল", "*.*")])
        self.add_paths(fs)

    def add_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.add_paths([d])

    def remove_sel(self):
        for i in reversed(self.lb.curselection()):
            self.lb.delete(i)
            del self.files[i]
        self.count.config(text=f"{len(self.files)}টি ফাইল")
        self.reset_edit()

    def clear(self):
        self.lb.delete(0, "end")
        self.files.clear()
        self.count.config(text="০টি ফাইল")
        self.reset_edit()

    # crop / pixel helpers
    def reset_edit(self):
        self.crop = None
        self.ops = []
        self.update_crop_lbl()

    def update_crop_lbl(self):
        if self.crop is None:
            t = "ক্রপ: নেই (পুরো ছবি)"
        else:
            c = self.crop
            t = (f"ক্রপ: {c[2]} × {c[3]} px, Position X={c[0]}, Y={c[1]}\n"
                 "(সব ছবিতে এই মাপ প্রয়োগ হবে; ছবির বাইরে গেলে ছবির সীমা পর্যন্ত কাটবে)")
        if self.ops:
            t += f"\nঘোরানো/ফ্লিপ প্রয়োগ করা আছে ({len(self.ops)} ধাপ)"
        self.crop_lbl.config(text=t)

    def open_crop(self):
        first = next((f for f in self.files if Path(f).suffix.lower() in IMG_EXT), None)
        if not first:
            messagebox.showinfo("Converter", "আগে ছবি যোগ করো")
            return
        try:
            CropDialog(self, first)
        except Exception as e:
            messagebox.showinfo("Converter", f"ছবি লোড হয়নি: {e}")

    def clear_crop(self):
        self.crop = None
        self.ops = []
        self.update_crop_lbl()

    def clear_px(self):
        self.out_w.set("")
        self.out_h.set("")

    @staticmethod
    def _px(var):
        try:
            return max(0, min(int(var.get().strip()), 12000))
        except ValueError:
            return 0

    def pick_out(self):
        d = filedialog.askdirectory()
        if d:
            self.outdir.set(d)

    def open_out(self):
        p = Path(self.outdir.get())
        p.mkdir(parents=True, exist_ok=True)
        os.startfile(str(p))

    # conversion
    def start(self):
        if not self.files:
            messagebox.showinfo("Converter", "আগে ফাইল যোগ করো")
            return
        fmt, mode = self.fmt.current(), self.mode.current()
        target = 0
        t = self.size.get().strip().replace(",", ".")
        if t:
            try:
                target = int(float(t) * (1024 if self.unit.current() == 0 else 1048576))
            except ValueError:
                messagebox.showinfo("Converter", "সাইজ সংখ্যায় লেখো")
                return
        if mode > 0 and target <= 0:
            messagebox.showinfo("Converter", "এই মোডে সাইজ লিখে দিতে হবে")
            return
        opts = {"ops": list(self.ops), "crop": self.crop,
                "out_w": self._px(self.out_w), "out_h": self._px(self.out_h), "keep": self.keep.get(), "rmbg": self.rmbg.get()}
        if opts["rmbg"]:
            key = load_key()
            if not key:
                from tkinter import simpledialog
                key = (simpledialog.askstring("remove.bg API key", "remove.bg API key দাও (একবারই লাগবে, সেভ থাকবে):", parent=self.root) or "").strip()
                if not key:
                    return
                save_key(key)
            opts["key"] = key
        self.btn.state(["disabled"])
        threading.Thread(target=self.work, args=(list(self.files), fmt, target, mode, Path(self.outdir.get()), opts), daemon=True).start()
        self.poll()

    def work(self, files, fmt, target, mode, outdir, opts):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None
        try:
            msg = run(files, fmt, target, mode, outdir,
                      lambda s: self.q.put(("s", s)),
                      lambda d, t: self.q.put(("p", (d, t))), opts)
        except Exception as e:
            msg = f"ভুল: {e}"
        self.q.put(("done", msg))
        try:
            if pythoncom:
                pythoncom.CoUninitialize()
        except Exception:
            pass

    def poll(self):
        finished = False
        try:
            while True:
                k, v = self.q.get_nowait()
                if k == "s":
                    self.status.config(text=v)
                elif k == "p":
                    self.bar.config(maximum=v[1], value=v[0])
                else:
                    self.status.config(text=v)
                    finished = True
        except queue.Empty:
            pass
        if finished:
            self.btn.state(["!disabled"])
        else:
            self.root.after(120, self.poll)


def main():
    try:
        root = TkinterDnD.Tk() if HAVE_DND else tk.Tk()
    except Exception:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
