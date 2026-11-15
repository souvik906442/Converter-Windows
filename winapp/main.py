import io, os, re, struct, zlib, zipfile, threading, queue, time, tempfile, shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
from PIL import Image, ImageOps

Image.MAX_IMAGE_PIXELS = None
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


def image_pages(path):
    im = Image.open(str(path))
    im = ImageOps.exif_transpose(im)
    im.load()
    if max(im.size) > 6000:
        im.thumbnail((6000, 6000))
    yield im


def run(files, fmt, target, mode, outdir, say, prog):
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
                gen = pdf_pages(p) if kind == "pdf" else image_pages(p)
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


# ---------------- GUI ----------------

class App:
    def __init__(self, root):
        self.root = root
        root.title("Offline Converter")
        root.geometry("700x720")
        try:
            for nm in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
                tkfont.nametofont(nm).configure(family="Nirmala UI", size=10)
        except Exception:
            pass
        self.files = []
        self.q = queue.Queue()
        self.outdir = tk.StringVar(value=str(Path.home() / "Downloads" / "Converter"))
        self.size = tk.StringVar()
        P = dict(padx=10, pady=4)

        ttk.Label(root, text="ফাইল বাছাই করো (ছবি / PDF / Word / Excel / PowerPoint) — ড্র্যাগ করেও ফেলতে পারো").pack(anchor="w", **P)
        fr = ttk.Frame(root)
        fr.pack(fill="both", expand=True, **P)
        self.lb = tk.Listbox(fr, height=10, selectmode="extended")
        sb = ttk.Scrollbar(fr, command=self.lb.yview)
        self.lb.config(yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        bf = ttk.Frame(root)
        bf.pack(fill="x", **P)
        ttk.Button(bf, text="ফাইল যোগ করো", command=self.add_dialog).pack(side="left", padx=2)
        ttk.Button(bf, text="ফোল্ডার যোগ করো", command=self.add_folder).pack(side="left", padx=2)
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

        ttk.Label(root, text="আউটপুট ফোল্ডার:").pack(anchor="w", **P)
        of = ttk.Frame(root)
        of.pack(fill="x", **P)
        ttk.Entry(of, textvariable=self.outdir).pack(side="left", fill="x", expand=True)
        ttk.Button(of, text="বদলাও", command=self.pick_out).pack(side="left", padx=4)
        ttk.Button(of, text="খোলো", command=self.open_out).pack(side="left")

        self.btn = ttk.Button(root, text="কনভার্ট করো", command=self.start)
        self.btn.pack(fill="x", padx=10, pady=8)
        self.bar = ttk.Progressbar(root, mode="determinate")
        self.bar.pack(fill="x", **P)
        self.status = ttk.Label(root, text="প্রস্তুত", wraplength=660, justify="left")
        self.status.pack(anchor="w", **P)

        if HAVE_DND and hasattr(root, "drop_target_register"):
            try:
                root.drop_target_register(DND_FILES)
                root.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

    # file list
    def add_paths(self, paths):
        for p in paths:
            p = Path(p)
            if p.is_dir():
                self.add_paths(sorted(x for x in p.iterdir() if x.is_file()))
            elif p.suffix.lower() in ALL_EXT and str(p) not in self.files:
                self.files.append(str(p))
                self.lb.insert("end", p.name)
        self.count.config(text=f"{len(self.files)}টি ফাইল")

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

    def clear(self):
        self.lb.delete(0, "end")
        self.files.clear()
        self.count.config(text="০টি ফাইল")

    def pick_out(self):
        d = filedialog.askdirectory()
        if d:
            self.outdir.set(d)

    def open_out(self):
        p = Path(self.outdir.get())
        p.mkdir(parents=True, exist_ok=True)
   
