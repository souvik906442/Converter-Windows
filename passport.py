"""Passport-photo logic (Python port of the Android PassportActivity): crop box maths,
background removal (offline + remove.bg), background colour compositing and the A4 sheet.
No GUI code here, so it can be tested without Tk."""
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

import numpy as np
from PIL import Image

# name, width px, height px  (300 DPI)
SIZES = [("35×45 mm", 413, 531), ("2×2 inch", 600, 600)]
PAGE_W, PAGE_H, MARGIN, GAP = 2480, 3508, 59, 24      # A4 @ 300 DPI
WHITE = (255, 255, 255)
SKY = (135, 206, 235)
REMOVEBG_URL = "https://api.remove.bg/v1.0/removebg"
MODEL_NAME = "selfie_segmenter.tflite"

try:                                   # written by the GitHub workflow from the REMOVEBG_KEY secret
    from embedded_key import KEY as EMBEDDED_KEY
except Exception:
    EMBEDDED_KEY = ""
EMBEDDED_KEY = (EMBEDDED_KEY or "").strip()


# ---------------- sheet geometry ----------------

def max_cols(pw):
    return max(1, (PAGE_W - 2 * MARGIN + GAP) // (pw + GAP))


def max_rows(ph):
    return max(1, (PAGE_H - 2 * MARGIN + GAP) // (ph + GAP))


def max_copies(pw, ph):
    return max_cols(pw) * max_rows(ph)


def build_sheet(photo, n):
    pw, ph = photo.size
    rows = -(-n // max_cols(pw))
    cols = -(-n // rows)
    sheet = Image.new("RGB", (PAGE_W, PAGE_H), WHITE)
    from PIL import ImageDraw
    d = ImageDraw.Draw(sheet)
    tw, th = cols * pw + (cols - 1) * GAP, rows * ph + (rows - 1) * GAP
    x0, y0 = (PAGE_W - tw) // 2, (PAGE_H - th) // 2
    for i in range(n):
        r, c = divmod(i, cols)
        x, y = x0 + c * (pw + GAP), y0 + r * (ph + GAP)
        sheet.paste(photo, (x, y))
        d.rectangle([x, y, x + pw, y + ph], outline=(187, 187, 187), width=1)   # cut guide
    return sheet


def save_jpeg(img, path):
    img.convert("RGB").save(str(path), "JPEG", quality=95, dpi=(300, 300))


def save_pdf(sheet, path):
    """A4 page: 2480 px at 300 DPI = 595 pt, so a 100 % print keeps the real photo size."""
    sheet.convert("RGB").save(str(path), "PDF", resolution=300.0, quality=95)


# ---------------- crop box ----------------
# box = [x, y, w, h] in source-image pixels, w/h always equals the passport aspect ratio

def max_box_w(iw, ih, aspect):
    return min(float(iw), ih * aspect)


def init_box(iw, ih, aspect):
    w = max_box_w(iw, ih, aspect)
    h = w / aspect
    return [(iw - w) / 2, (ih - h) / 2, w, h]


def clamp_box(box, iw, ih, aspect, min_w=40.0):
    x, y, w, h = box
    w = max(min(min_w, max_box_w(iw, ih, aspect)), min(w, max_box_w(iw, ih, aspect)))
    h = w / aspect
    x = max(0.0, min(x, iw - w))
    y = max(0.0, min(y, ih - h))
    return [x, y, w, h]


def zoom_box(box, factor, iw, ih, aspect, fx=0.5, fy=0.5):
    """factor > 1 zooms in. (fx, fy) = fixed point inside the frame (0..1)."""
    x, y, w, h = box
    nw = w / factor
    nh = nw / aspect
    nx = x + (w - nw) * fx
    ny = y + (h - nh) * fy
    return clamp_box([nx, ny, nw, nh], iw, ih, aspect)


def crop_image(img, box, pw, ph):
    x, y, w, h = box
    return img.convert("RGB").resize((pw, ph), Image.LANCZOS, box=(x, y, x + w, y + h))


def crop_mask(mask, src_size, box, pw, ph):
    """mask: float array (mh, mw) covering the whole source image. Returns float array (ph, pw)."""
    mh, mw = mask.shape
    sx, sy = mw / src_size[0], mh / src_size[1]
    x, y, w, h = box
    m = Image.fromarray((np.clip(mask, 0, 1) * 255).astype(np.uint8), "L")
    m = m.resize((pw, ph), Image.BILINEAR, box=(x * sx, y * sy, (x + w) * sx, (y + h) * sy))
    return np.asarray(m, dtype=np.float32) / 255.0


# ---------------- compositing ----------------

def composite(fg, mask, bg, edge=50, is_alpha=False):
    """fg: PIL RGB, mask: float (h, w), bg: (r, g, b). edge 0..100 only matters for model masks."""
    a = mask.astype(np.float32)
    if not is_alpha:
        lo = 0.2 + 0.6 * edge / 100.0
        a = np.clip((a - lo) / 0.3, 0.0, 1.0)
        a = a * a * (3.0 - 2.0 * a)                   # soften the edge a little
    else:
        a = np.clip(a, 0.0, 1.0)
    f = np.asarray(fg.convert("RGB"), dtype=np.float32)
    b = np.array(bg, dtype=np.float32).reshape(1, 1, 3)
    out = f * a[..., None] + b * (1.0 - a[..., None])
    return Image.fromarray(np.clip(out, 0, 255).astype(np.uint8), "RGB")


# ---------------- remove.bg (online) ----------------

def remove_bg_mask(img, key, url=REMOVEBG_URL, timeout=60):
    """Send the cropped photo to remove.bg, return its alpha channel as a float mask."""
    buf = io.BytesIO()
    img.convert("RGB").save(buf, "JPEG", quality=95)
    boundary = "----pp%d" % int(time.time() * 1000)
    nl = b"\r\n"
    body = b""
    for k, v in (("size", "auto"), ("format", "png")):
        body += (("--%s" % boundary).encode() + nl +
                 ('Content-Disposition: form-data; name="%s"' % k).encode() + nl + nl + v.encode() + nl)
    body += (("--%s" % boundary).encode() + nl +
             b'Content-Disposition: form-data; name="image_file"; filename="photo.jpg"' + nl +
             b"Content-Type: image/jpeg" + nl + nl + buf.getvalue() + nl +
             ("--%s--" % boundary).encode() + nl)
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("X-Api-Key", key)
    req.add_header("Content-Type", "multipart/form-data; boundary=" + boundary)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = r.read()
    except urllib.error.HTTPError as e:
        txt = e.read().decode("utf-8", "replace")
        m = re.search(r'"title"\s*:\s*"([^"]*)"', txt)
        raise RuntimeError("HTTP %d%s" % (e.code, (": " + m.group(1)) if m else ""))
    try:
        res = Image.open(io.BytesIO(data)).convert("RGBA")
    except Exception:
        raise RuntimeError("রেজাল্ট পড়া যায়নি")
    if res.size != img.size:
        res = res.resize(img.size, Image.BILINEAR)
    return np.asarray(res.split()[3], dtype=np.float32) / 255.0


def split_keys(raw):
    return [k for k in re.split(r"[,;\s]+", raw or "") if k]


def remove_bg_with_keys(img, keys, start=0, url=REMOVEBG_URL):
    """Try keys in turn (quota 402 / key rejected 403 / rate limit 429 -> next key).
    Returns (mask, index of the key that worked)."""
    if not keys:
        raise RuntimeError("আগে remove.bg API key বসাও")
    start = start if 0 <= start < len(keys) else 0
    err = ""
    for n in range(len(keys)):
        i = (start + n) % len(keys)
        try:
            return remove_bg_mask(img, keys[i], url), i
        except Exception as e:
            err = str(e) or e.__class__.__name__
            if not (err.startswith("HTTP 402") or err.startswith("HTTP 403") or err.startswith("HTTP 429")):
                raise RuntimeError(err)
    raise RuntimeError(err + " (সব key চেষ্টা করা হয়েছে)")


# ---------------- offline background removal ----------------

def _model_file():
    cands = []
    if getattr(sys, "_MEIPASS", None):
        cands.append(Path(sys._MEIPASS) / MODEL_NAME)
    cands += [Path(__file__).resolve().parent / MODEL_NAME, Path.cwd() / MODEL_NAME]
    if getattr(sys, "frozen", False):
        cands.append(Path(sys.executable).resolve().parent / MODEL_NAME)
    for c in cands:
        if c.exists():
            return str(c)
    return None


def _orient(m):
    """Make sure 1 = person. The border of a portrait is almost always background."""
    h, w = m.shape
    bh, bw = max(1, h // 20), max(1, w // 20)
    border = np.concatenate([m[:bh].ravel(), m[-bh:].ravel(), m[:, :bw].ravel(), m[:, -bw:].ravel()])
    return 1.0 - m if border.mean() > 0.5 else m


def _mediapipe_mask(rgb):
    model = _model_file()
    if not model:
        raise RuntimeError("মডেল ফাইল নেই")
    import mediapipe as mp
    from mediapipe.tasks.python import vision
    from mediapipe.tasks.python.core.base_options import BaseOptions
    opts = vision.ImageSegmenterOptions(base_options=BaseOptions(model_asset_path=model),
                                        output_confidence_masks=True, output_category_mask=False)
    with vision.ImageSegmenter.create_from_options(opts) as seg:
        res = seg.segment(mp.Image(image_format=mp.ImageFormat.SRGB, data=np.ascontiguousarray(rgb)))
    masks = res.confidence_masks
    m = np.array(masks[1 if len(masks) > 1 else 0].numpy_view(), dtype=np.float32).reshape(rgb.shape[0], rgb.shape[1])
    return _orient(m)


def _grabcut_mask(rgb):
    """Simple fallback without any model: GrabCut seeded with a centred head-and-shoulders ellipse."""
    import cv2
    h, w = rgb.shape[:2]
    s = 512.0 / max(h, w)
    sm = cv2.resize(rgb, (max(8, round(w * s)), max(8, round(h * s))), interpolation=cv2.INTER_AREA)
    sh, sw = sm.shape[:2]
    gc = np.full((sh, sw), cv2.GC_PR_BGD, np.uint8)
    cv2.ellipse(gc, (sw // 2, int(sh * 0.50)), (int(sw * 0.32), int(sh * 0.40)), 0, 0, 360, cv2.GC_PR_FGD, -1)
    gc[int(sh * 0.80):, int(sw * 0.12):int(sw * 0.88)] = cv2.GC_PR_FGD          # shoulders
    cv2.ellipse(gc, (sw // 2, int(sh * 0.48)), (int(sw * 0.16), int(sh * 0.22)), 0, 0, 360, cv2.GC_FGD, -1)
    b = max(2, int(min(sh, sw) * 0.03))
    gc[:b], gc[-b:], gc[:, :b], gc[:, -b:] = cv2.GC_BGD, cv2.GC_BGD, cv2.GC_BGD, cv2.GC_BGD
    bgd, fgd = np.zeros((1, 65), np.float64), np.zeros((1, 65), np.float64)
    cv2.grabCut(cv2.cvtColor(sm, cv2.COLOR_RGB2BGR), gc, None, bgd, fgd, 5, cv2.GC_INIT_WITH_MASK)
    m = np.where((gc == cv2.GC_FGD) | (gc == cv2.GC_PR_FGD), 1.0, 0.0).astype(np.float32)
    m = cv2.GaussianBlur(m, (0, 0), max(1.0, min(sh, sw) / 150.0))
    return cv2.resize(m, (w, h), interpolation=cv2.INTER_LINEAR)


def offline_mask(img):
    """Person mask for the whole image (float 0..1, longest side <= 1024).
    Returns (mask, engine) with engine 'mediapipe' or 'grabcut'."""
    im = img.convert("RGB")
    m = max(im.size)
    if m > 1024:
        im = im.resize((round(im.width * 1024 / m), round(im.height * 1024 / m)), Image.LANCZOS)
    rgb = np.asarray(im, dtype=np.uint8)
    try:
        return _mediapipe_mask(rgb), "mediapipe"
    except Exception as e1:
        try:
            return _grabcut_mask(rgb), "grabcut"
        except Exception as e2:
            raise RuntimeError("অফলাইন মডেল চলেনি (%s / %s)" % (e1, e2))


# ---------------- settings ----------------

def cfg_path():
    base = os.environ.get("APPDATA") or str(Path.home())
    return Path(base) / "OfflineConverter" / "passport.json"


def load_cfg():
    try:
        return json.loads(cfg_path().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_cfg(d):
    try:
        p = cfg_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(d), encoding="utf-8")
    except Exception:
        pass
