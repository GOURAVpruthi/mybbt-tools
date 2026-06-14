"""
GST Tools Suite
===============
File Collector  +  PDF Compressor

Run by double-clicking this file, or:
    python gst_tools_suite.py
"""

# ── auto-install missing packages BEFORE any other import ───────
import subprocess, sys, os

def _install(pkg):
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", pkg, "--quiet"],
                       capture_output=True, timeout=240)
    except Exception:
        pass

# Try each package. pikepdf is OPTIONAL — app works without it.
for pip_name, import_name in [("pillow", "PIL"),
                              ("PyMuPDF", "fitz"),
                              ("pikepdf", "pikepdf"),
                              ("openpyxl", "openpyxl"),
                              ("pdfplumber", "pdfplumber")]:
    try:
        __import__(import_name)
    except Exception:
        _install(pip_name)

# ── now import what we can, with fallbacks ──────────────────────
import re
import io
import shutil
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import fitz
    HAS_PYMUPDF = True
except Exception:
    HAS_PYMUPDF = False

try:
    from PIL import Image
    HAS_PILLOW = True
except Exception:
    HAS_PILLOW = False

try:
    import pikepdf
    HAS_PIKEPDF = True
except Exception:
    HAS_PIKEPDF = False

try:
    import openpyxl as _openpyxl_check
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

try:
    import pdfplumber as _pdfplumber_check
    HAS_PDFPLUMBER = True
except Exception:
    HAS_PDFPLUMBER = False


# ════════════════════════════════════════════════════════════════
#  PALETTE
# ════════════════════════════════════════════════════════════════
BG       = "#1e1e2e"
CARD     = "#2a2a3e"
ACCENT   = "#6c63ff"
ACCENT2  = "#a78bfa"
TEXT     = "#e2e8f0"
SUBTEXT  = "#94a3b8"
SUCCESS  = "#22c55e"
WARNING  = "#f59e0b"
ERROR    = "#ef4444"
ENTRY_BG = "#16213e"
BORDER   = "#3d3d5c"


# ════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════
_SEPARATORS_RE = re.compile(r"[\s\-_\.]+")

def _normalize_for_match(s):
    return _SEPARATORS_RE.sub("", s.lower())

def _keyword_matches(filename, keyword):
    return _normalize_for_match(keyword) in _normalize_for_match(filename)

EXT_PRESETS = {
    "Any (all formats)":      None,
    "PDF (.pdf)":             (".pdf",),
    "Excel (.xlsx, .xls)":    (".xlsx", ".xls", ".xlsm", ".xlsb"),
    "CSV (.csv)":             (".csv",),
    "JSON (.json)":           (".json",),
    "Word (.docx, .doc)":     (".docx", ".doc"),
    "Text (.txt)":            (".txt",),
    "Custom...":              "CUSTOM",
}

def _ext_matches(filename, ext_tuple):
    if ext_tuple is None: return True
    fn = filename.lower()
    return any(fn.endswith(e) for e in ext_tuple)


# ════════════════════════════════════════════════════════════════
#  PDF COMPRESSION ENGINES
# ════════════════════════════════════════════════════════════════
def compress_lossless(in_path, out_path):
    doc = fitz.open(in_path)
    try:
        doc.save(out_path, garbage=4, deflate=True, deflate_images=True,
                 deflate_fonts=True, clean=True)
    finally:
        doc.close()


def compress_images_preserve_text(in_path, out_path, jpeg_quality, downscale=1.0):
    if not HAS_PIKEPDF or not HAS_PILLOW:
        raise RuntimeError("pikepdf or pillow not available")
    pdf = pikepdf.Pdf.open(in_path)
    try:
        for page in pdf.pages:
            try: images = page.images
            except Exception: continue
            for name in list(images.keys()):
                raw_image = images[name]
                try:
                    pdfimg = pikepdf.PdfImage(raw_image)
                    pil = pdfimg.as_pil_image()
                    if pil.mode in ("RGBA", "P", "LA"):
                        pil = pil.convert("RGB")
                    elif pil.mode == "CMYK":
                        pil = pil.convert("RGB")
                    if downscale < 1.0:
                        w, h = pil.size
                        pil = pil.resize((max(1, int(w * downscale)),
                                          max(1, int(h * downscale))),
                                         Image.LANCZOS)
                    buf = io.BytesIO()
                    pil.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
                    new_bytes = buf.getvalue()
                    try: raw_size = len(bytes(raw_image.read_raw_bytes()))
                    except Exception: raw_size = 10**9
                    if len(new_bytes) >= raw_size: continue
                    cs = pikepdf.Name("/DeviceGray") if pil.mode == "L" \
                         else pikepdf.Name("/DeviceRGB")
                    raw_image.write(new_bytes, filter=pikepdf.Name("/DCTDecode"))
                    raw_image.Width = pil.width
                    raw_image.Height = pil.height
                    raw_image.BitsPerComponent = 8
                    raw_image.ColorSpace = cs
                    for k in ("/DecodeParms", "/Decode"):
                        if k in raw_image: del raw_image[k]
                except Exception: continue
        pdf.save(out_path, compress_streams=True,
                 object_stream_mode=pikepdf.ObjectStreamMode.generate,
                 recompress_flate=True, linearize=False)
    finally:
        pdf.close()


def compress_rasterize(in_path, out_path, dpi, jpeg_quality):
    doc = fitz.open(in_path)
    new_doc = fitz.open()
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            img_bytes = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
            new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=img_bytes)
        new_doc.save(out_path, garbage=4, deflate=True)
    finally:
        new_doc.close()
        doc.close()


def smart_compress(in_path, out_path, target_mb, allow_text_loss, log):
    # SAFETY: never let output be the same path as input
    if os.path.abspath(in_path) == os.path.abspath(out_path):
        raise ValueError(
            f"Input and output paths are the same:\n  {in_path}\n"
            "Pick a different output folder, or this tool will refuse "
            "to overwrite the source file.")

    target_bytes = int(target_mb * 1024 * 1024)
    orig_bytes = os.path.getsize(in_path)

    if orig_bytes <= target_bytes:
        shutil.copy2(in_path, out_path)
        log(f"   ℹ  Already ≤ target — copied as-is.", "dim")
        return orig_bytes, "already optimal", True

    best_size, best_strategy = orig_bytes, "original"
    try:
        compress_lossless(in_path, out_path)
        sz = os.path.getsize(out_path)
        pct = (1 - sz / orig_bytes) * 100
        if sz <= target_bytes:
            log(f"   ✓  Lossless: {sz/1024/1024:.2f} MB  (-{pct:.0f}%)  ✅ Target hit", "ok")
            return sz, "lossless", True
        log(f"   ↺  Lossless: {sz/1024/1024:.2f} MB  (-{pct:.0f}%), > target", "ren")
        best_size, best_strategy = sz, "lossless"
    except Exception as e:
        log(f"   ⚠  Lossless step failed: {e}", "err")
        # Fall back to copy of original as starting baseline
        try:
            shutil.copy2(in_path, out_path)
        except Exception as e2:
            log(f"   ⚠  Could not even copy original: {e2}", "err")
            raise

    if HAS_PIKEPDF and HAS_PILLOW:
        for q, scale in [(80, 1.0), (70, 1.0), (60, 0.9),
                         (50, 0.8), (40, 0.7), (30, 0.6)]:
            try:
                tmp = out_path + ".t2.tmp.pdf"
                compress_images_preserve_text(in_path, tmp, q, scale)
                sz = os.path.getsize(tmp)
                if sz < best_size:
                    shutil.move(tmp, out_path)
                    best_size = sz
                    best_strategy = f"images q={q} scale={scale:.2f}"
                else:
                    try: os.remove(tmp)
                    except: pass
                if sz <= target_bytes:
                    pct = (1 - sz / orig_bytes) * 100
                    log(f"   ✓  Images q={q} s={scale}: "
                        f"{sz/1024/1024:.2f} MB (-{pct:.0f}%) ✅", "ok")
                    return sz, best_strategy, True
            except Exception:
                continue
    else:
        log("   ⚠  pikepdf not available — image-recompress tier skipped.", "dim")

    if not allow_text_loss:
        log(f"   ⚠  Target not reached. Best: {best_size/1024/1024:.2f} MB "
            f"({best_strategy}). Switch to 'Allow text loss' to go smaller.", "dim")
        return best_size, best_strategy, False

    for dpi, q in [(200, 80), (150, 75), (120, 70), (100, 60),
                   (85, 55), (72, 50), (60, 40), (50, 30)]:
        try:
            tmp = out_path + ".t3.tmp.pdf"
            compress_rasterize(in_path, tmp, dpi, q)
            sz = os.path.getsize(tmp)
            if sz < best_size:
                shutil.move(tmp, out_path)
                best_size = sz
                best_strategy = f"rasterized @ {dpi}dpi q{q} (TEXT LOST)"
            else:
                try: os.remove(tmp)
                except: pass
            if sz <= target_bytes:
                pct = (1 - sz / orig_bytes) * 100
                log(f"   ⚠  Rasterize @ {dpi}dpi q{q}: "
                    f"{sz/1024/1024:.2f} MB ⚠ TEXT LOST", "ren")
                return sz, best_strategy, True
        except Exception:
            continue

    for sfx in (".t2.tmp.pdf", ".t3.tmp.pdf"):
        p = out_path + sfx
        if os.path.exists(p):
            try: os.remove(p)
            except: pass
    log(f"   ⚠  Target unreachable. Best: {best_size/1024/1024:.2f} MB", "ren")
    return best_size, best_strategy, False


# ════════════════════════════════════════════════════════════════
#  GSTR-1 / GSTR-3B  EXTRACTOR ENGINES  (embedded)
# ════════════════════════════════════════════════════════════════
# The full GSTR-1 and GSTR-3B PDF-to-Excel extractor engines are
# embedded here as base64-encoded source.  They run in their own
# isolated namespaces so their constants (STATE_CODES, NUM_FMT,
# etc.) don't collide with the rest of the suite.

import base64 as _base64

_GSTR1_ENGINE_B64 = (
    "U1RBVEVfQ09ERVMgPSB7CiAgICAiMDEiOiAiSmFtbXUgJiBLYXNobWlyIiwgIjAyIjogIkhpbWFj"
    "aGFsIFByYWRlc2giLCAiMDMiOiAiUHVuamFiIiwKICAgICIwNCI6ICJDaGFuZGlnYXJoIiwgIjA1"
    "IjogIlV0dGFyYWtoYW5kIiwgIjA2IjogIkhhcnlhbmEiLCAiMDciOiAiRGVsaGkiLAogICAgIjA4"
    "IjogIlJhamFzdGhhbiIsICIwOSI6ICJVdHRhciBQcmFkZXNoIiwgIjEwIjogIkJpaGFyIiwgIjEx"
    "IjogIlNpa2tpbSIsCiAgICAiMTIiOiAiQXJ1bmFjaGFsIFByYWRlc2giLCAiMTMiOiAiTmFnYWxh"
    "bmQiLCAiMTQiOiAiTWFuaXB1ciIsICIxNSI6ICJNaXpvcmFtIiwKICAgICIxNiI6ICJUcmlwdXJh"
    "IiwgIjE3IjogIk1lZ2hhbGF5YSIsICIxOCI6ICJBc3NhbSIsICIxOSI6ICJXZXN0IEJlbmdhbCIs"
    "CiAgICAiMjAiOiAiSmhhcmtoYW5kIiwgIjIxIjogIk9kaXNoYSIsICIyMiI6ICJDaGhhdHRpc2dh"
    "cmgiLCAiMjMiOiAiTWFkaHlhIFByYWRlc2giLAogICAgIjI0IjogIkd1amFyYXQiLCAiMjUiOiAi"
    "RGFtYW4gJiBEaXUiLCAiMjYiOiAiRGFkcmEgJiBOYWdhciBIYXZlbGkgYW5kIERhbWFuICYgRGl1"
    "IiwKICAgICIyNyI6ICJNYWhhcmFzaHRyYSIsICIyOCI6ICJBbmRocmEgUHJhZGVzaCAoT2xkKSIs"
    "ICIyOSI6ICJLYXJuYXRha2EiLAogICAgIjMwIjogIkdvYSIsICIzMSI6ICJMYWtzaGFkd2VlcCIs"
    "ICIzMiI6ICJLZXJhbGEiLCAiMzMiOiAiVGFtaWwgTmFkdSIsCiAgICAiMzQiOiAiUHVkdWNoZXJy"
    "eSIsICIzNSI6ICJBbmRhbWFuICYgTmljb2JhciBJc2xhbmRzIiwgIjM2IjogIlRlbGFuZ2FuYSIs"
    "CiAgICAiMzciOiAiQW5kaHJhIFByYWRlc2giLCAiMzgiOiAiTGFkYWtoIiwgIjk3IjogIk90aGVy"
    "IFRlcnJpdG9yeSIsICI5OSI6ICJDZW50cmUgSnVyaXNkaWN0aW9uIiwKfQoKTU9OVEhfQUJCUiA9"
    "IHsKICAgICJKYW51YXJ5IjogIkphbiIsICJGZWJydWFyeSI6ICJGZWIiLCAiTWFyY2giOiAiTWFy"
    "IiwgIkFwcmlsIjogIkFwciIsCiAgICAiTWF5IjogICAgICJNYXkiLCAiSnVuZSI6ICAgICAiSnVu"
    "IiwgIkp1bHkiOiAgIkp1bCIsICJBdWd1c3QiOiAgICJBdWciLAogICAgIlNlcHRlbWJlciI6ICJT"
    "ZXAiLCAiT2N0b2JlciI6ICJPY3QiLCAiTm92ZW1iZXIiOiAiTm92IiwgIkRlY2VtYmVyIjogIkRl"
    "YyIsCn0KCiMgQ29sdW1uIHNjaGVtYXMg4oCUIHdoaWNoIG51bWVyaWMgY29sdW1ucyBlYWNoIHZh"
    "bHVlLWxpbmUgY2FycmllcyAoaW4gb3JkZXIpClNDSEVNQV81ID0gWyJ0YXhhYmxlIiwgImlnc3Qi"
    "LCAiY2dzdCIsICJzZ3N0IiwgImNlc3MiXSAgICMgaW50cmEtc3RhdGUgZnVsbCBzZXQKU0NIRU1B"
    "XzMgPSBbInRheGFibGUiLCAiaWdzdCIsICJjZXNzIl0gICAgICAgICAgICAgICAgICAgICMgaW50"
    "ZXItc3RhdGUgb25seSAobm8gQ0dTVC9TR1NUKQpTQ0hFTUFfMSA9IFsidGF4YWJsZSJdICAgICAg"
    "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIyB2YWx1ZS1vbmx5IHJvd3MKCgojIFRhYmxl"
    "IGV4dHJhY3Rpb24gcnVsZXMuCiMgRWFjaCBydWxlIHNheXM6IGxvY2F0ZSB0aGlzIGhlYWRpbmcs"
    "IHRoZW4gZmluZCB0aGUgZGF0YSBsaW5lIHRoYXQgYmVnaW5zCiMgd2l0aCB0aGlzIG1hcmtlciwg"
    "dGhlbiBtYXAgdGhlIG51bWJlcnMgZm91bmQgdGhlcmUgdG8gdGhlc2UgY29sdW1ucy4KVEFSR0VU"
    "UyA9IFsKICAgICMgLS0tLSBDdXJyZW50IHBlcmlvZCBvdXR3YXJkIHN1cHBsaWVzIC0tLS0KICAg"
    "ICgiNEEiLCAgIkIyQiBSZWd1bGFyIiwKICAgICAiNEEgLSBUYXhhYmxlIG91dHdhcmQgc3VwcGxp"
    "ZXMgbWFkZSB0byByZWdpc3RlcmVkIHBlcnNvbnMgKG90aGVyIHRoYW4iLAogICAgICJUb3RhbCAi"
    "LCBTQ0hFTUFfNSksCiAgICAoIjRCIiwgICJCMkIgUmV2ZXJzZSBDaGFyZ2UiLAogICAgICI0QiAt"
    "IFRheGFibGUgb3V0d2FyZCBzdXBwbGllcyBtYWRlIHRvIHJlZ2lzdGVyZWQgcGVyc29ucyBhdHRy"
    "YWN0aW5nIHRheCBvbiByZXZlcnNlIiwKICAgICAiVG90YWwgIiwgU0NIRU1BXzUpLAogICAgKCI1"
    "IiwgICAiQjJDTCAoTGFyZ2UpIiwKICAgICAiNSAtIFRheGFibGUgb3V0d2FyZCBpbnRlci1zdGF0"
    "ZSBzdXBwbGllcyBtYWRlIHRvIHVucmVnaXN0ZXJlZCBwZXJzb25zIiwKICAgICAiVG90YWwgIiwg"
    "U0NIRU1BXzMpLAoKICAgICMgNkEgRXhwb3J0cyDigJQgVG90YWwgcm93IChhZ2dyZWdhdGVzIEVY"
    "UFdQICsgRVhQV09QKSArIHN1Yi1yb3dzCiAgICAoIjZBIiwgICJFeHBvcnRzIC0gVG90YWwgKDZB"
    "KSIsCiAgICAgIjZBIOKAkyBFeHBvcnRzIiwgIlRvdGFsICIsIFNDSEVNQV8zKSwKICAgICgiNkEi"
    "LCAgIkV4cG9ydHMgLSBFWFBXUCAod2l0aCBwYXltZW50KSIsCiAgICAgIjZBIOKAkyBFeHBvcnRz"
    "IiwgIi0gRVhQV1AiLCBTQ0hFTUFfMyksCiAgICAoIjZBIiwgICJFeHBvcnRzIC0gRVhQV09QICh3"
    "aXRob3V0IHBheW1lbnQpIiwKICAgICAiNkEg4oCTIEV4cG9ydHMiLCAiLSBFWFBXT1AiLCBTQ0hF"
    "TUFfMSksCgogICAgIyA2QiBTRVog4oCUIFRvdGFsIHJvdyArIHN1Yi1yb3dzCiAgICAoIjZCIiwg"
    "ICJTRVogLSBUb3RhbCAoNkIpIiwKICAgICAiNkIgLSBTdXBwbGllcyBtYWRlIHRvIFNFWiIsICJU"
    "b3RhbCAiLCBTQ0hFTUFfMyksCiAgICAoIjZCIiwgICJTRVogLSBTRVpXUCAod2l0aCBwYXltZW50"
    "KSIsCiAgICAgIjZCIC0gU3VwcGxpZXMgbWFkZSB0byBTRVoiLCAiLSBTRVpXUCIsIFNDSEVNQV8z"
    "KSwKICAgICgiNkIiLCAgIlNFWiAtIFNFWldPUCAod2l0aG91dCBwYXltZW50KSIsCiAgICAgIjZC"
    "IC0gU3VwcGxpZXMgbWFkZSB0byBTRVoiLCAiLSBTRVpXT1AiLCBTQ0hFTUFfMSksCgogICAgKCI2"
    "QyIsICAiRGVlbWVkIEV4cG9ydHMgKERFKSIsCiAgICAgIjZDIC0gRGVlbWVkIEV4cG9ydHMiLCAi"
    "VG90YWwgIiwgU0NIRU1BXzUpLAoKICAgICgiNyIsICAgIkIyQ1MgKE90aGVycykiLAogICAgICI3"
    "LSBUYXhhYmxlIHN1cHBsaWVzIiwgIlRvdGFsICIsIFNDSEVNQV81KSwKCiAgICAjIDgg4oCUIHN1"
    "Yi1yb3dzCiAgICAoIjgiLCAgICJOaWwgcmF0ZWQiLAogICAgICI4IC0gTmlsIHJhdGVkLCBleGVt"
    "cHRlZCIsICItIE5pbCIsIFNDSEVNQV8xKSwKICAgICgiOCIsICAgIkV4ZW1wdGVkIiwKICAgICAi"
    "OCAtIE5pbCByYXRlZCwgZXhlbXB0ZWQiLCAiLSBFeGVtcHRlZCIsIFNDSEVNQV8xKSwKICAgICgi"
    "OCIsICAgIk5vbi1HU1QiLAogICAgICI4IC0gTmlsIHJhdGVkLCBleGVtcHRlZCIsICItIE5vbi1H"
    "U1QiLCBTQ0hFTUFfMSksCgogICAgIyAtLS0tIDlBIEFtZW5kbWVudHMgLS0tLQogICAgKCI5QSIs"
    "ICAiQW1lbmRtZW50IC0gQjJCIFJlZ3VsYXIiLAogICAgICI5QSAtIEFtZW5kbWVudCB0byB0YXhh"
    "YmxlIG91dHdhcmQgc3VwcGxpZXMgbWFkZSB0byByZWdpc3RlcmVkIHBlcnNvbiBpbiByZXR1cm5z"
    "IG9mIGVhcmxpZXIgdGF4IHBlcmlvZHMgaW4gdGFibGUgNCAtIEIyQiBSZWd1bGFyIiwKICAgICAi"
    "TmV0IGRpZmZlcmVudGlhbCBhbW91bnQiLCBTQ0hFTUFfNSksCiAgICAoIjlBIiwgICJBbWVuZG1l"
    "bnQgLSBCMkIgUmV2ZXJzZSBDaGFyZ2UiLAogICAgICI5QSAtIEFtZW5kbWVudCB0byB0YXhhYmxl"
    "IG91dHdhcmQgc3VwcGxpZXMgbWFkZSB0byByZWdpc3RlcmVkIHBlcnNvbiBpbiByZXR1cm5zIG9m"
    "IGVhcmxpZXIgdGF4IHBlcmlvZHMgaW4gdGFibGUgNCAtIEIyQiBSZXZlcnNlIGNoYXJnZSIsCiAg"
    "ICAgIk5ldCBkaWZmZXJlbnRpYWwgYW1vdW50IiwgU0NIRU1BXzUpLAogICAgKCI5QSIsICAiQW1l"
    "bmRtZW50IC0gQjJDTCIsCiAgICAgIjlBIC0gQW1lbmRtZW50IHRvIEludGVyLVN0YXRlIHN1cHBs"
    "aWVzIG1hZGUgdG8gdW5yZWdpc3RlcmVkIHBlcnNvbiIsCiAgICAgIk5ldCBkaWZmZXJlbnRpYWwg"
    "YW1vdW50IiwgU0NIRU1BXzMpLAogICAgKCI5QSIsICAiQW1lbmRtZW50IC0gRXhwb3J0cyBFWFBX"
    "UCIsCiAgICAgIjlBIC0gQW1lbmRtZW50IHRvIEV4cG9ydCBzdXBwbGllcyIsICItIEVYUFdQIiwg"
    "U0NIRU1BXzMpLAogICAgKCI5QSIsICAiQW1lbmRtZW50IC0gRXhwb3J0cyBFWFBXT1AiLAogICAg"
    "ICI5QSAtIEFtZW5kbWVudCB0byBFeHBvcnQgc3VwcGxpZXMiLCAiLSBFWFBXT1AiLCBTQ0hFTUFf"
    "MSksCiAgICAoIjlBIiwgICJBbWVuZG1lbnQgLSBTRVogU0VaV1AiLAogICAgICI5QSAtIEFtZW5k"
    "bWVudCB0byBzdXBwbGllcyBtYWRlIHRvIFNFWiIsICItIFNFWldQIiwgU0NIRU1BXzMpLAogICAg"
    "KCI5QSIsICAiQW1lbmRtZW50IC0gU0VaIFNFWldPUCIsCiAgICAgIjlBIC0gQW1lbmRtZW50IHRv"
    "IHN1cHBsaWVzIG1hZGUgdG8gU0VaIiwgIi0gU0VaV09QIiwgU0NIRU1BXzEpLAogICAgKCI5QSIs"
    "ICAiQW1lbmRtZW50IC0gRGVlbWVkIEV4cG9ydHMiLAogICAgICI5QSAtIEFtZW5kbWVudCB0byBE"
    "ZWVtZWQgRXhwb3J0cyIsCiAgICAgIk5ldCBkaWZmZXJlbnRpYWwgYW1vdW50IiwgU0NIRU1BXzUp"
    "LAoKICAgICMgLS0tLSA5QiBDcmVkaXQvRGViaXQgTm90ZXMgLS0tLQogICAgKCI5QiIsICAiQ0RO"
    "UiAtIENyZWRpdC9EZWJpdCBOb3RlcyAoUmVnaXN0ZXJlZCkiLAogICAgICI5QiAtIENyZWRpdC9E"
    "ZWJpdCBOb3RlcyAoUmVnaXN0ZXJlZCkiLAogICAgICJUb3RhbCAtIE5ldCBvZmYgZGViaXQvY3Jl"
    "ZGl0IG5vdGVzIiwgU0NIRU1BXzUpLAogICAgKCI5QiIsICAiQ0ROVVIgLSBCMkNMIiwKICAgICAi"
    "OUIgLSBDcmVkaXQvRGViaXQgTm90ZXMgKFVucmVnaXN0ZXJlZCkiLCAiLSBCMkNMIiwgU0NIRU1B"
    "XzMpLAogICAgKCI5QiIsICAiQ0ROVVIgLSBFWFBXUCIsCiAgICAgIjlCIC0gQ3JlZGl0L0RlYml0"
    "IE5vdGVzIChVbnJlZ2lzdGVyZWQpIiwgIi0gRVhQV1AiLCBTQ0hFTUFfMyksCiAgICAoIjlCIiwg"
    "ICJDRE5VUiAtIEVYUFdPUCIsCiAgICAgIjlCIC0gQ3JlZGl0L0RlYml0IE5vdGVzIChVbnJlZ2lz"
    "dGVyZWQpIiwgIi0gRVhQV09QIiwgU0NIRU1BXzEpLAoKICAgICMgLS0tLSA5QyBBbWVuZGVkIENE"
    "TiAtLS0tCiAgICAoIjlDIiwgICJDRE5SQSAtIEFtZW5kZWQgQ0ROIChSZWdpc3RlcmVkKSIsCiAg"
    "ICAgIjlDIC0gQW1lbmRlZCBDcmVkaXQvRGViaXQgTm90ZXMgKFJlZ2lzdGVyZWQpIiwKICAgICAi"
    "TmV0IERpZmZlcmVudGlhbCBhbW91bnQiLCBTQ0hFTUFfNSksCiAgICAoIjlDIiwgICJDRE5VUkEg"
    "LSBCMkNMIiwKICAgICAiOUMgLSBBbWVuZGVkIENyZWRpdC9EZWJpdCBOb3RlcyAoVW5yZWdpc3Rl"
    "cmVkKSIsICItIEIyQ0wiLCBTQ0hFTUFfMyksCiAgICAoIjlDIiwgICJDRE5VUkEgLSBFWFBXUCIs"
    "CiAgICAgIjlDIC0gQW1lbmRlZCBDcmVkaXQvRGViaXQgTm90ZXMgKFVucmVnaXN0ZXJlZCkiLCAi"
    "LSBFWFBXUCIsIFNDSEVNQV8zKSwKICAgICgiOUMiLCAgIkNETlVSQSAtIEVYUFdPUCIsCiAgICAg"
    "IjlDIC0gQW1lbmRlZCBDcmVkaXQvRGViaXQgTm90ZXMgKFVucmVnaXN0ZXJlZCkiLCAiLSBFWFBX"
    "T1AiLCBTQ0hFTUFfMSksCgogICAgKCIxMCIsICAiQW1lbmRtZW50IHRvIEIyQyAoT3RoZXJzKSIs"
    "CiAgICAgIjEwIC0gQW1lbmRtZW50IHRvIHRheGFibGUgb3V0d2FyZCBzdXBwbGllcyBtYWRlIHRv"
    "IHVucmVnaXN0ZXJlZCBwZXJzb24iLAogICAgICJOZXQgZGlmZmVyZW50aWFsIGFtb3VudCIsIFND"
    "SEVNQV81KSwKCiAgICAjIC0tLS0gMTEgQWR2YW5jZXMgLS0tLQogICAgKCIxMUEiLCAiQWR2YW5j"
    "ZXMgcmVjZWl2ZWQiLAogICAgICIxMUEoMSksIDExQSgyKSAtIEFkdmFuY2VzIHJlY2VpdmVkIiwg"
    "IlRvdGFsICIsIFNDSEVNQV81KSwKICAgICgiMTFCIiwgIkFkdmFuY2VzIGFkanVzdGVkIGFnYWlu"
    "c3QgY3VycmVudC1wZXJpb2Qgc3VwcGxpZXMiLAogICAgICIxMUIoMSksIDExQigyKSAtIEFkdmFu"
    "Y2UgYW1vdW50IHJlY2VpdmVkIGluIGVhcmxpZXIgdGF4IHBlcmlvZCIsCiAgICAgIlRvdGFsICIs"
    "IFNDSEVNQV81KSwKICAgICgiMTFBIiwgIkFtZW5kbWVudCB0byBBZHZhbmNlcyByZWNlaXZlZCIs"
    "CiAgICAgIjExQSAtIEFtZW5kbWVudCB0byBhZHZhbmNlcyByZWNlaXZlZCIsICJUb3RhbCAiLCBT"
    "Q0hFTUFfNSksCiAgICAoIjExQiIsICJBbWVuZG1lbnQgdG8gQWR2YW5jZXMgYWRqdXN0ZWQiLAog"
    "ICAgICIxMUIgLSBBbWVuZG1lbnQgdG8gYWR2YW5jZXMgYWRqdXN0ZWQiLCAiVG90YWwgIiwgU0NI"
    "RU1BXzUpLAoKICAgICMgLS0tLSAxMiBIU04gLS0tLQogICAgKCIxMiIsICAiSFNOLXdpc2Ugc3Vt"
    "bWFyeSIsCiAgICAgIjEyIC0gSFNOLXdpc2Ugc3VtbWFyeSIsICJUb3RhbCAiLCBTQ0hFTUFfNSks"
    "CgogICAgIyAtLS0tIDEzIERvY3VtZW50cyAoY291bnQgb25seSkgLS0tLQogICAgKCIxMyIsICAi"
    "RG9jdW1lbnRzIGlzc3VlZCAoY291bnQgb25seSkiLAogICAgICIxMyAtIERvY3VtZW50cyBpc3N1"
    "ZWQiLCAiTmV0IGlzc3VlZCBkb2N1bWVudHMiLCBbXSksCgogICAgIyAtLS0tIDE0IEVDTyAtLS0t"
    "CiAgICAoIjE0IiwgICJTdXBwbGllcyB0aHJvdWdoIEUtQ29tbWVyY2UgT3BlcmF0b3JzIiwKICAg"
    "ICAiMTQgLSBTdXBwbGllcyBtYWRlIHRocm91Z2ggRS1Db21tZXJjZSBPcGVyYXRvcnMiLCAiVG90"
    "YWwgIiwgU0NIRU1BXzUpLAogICAgKCIxNEEiLCAiQW1lbmRlZCBTdXBwbGllcyB0aHJvdWdoIEVD"
    "TyIsCiAgICAgIjE0QSAtIEFtZW5kZWQgU3VwcGxpZXMgbWFkZSB0aHJvdWdoIEUtQ29tbWVyY2Ug"
    "T3BlcmF0b3JzIiwKICAgICAiTmV0IGRpZmZlcmVudGlhbCBhbW91bnQiLCBTQ0hFTUFfNSksCgog"
    "ICAgIyAtLS0tIDE1IFN1cHBsaWVzIHUvcyA5KDUpIC0tLS0KICAgICgiMTUiLCAgIlN1cHBsaWVz"
    "IHUvcyA5KDUpIiwKICAgICAiMTUgLSBTdXBwbGllcyBVL3MgOSg1KSIsICJUb3RhbCAiLCBTQ0hF"
    "TUFfNSksCiAgICAoIjE1QShJKSIsICAiQW1lbmRlZCA5KDUpIC0gUmVnaXN0ZXJlZCBSZWNpcGll"
    "bnRzIiwKICAgICAiMTVBIChJKSAtIEFtZW5kZWQgU3VwcGxpZXMgVS9zIDkoNSkiLAogICAgICJO"
    "ZXQgZGlmZmVyZW50aWFsIGFtb3VudCIsIFNDSEVNQV81KSwKICAgICgiMTVBKElJKSIsICJBbWVu"
    "ZGVkIDkoNSkgLSBVbnJlZ2lzdGVyZWQgUmVjaXBpZW50cyIsCiAgICAgIjE1QSAoSUkpIC0gQW1l"
    "bmRlZCBTdXBwbGllcyIsCiAgICAgIk5ldCBkaWZmZXJlbnRpYWwgYW1vdW50IiwgU0NIRU1BXzUp"
    "LApdCgoKIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0gIwojICBQYXJzaW5nIGhlbHBlcnMKIyAtLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0gIwoKIyBOdW1iZXIgcmVnZXgg4oCUIGNhcHR1cmVzOgojICAgMSkgICgxMiwz"
    "NDUuNjcpICBvciAgKCAxMiwzNDUuNjcgKSAgICAgICDihpIgcGFyZW50aGVzaXNlZCBuZWdhdGl2"
    "ZSAgKEdTVFItMSBzdGFuZGFyZCkKIyAgIDIpICAtMTIsMzQ1LjY3ICAgICAgICAgICAgICAgICAg"
    "ICAgICAgICAgIOKGkiBsZWFkaW5nIG1pbnVzCiMgICAzKSAgMTIsMzQ1LjY3LSAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAgICDihpIgdHJhaWxpbmcgbWludXMgKHJhcmUpCiMgICA0KSAgMTIsMzQ1"
    "LjY3ICAgICAgICAgICAgICAgICAgICAgICAgICAgICDihpIgcG9zaXRpdmUKIyBJbmRpYW4gY29t"
    "bWEgZ3JvdXBpbmcgKGxha2gvY3JvcmU6IDEsMjMsNDUsNjc4LjkwKSBpcyBzdXBwb3J0ZWQuCk5V"
    "TV9SRSA9IHJlLmNvbXBpbGUoCiAgICByIiIiCiAgICAoP1A8cGFyZW4+XChccypcZHsxLDN9KD86"
    "LFxkezIsM30pKlwuXGR7Mn1ccypcKSkgICAgICAjICgxMiwzNDUuNjcpCiAgICB8CiAgICAoP1A8"
    "bGVhZD4tXGR7MSwzfSg/OixcZHsyLDN9KSpcLlxkezJ9KSAgICAgICAgICAgICAgICAjIC0xMiwz"
    "NDUuNjcKICAgIHwKICAgICg/UDx0cmFpbD5cZHsxLDN9KD86LFxkezIsM30pKlwuXGR7Mn0tKSAg"
    "ICAgICAgICAgICAgICMgMTIsMzQ1LjY3LQogICAgfAogICAgKD9QPHBvcz5cZHsxLDN9KD86LFxk"
    "ezIsM30pKlwuXGR7Mn0pICAgICAgICAgICAgICAgICAgIyAxMiwzNDUuNjcgIG9yICAxMjMuNDUK"
    "ICAgICIiIiwKICAgIHJlLlZFUkJPU0UsCikKV1NfUkUgID0gcmUuY29tcGlsZShyIlxzKyIpCgoK"
    "ZGVmIG5vcm0ocyk6CiAgICAiIiJDb2xsYXBzZSB3aGl0ZXNwYWNlOyBsb3dlcmNhc2UtZnJlZS4i"
    "IiIKICAgIHJldHVybiBXU19SRS5zdWIoIiAiLCBzKS5zdHJpcCgpCgoKZGVmIF90b19mbG9hdCh0"
    "b2tlbik6CiAgICAiIiJDb252ZXJ0IGEgbWF0Y2hlZCBudW1lcmljIHRva2VuIChhbnkgb2YgdGhl"
    "IDQgZm9ybWF0cyBhYm92ZSkgdG8gYSBmbG9hdC4iIiIKICAgIHQgPSB0b2tlbi5zdHJpcCgpCiAg"
    "ICBzaWduID0gMS4wCiAgICBpZiB0LnN0YXJ0c3dpdGgoIigiKSBhbmQgdC5lbmRzd2l0aCgiKSIp"
    "OgogICAgICAgIHNpZ24gPSAtMS4wCiAgICAgICAgdCA9IHRbMTotMV0uc3RyaXAoKQogICAgZWxp"
    "ZiB0LnN0YXJ0c3dpdGgoIi0iKToKICAgICAgICBzaWduID0gLTEuMAogICAgICAgIHQgPSB0WzE6"
    "XQogICAgZWxpZiB0LmVuZHN3aXRoKCItIik6CiAgICAgICAgc2lnbiA9IC0xLjAKICAgICAgICB0"
    "ID0gdFs6LTFdCiAgICByZXR1cm4gc2lnbiAqIGZsb2F0KHQucmVwbGFjZSgiLCIsICIiKSkKCgoj"
    "IFdhdGVybWFyayBsZXR0ZXJzIHRoYXQgdGhlIEdTVCBwb3J0YWwgc29tZXRpbWVzIG92ZXJsYXlz"
    "IGRpYWdvbmFsbHkgb24gdGhlIFBERiwKIyB3aGljaCBwZGZwbHVtYmVyIG9jY2FzaW9uYWxseSBp"
    "bnNlcnRzIG1pZC10ZXh0LiBTdHJpcHBpbmcgaXNvbGF0ZWQgc2luZ2xlIGxldHRlcnMKIyAobm90"
    "IHByZWNlZGVkL2ZvbGxvd2VkIGJ5IGEgd29yZCBjaGFyYWN0ZXIpIG1ha2VzIG51bWVyaWMgcGFy"
    "c2luZyByb2J1c3QuCldBVEVSTUFSS19MRVRURVJTID0gcmUuY29tcGlsZShyIig/PCFbQS1aYS16"
    "MC05XSlbRklOQUxdKD8hW0EtWmEtejAtOV0pIikKCgpkZWYgY2xlYW5fbGluZV9mb3JfbnVtYmVy"
    "cyhsaW5lKToKICAgICIiIgogICAgUmVtb3ZlIGlzb2xhdGVkIHNpbmdsZS1sZXR0ZXIgd2F0ZXJt"
    "YXJrIGZyYWdtZW50cyAoJ0YnLCAnSScsICdOJywgJ0EnLCAnTCcpCiAgICB0aGF0IHRoZSAnRklO"
    "QUwnIGRpYWdvbmFsIHdhdGVybWFyayBzb21ldGltZXMgaW5qZWN0cyBpbnRvIGEgbGluZS4KICAg"
    "IE9ubHkgcmVtb3ZlcyB0aGVtIHdoZW4gdGhleSBhcmUgY2xlYXJseSBub3QgcGFydCBvZiBhIHdv"
    "cmQuCiAgICAiIiIKICAgICMgTXVsdGlwbGUgcGFzc2VzIGJlY2F1c2UgcmVtb3ZhbCBjYW4gZXhw"
    "b3NlIG5ldyBpc29sYXRlZCBsZXR0ZXJzCiAgICBwcmV2ID0gTm9uZQogICAgY3VyID0gbGluZQog"
    "ICAgd2hpbGUgcHJldiAhPSBjdXI6CiAgICAgICAgcHJldiA9IGN1cgogICAgICAgIGN1ciA9IFdB"
    "VEVSTUFSS19MRVRURVJTLnN1YigiIiwgY3VyKQogICAgIyBDb2xsYXBzZSB3aGl0ZXNwYWNlIGlu"
    "dHJvZHVjZWQgYnkgcmVtb3ZhbAogICAgcmV0dXJuIFdTX1JFLnN1YigiICIsIGN1cikuc3RyaXAo"
    "KQoKCmRlZiBleHRyYWN0X251bWJlcnMobGluZSk6CiAgICAiIiIKICAgIFJldHVybiBsaXN0IG9m"
    "IGZsb2F0cyBmcm9tIGEgbGluZS4KICAgIEhhbmRsZXMgSW5kaWFuIGNvbW1hIGZvcm1hdCBBTkQg"
    "bmVnYXRpdmUgdmFsdWVzIHdyaXR0ZW4gYXMKICAgIHBhcmVudGhlc2VzLCBsZWFkaW5nIG1pbnVz"
    "LCBvciB0cmFpbGluZyBtaW51cy4KICAgIFN0cmlwcyBkaWFnb25hbC13YXRlcm1hcmsgc2luZ2xl"
    "LWxldHRlciBmcmFnbWVudHMgYmVmb3JlIHBhcnNpbmcuCiAgICAiIiIKICAgIGNsZWFuZWQgPSBj"
    "bGVhbl9saW5lX2Zvcl9udW1iZXJzKGxpbmUpCiAgICBvdXQgPSBbXQogICAgZm9yIG0gaW4gTlVN"
    "X1JFLmZpbmRpdGVyKGNsZWFuZWQpOgogICAgICAgIG91dC5hcHBlbmQoX3RvX2Zsb2F0KG0uZ3Jv"
    "dXAoMCkpKQogICAgcmV0dXJuIG91dAoKCmRlZiBleHRyYWN0X3JlY29yZHMobGluZSwgdmFsdWVf"
    "bWFya2VyKToKICAgICIiIgogICAgVHJ5IHRvIHB1bGwgdGhlIGludGVnZXIgcmVjb3JkIGNvdW50"
    "IG91dCBvZiBhIHZhbHVlIGxpbmUuCiAgICBGb3JtYXQgZXhhbXBsZTogJ1RvdGFsIDEgSW52b2lj"
    "ZSAxNSw0MSw4MDEuMDAgLi4uJwogICAgICAgICAgICAgICAgICAgICdOZXQgVG90YWwgKERlYml0"
    "IG5vdGVzIOKAkyBDcmVkaXQgbm90ZXMpIDAgTm90ZSAwLjAwIC4uLicKICAgIFJldHVybnMgaW50"
    "IG9yIDAgaWYgbm90IGZvdW5kIC8gdW5wYXJzZWFibGUuCiAgICBUaGUgd2F0ZXJtYXJrIHNvbWV0"
    "aW1lcyBwcmVmaXhlcyBhIGxldHRlciB0byB0aGUgZGlnaXQgKGUuZy4gIkkwIiwgIkYwIiksCiAg"
    "ICBzbyB3ZSBzdHJpcCBub24tZGlnaXRzLgogICAgIiIiCiAgICAjIFRha2UgdGhlIHBvcnRpb24g"
    "YWZ0ZXIgdGhlIG1hcmtlcgogICAgaWR4ID0gbGluZS5sb3dlcigpLmZpbmQodmFsdWVfbWFya2Vy"
    "LnN0cmlwKCkubG93ZXIoKSkKICAgIGlmIGlkeCA9PSAtMToKICAgICAgICByZXR1cm4gMAogICAg"
    "dGFpbCA9IGxpbmVbaWR4ICsgbGVuKHZhbHVlX21hcmtlcik6XS5zdHJpcCgpCiAgICAjIEZpcnN0"
    "IHdoaXRlc3BhY2UtdG9rZW4gaXMgdGhlIHJlY29yZHMgY291bnQgKHBvc3NpYmx5IHByZWZpeGVk"
    "IGJ5IHdhdGVybWFyaykKICAgIGZpcnN0X3RvayA9IHRhaWwuc3BsaXQoKVswXSBpZiB0YWlsLnNw"
    "bGl0KCkgZWxzZSAiIgogICAgZGlnaXRzID0gcmUuc3ViKHIiXEQiLCAiIiwgZmlyc3RfdG9rKQog"
    "ICAgcmV0dXJuIGludChkaWdpdHMpIGlmIGRpZ2l0cyBlbHNlIDAKCgpkZWYgZXh0cmFjdF9wZGZf"
    "dGV4dChwZGZfcGF0aCk6CiAgICAiIiJDb25jYXRlbmF0ZSB0ZXh0IGZyb20gZXZlcnkgcGFnZS4i"
    "IiIKICAgIHBhcnRzID0gW10KICAgIHdpdGggcGRmcGx1bWJlci5vcGVuKHBkZl9wYXRoKSBhcyBw"
    "ZGY6CiAgICAgICAgZm9yIHBhZ2UgaW4gcGRmLnBhZ2VzOgogICAgICAgICAgICB0ID0gcGFnZS5l"
    "eHRyYWN0X3RleHQoKSBvciAiIgogICAgICAgICAgICBwYXJ0cy5hcHBlbmQodCkKICAgIHJldHVy"
    "biAiXG4iLmpvaW4ocGFydHMpCgoKZGVmIHBhcnNlX21ldGEodGV4dCk6CiAgICAiIiJQdWxsIEdT"
    "VElOLCBwZXJpb2QsIEZZLCBsZWdhbCBuYW1lLCBBUk4sIEFSTiBkYXRlLiIiIgogICAgZ3N0aW5f"
    "bSAgICA9IHJlLnNlYXJjaChyIkdTVElOXHMrKFswLTldezJ9W0EtWjAtOV17MTN9KSIsIHRleHQp"
    "CiAgICBwZXJpb2RfbSAgID0gcmUuc2VhcmNoKHIiVGF4XHMqcGVyaW9kXHMrKFtBLVphLXpdKyki"
    "LCB0ZXh0KQogICAgZnlfbSAgICAgICA9IHJlLnNlYXJjaChyIkZpbmFuY2lhbFxzKnllYXJccyso"
    "XGR7NH0tXGR7Mn0pIiwgdGV4dCkKICAgIGxlZ2FsX20gICAgPSByZS5zZWFyY2gociJMZWdhbCBu"
    "YW1lIG9mIHRoZSByZWdpc3RlcmVkIHBlcnNvblxzKyguKykiLCB0ZXh0KQogICAgYXJuX20gICAg"
    "ICA9IHJlLnNlYXJjaChyIlwoY1wpXHMqQVJOXHMrKFxTKykiLCB0ZXh0KQogICAgYXJuX2RhdGVf"
    "bSA9IHJlLnNlYXJjaChyIlwoZFwpXHMqQVJOXHMqZGF0ZVxzKyhcUyspIiwgdGV4dCkKCiAgICBn"
    "c3RpbiA9IGdzdGluX20uZ3JvdXAoMSkgaWYgZ3N0aW5fbSBlbHNlICIiCiAgICBzdGF0ZV9jb2Rl"
    "ID0gZ3N0aW5bOjJdIGlmIGdzdGluIGVsc2UgIiIKICAgIHN0YXRlX25hbWUgPSBTVEFURV9DT0RF"
    "Uy5nZXQoc3RhdGVfY29kZSwgIlVua25vd24iKQogICAgcGVyaW9kID0gcGVyaW9kX20uZ3JvdXAo"
    "MSkgaWYgcGVyaW9kX20gZWxzZSAiIgogICAgZnkgPSBmeV9tLmdyb3VwKDEpIGlmIGZ5X20gZWxz"
    "ZSAiIgogICAgIyBJbmRpYW4gRlk6IEFwcuKAk0RlYyA9IGZpcnN0IHllYXIgKHl5MSksIEphbuKA"
    "k01hciA9IHNlY29uZCB5ZWFyICh5eTIpCiAgICB5eSA9ICIiCiAgICBpZiBmeSBhbmQgIi0iIGlu"
    "IGZ5OgogICAgICAgIHl5MSA9IGZ5LnNwbGl0KCItIilbMF1bLTI6XQogICAgICAgIHl5MiA9IGZ5"
    "LnNwbGl0KCItIilbMV0KICAgICAgICBpZiBwZXJpb2QgaW4gKCJKYW51YXJ5IiwgIkZlYnJ1YXJ5"
    "IiwgIk1hcmNoIik6CiAgICAgICAgICAgIHl5ID0geXkyCiAgICAgICAgZWxzZToKICAgICAgICAg"
    "ICAgeXkgPSB5eTEKICAgIG1vbnRoID0gZiJ7TU9OVEhfQUJCUi5nZXQocGVyaW9kLCBwZXJpb2Rb"
    "OjNdKX0te3l5fSIgaWYgcGVyaW9kIGVsc2UgIiIKCiAgICByZXR1cm4gewogICAgICAgICJnc3Rp"
    "biI6ICAgICAgZ3N0aW4sCiAgICAgICAgInN0YXRlX2NvZGUiOiBzdGF0ZV9jb2RlLAogICAgICAg"
    "ICJzdGF0ZV9uYW1lIjogc3RhdGVfbmFtZSwKICAgICAgICAibW9udGgiOiAgICAgIG1vbnRoLAog"
    "ICAgICAgICJmeSI6ICAgICAgICAgZnksCiAgICAgICAgInRheF9wZXJpb2QiOiBwZXJpb2QsCiAg"
    "ICAgICAgImxlZ2FsX25hbWUiOiBsZWdhbF9tLmdyb3VwKDEpLnN0cmlwKCkgaWYgbGVnYWxfbSBl"
    "bHNlICIiLAogICAgICAgICJhcm4iOiAgICAgICAgYXJuX20uZ3JvdXAoMSkgaWYgYXJuX20gZWxz"
    "ZSAiIiwKICAgICAgICAiYXJuX2RhdGUiOiAgIGFybl9kYXRlX20uZ3JvdXAoMSkgaWYgYXJuX2Rh"
    "dGVfbSBlbHNlICIiLAogICAgfQoKCkhFQURfUkUgPSByZS5jb21waWxlKHIiXlxzKlxkK1tBLVpd"
    "P1xzKig/OlwoW0lWWF0rXCkpP1xzKlst4oCTXSIpCgoKREFTSF9SRSA9IHJlLmNvbXBpbGUociJb"
    "XHUyMDEwXHUyMDExXHUyMDEyXHUyMDEzXHUyMDE0XHUyMDE1XC1dIikgICMgYW55IGtpbmQgb2Yg"
    "ZGFzaAoKCmRlZiBub3JtX2Rhc2hlcyhzKToKICAgICIiIlJlcGxhY2UgYWxsIGRhc2ggdmFyaWFu"
    "dHMgKGVuLWRhc2gsIGVtLWRhc2gsIGh5cGhlbi1taW51cywgZXRjKSB3aXRoIHNpbXBsZSAnLScu"
    "IiIiCiAgICByZXR1cm4gREFTSF9SRS5zdWIoIi0iLCBzKQoKCmRlZiBmaW5kX2hlYWRpbmdfaWR4"
    "KGxpbmVzLCBoZWFkaW5nX3N1YnN0ciwgc3RhcnQ9MCk6CiAgICAiIiJSZXR1cm4gaW5kZXggb2Yg"
    "Zmlyc3QgbGluZSB0aGF0IGNvbnRhaW5zIGhlYWRpbmdfc3Vic3RyIChkYXNoLSBhbmQgd3Mtbm9y"
    "bWFsaXplZCkuIiIiCiAgICBuZWVkbGUgPSBub3JtX2Rhc2hlcyhub3JtKGhlYWRpbmdfc3Vic3Ry"
    "KSkKICAgIGZvciBpIGluIHJhbmdlKHN0YXJ0LCBsZW4obGluZXMpKToKICAgICAgICBpZiBuZWVk"
    "bGUgaW4gbm9ybV9kYXNoZXMobm9ybShsaW5lc1tpXSkpOgogICAgICAgICAgICByZXR1cm4gaQog"
    "ICAgcmV0dXJuIC0xCgoKZGVmIGZpbmRfdmFsdWVfaWR4KGxpbmVzLCBzdGFydF9pZHgsIG1hcmtl"
    "ciwgbWF4X2xpbmVzPTIwKToKICAgICIiIgogICAgRmluZCBhIGRhdGEgbGluZSBhZnRlciBzdGFy"
    "dF9pZHggd2hvc2Ugbm9ybWFsaXplZCB0ZXh0IGJlZ2lucwogICAgd2l0aCAob3IgY29udGFpbnMs"
    "IGZvciBzdWItcm93IG1hcmtlcnMgbGlrZSAnLSBFWFBXUCcpIHRoZSBtYXJrZXIuCiAgICBTdG9w"
    "cyBpZiBhIG5ldyB0YWJsZSBoZWFkaW5nIGFwcGVhcnMgaW4gYmV0d2Vlbi4KCiAgICBUb2xlcmF0"
    "ZXMgd2F0ZXJtYXJrIGZyYWdtZW50cyAoc2luZ2xlLWNoYXIgJ0YnLCAnSScsICdOJywgJ0EnLCAn"
    "TCcgbGluZXMpCiAgICBhbmQgcGFnaW5hdGlvbiBub2lzZSBsaWtlICdJUCBBZGRyZXNzOicsICdE"
    "ZXNjcmlwdGlvbiBOby4gb2YuLi4nIGNvbHVtbnMKICAgIHRoYXQgbWF5IGludGVybGVhdmUgYmV0"
    "d2VlbiBoZWFkaW5nIGFuZCBkYXRhIG9uIG11bHRpLXBhZ2UgcmV0dXJucy4KICAgICIiIgogICAg"
    "bmVlZGxlID0gbm9ybV9kYXNoZXMobm9ybShtYXJrZXIpKQogICAgZm9yIGkgaW4gcmFuZ2Uoc3Rh"
    "cnRfaWR4ICsgMSwgbWluKHN0YXJ0X2lkeCArIDEgKyBtYXhfbGluZXMsIGxlbihsaW5lcykpKToK"
    "ICAgICAgICByYXcgPSBsaW5lc1tpXQogICAgICAgIGxpbmVfbm9ybSA9IG5vcm1fZGFzaGVzKG5v"
    "cm0ocmF3KSkKCiAgICAgICAgIyBTa2lwIGVtcHR5IC8gd2F0ZXJtYXJrLWZyYWdtZW50IC8gcGFn"
    "aW5hdGlvbiBub2lzZSBsaW5lcyAoY29udGludWUgc2Nhbm5pbmcpCiAgICAgICAgaWYgbm90IGxp"
    "bmVfbm9ybToKICAgICAgICAgICAgY29udGludWUKICAgICAgICBpZiBsZW4obGluZV9ub3JtKSA8"
    "PSAyOiAgICAgICAgICAgICAgICAgICAgICAgIyAnRicsICdJTicsIGV0YyDigJQgd2F0ZXJtYXJr"
    "IGRlYnJpcwogICAgICAgICAgICBjb250aW51ZQogICAgICAgIGlmIGxpbmVfbm9ybS5zdGFydHN3"
    "aXRoKCJJUCBBZGRyZXNzIik6CiAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgaWYgbGluZV9u"
    "b3JtLnN0YXJ0c3dpdGgoIkRlc2NyaXB0aW9uIE5vLiBvZiIpOiAgIyBoZWFkZXIgcm93IHJlcGVh"
    "dHMgb24gcGFnZSBicmVhawogICAgICAgICAgICBjb250aW51ZQogICAgICAgIGlmIGxpbmVfbm9y"
    "bS5zdGFydHN3aXRoKCJyZWNvcmRzIERvY3VtZW50IFR5cGUiKToKICAgICAgICAgICAgY29udGlu"
    "dWUKCiAgICAgICAgIyBTdG9wIG9ubHkgaWYgd2UgaGl0IGEgTkVXIGhlYWRpbmcgcm93IHRoYXQg"
    "ZG9lc24ndCBpdHNlbGYgbWF0Y2ggb3VyIG1hcmtlcgogICAgICAgIGlmIEhFQURfUkUubWF0Y2go"
    "cmF3KSBhbmQgbm90IGxpbmVfbm9ybS5zdGFydHN3aXRoKG5lZWRsZSk6CiAgICAgICAgICAgIHJl"
    "dHVybiAtMQoKICAgICAgICBpZiBsaW5lX25vcm0uc3RhcnRzd2l0aChuZWVkbGUpIG9yICgiICIg"
    "KyBuZWVkbGUgKyAiICIpIGluICgiICIgKyBsaW5lX25vcm0gKyAiICIpOgogICAgICAgICAgICBy"
    "ZXR1cm4gaQogICAgcmV0dXJuIC0xCgoKZGVmIHBhcnNlX3BkZihwZGZfcGF0aCk6CiAgICAiIiJS"
    "ZXR1cm4gKG1ldGFfZGljdCwgbGlzdF9vZl90YWJsZV9yb3dfZGljdHMpIG9yIHJhaXNlIG9uIGhh"
    "cmQgZmFpbHVyZS4iIiIKICAgIHRleHQgPSBleHRyYWN0X3BkZl90ZXh0KHBkZl9wYXRoKQogICAg"
    "aWYgbm90IHRleHQuc3RyaXAoKToKICAgICAgICByYWlzZSBWYWx1ZUVycm9yKCJQREYgaGFzIG5v"
    "IGV4dHJhY3RhYmxlIHRleHQgKHNjYW5uZWQgLyBpbWFnZS1vbmx5KS4iKQoKICAgIG1ldGEgPSBw"
    "YXJzZV9tZXRhKHRleHQpCiAgICBpZiBub3QgbWV0YVsiZ3N0aW4iXToKICAgICAgICByYWlzZSBW"
    "YWx1ZUVycm9yKCJDb3VsZCBub3QgbG9jYXRlIEdTVElOIGluIFBERi4iKQoKICAgIGxpbmVzID0g"
    "dGV4dC5zcGxpdCgiXG4iKQogICAgb3V0X3Jvd3MgPSBbXQoKICAgIGZvciB0YWJsZV9ubywgZGVz"
    "YywgaGVhZGluZywgbWFya2VyLCBzY2hlbWEgaW4gVEFSR0VUUzoKICAgICAgICBoX2lkeCA9IGZp"
    "bmRfaGVhZGluZ19pZHgobGluZXMsIGhlYWRpbmcpCiAgICAgICAgcm93ID0gewogICAgICAgICAg"
    "ICAidGFibGVfbm8iOiAgIHRhYmxlX25vLAogICAgICAgICAgICAiZGVzY3JpcHRpb24iOiBkZXNj"
    "LAogICAgICAgICAgICAicmVjb3JkcyI6ICAgIDAsCiAgICAgICAgICAgICJ0YXhhYmxlIjogICAg"
    "MC4wLAogICAgICAgICAgICAiaWdzdCI6ICAgICAgIDAuMCwKICAgICAgICAgICAgImNnc3QiOiAg"
    "ICAgICAwLjAsCiAgICAgICAgICAgICJzZ3N0IjogICAgICAgMC4wLAogICAgICAgICAgICAiY2Vz"
    "cyI6ICAgICAgIDAuMCwKICAgICAgICAgICAgImZvdW5kIjogICAgICBGYWxzZSwKICAgICAgICB9"
    "CiAgICAgICAgaWYgaF9pZHggPT0gLTE6CiAgICAgICAgICAgIG91dF9yb3dzLmFwcGVuZChyb3cp"
    "CiAgICAgICAgICAgIGNvbnRpbnVlCgogICAgICAgIHZfaWR4ID0gZmluZF92YWx1ZV9pZHgobGlu"
    "ZXMsIGhfaWR4LCBtYXJrZXIpCiAgICAgICAgaWYgdl9pZHggPT0gLTE6CiAgICAgICAgICAgIG91"
    "dF9yb3dzLmFwcGVuZChyb3cpCiAgICAgICAgICAgIGNvbnRpbnVlCgogICAgICAgIGxpbmUgPSBs"
    "aW5lc1t2X2lkeF0KICAgICAgICByb3dbImZvdW5kIl0gPSBUcnVlCgogICAgICAgICMgUmVjb3Jk"
    "cyAoYmVzdC1lZmZvcnQpCiAgICAgICAgdHJ5OgogICAgICAgICAgICByb3dbInJlY29yZHMiXSA9"
    "IGV4dHJhY3RfcmVjb3JkcyhsaW5lLCBtYXJrZXIpCiAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbjoK"
    "ICAgICAgICAgICAgcm93WyJyZWNvcmRzIl0gPSAwCgogICAgICAgICMgSWYgdGhpcyBpcyB0aGUg"
    "ZG9jdW1lbnRzLWNvdW50IGxpbmUgKFRhYmxlIDEzKSwgcmVjb3Jkcy1vbmx5CiAgICAgICAgaWYg"
    "bm90IHNjaGVtYToKICAgICAgICAgICAgb3V0X3Jvd3MuYXBwZW5kKHJvdykKICAgICAgICAgICAg"
    "Y29udGludWUKCiAgICAgICAgIyBQdWxsIG51bWVyaWMgdmFsdWVzIGZyb20gdGhlIGxpbmUKICAg"
    "ICAgICBudW1zID0gZXh0cmFjdF9udW1iZXJzKGxpbmUpCiAgICAgICAgZm9yIGNvbF9uYW1lLCB2"
    "YWwgaW4gemlwKHNjaGVtYSwgbnVtcyk6CiAgICAgICAgICAgIHJvd1tjb2xfbmFtZV0gPSB2YWwK"
    "CiAgICAgICAgb3V0X3Jvd3MuYXBwZW5kKHJvdykKCiAgICByZXR1cm4gbWV0YSwgb3V0X3Jvd3MK"
    "CgojIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLSAjCiMgIEV4Y2VsIHdyaXRlcgojIC0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLSAjCgpIRFJfRklMTCAgID0gUGF0dGVybkZpbGwoInNvbGlkIiwgc3RhcnRfY29sb3I9IjFG"
    "NEU3OCIpClNVQl9GSUxMICAgPSBQYXR0ZXJuRmlsbCgic29saWQiLCBzdGFydF9jb2xvcj0iMkU3"
    "NUI2IikKVE9UQUxfRklMTCA9IFBhdHRlcm5GaWxsKCJzb2xpZCIsIHN0YXJ0X2NvbG9yPSJGRkU2"
    "OTkiKQpHUk9VUF9GSUxMID0gUGF0dGVybkZpbGwoInNvbGlkIiwgc3RhcnRfY29sb3I9IkRERUJG"
    "NyIpCkVSUl9GSUxMICAgPSBQYXR0ZXJuRmlsbCgic29saWQiLCBzdGFydF9jb2xvcj0iRkZDN0NF"
    "IikKTUVUQV9GSUxMICA9IFBhdHRlcm5GaWxsKCJzb2xpZCIsIHN0YXJ0X2NvbG9yPSJGMkYyRjIi"
    "KQoKV0hJVEUgPSBGb250KG5hbWU9IkFyaWFsIiwgYm9sZD1UcnVlLCBjb2xvcj0iRkZGRkZGIiwg"
    "c2l6ZT0xMSkKQk9MRCAgPSBGb250KG5hbWU9IkFyaWFsIiwgYm9sZD1UcnVlLCBzaXplPTEwKQpS"
    "RUcgICA9IEZvbnQobmFtZT0iQXJpYWwiLCBzaXplPTEwKQpUSVRMRSA9IEZvbnQobmFtZT0iQXJp"
    "YWwiLCBib2xkPVRydWUsIHNpemU9MTQsIGNvbG9yPSJGRkZGRkYiKQoKdGhpbiA9IFNpZGUoYm9y"
    "ZGVyX3N0eWxlPSJ0aGluIiwgY29sb3I9IkI0QjRCNCIpCkJPUkRFUiA9IEJvcmRlcihsZWZ0PXRo"
    "aW4sIHJpZ2h0PXRoaW4sIHRvcD10aGluLCBib3R0b209dGhpbikKQ0VOVEVSID0gQWxpZ25tZW50"
    "KGhvcml6b250YWw9ImNlbnRlciIsIHZlcnRpY2FsPSJjZW50ZXIiLCB3cmFwX3RleHQ9VHJ1ZSkK"
    "TEVGVCAgID0gQWxpZ25tZW50KGhvcml6b250YWw9ImxlZnQiLCAgIHZlcnRpY2FsPSJjZW50ZXIi"
    "LCB3cmFwX3RleHQ9VHJ1ZSkKUklHSFQgID0gQWxpZ25tZW50KGhvcml6b250YWw9InJpZ2h0Iiwg"
    "IHZlcnRpY2FsPSJjZW50ZXIiKQoKTlVNX0ZNVCA9ICcjLCMjMC4wMDsoIywjIzAuMDApOyItIicK"
    "CgpkZWYgd3JpdGVfZXhjZWwocmV0dXJucywgb3V0cHV0X3BhdGgpOgogICAgIiIiCiAgICByZXR1"
    "cm5zOiBsaXN0IG9mIGRpY3RzOgogICAgICAgIHsgIm1ldGEiOiB7Li4ufSwgInJvd3MiOiBbIHsu"
    "Li59LCAuLi4gXSwgInNvdXJjZV9maWxlIjogIm5hbWUucGRmIiB9CiAgICAiIiIKICAgIHdiID0g"
    "V29ya2Jvb2soKQoKICAgICMgLS0tLSBTaGVldCAxOiBDb25zb2xpZGF0ZWQgLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KICAgIHdzID0gd2IuYWN0aXZlCiAgICB3cy50"
    "aXRsZSA9ICJDb25zb2xpZGF0ZWQiCgogICAgd3MubWVyZ2VfY2VsbHMoIkExOk8xIikKICAgIHdz"
    "WyJBMSJdID0gIkdTVFItMSBDT05TT0xJREFURUQg4oCUIFNUQVRFLVdJU0UgLyBNT05USC1XSVNF"
    "IFRBQkxFIFNVTU1BUlkiCiAgICB3c1siQTEiXS5mb250ID0gVElUTEUKICAgIHdzWyJBMSJdLmZp"
    "bGwgPSBIRFJfRklMTAogICAgd3NbIkExIl0uYWxpZ25tZW50ID0gQ0VOVEVSCiAgICB3cy5yb3df"
    "ZGltZW5zaW9uc1sxXS5oZWlnaHQgPSAyOAoKICAgIGhlYWRlcnMgPSBbCiAgICAgICAgIlNyIE5v"
    "IiwgIk1vbnRoIiwgIkZZIiwgIlN0YXRlIENvZGUiLCAiU3RhdGUgTmFtZSIsICJHU1RJTiIsCiAg"
    "ICAgICAgIkxlZ2FsIE5hbWUiLCAiQVJOIiwKICAgICAgICAiVGFibGUgTm8uIiwgIlRhYmxlIERl"
    "c2NyaXB0aW9uIiwKICAgICAgICAiTm8uIG9mIFJlY29yZHMiLAogICAgICAgICJUYXhhYmxlIFZh"
    "bHVlICjigrkpIiwgIklHU1QgKOKCuSkiLCAiQ0dTVCAo4oK5KSIsICJTR1NUL1VUR1NUICjigrkp"
    "IiwgIkNFU1MgKOKCuSkiCiAgICBdCiAgICBIRFJfUk9XID0gMwogICAgZm9yIGksIGggaW4gZW51"
    "bWVyYXRlKGhlYWRlcnMsIHN0YXJ0PTEpOgogICAgICAgIGMgPSB3cy5jZWxsKHJvdz1IRFJfUk9X"
    "LCBjb2x1bW49aSwgdmFsdWU9aCkKICAgICAgICBjLmZvbnQgPSBXSElURTsgYy5maWxsID0gU1VC"
    "X0ZJTEw7IGMuYWxpZ25tZW50ID0gQ0VOVEVSOyBjLmJvcmRlciA9IEJPUkRFUgogICAgd3Mucm93"
    "X2RpbWVuc2lvbnNbSERSX1JPV10uaGVpZ2h0ID0gMzYKCiAgICBzciA9IDAKICAgIHIgPSBIRFJf"
    "Uk9XICsgMQogICAgZGF0YV9zdGFydCA9IHIKCiAgICBmb3IgcmV0IGluIHJldHVybnM6CiAgICAg"
    "ICAgbSA9IHJldFsibWV0YSJdCiAgICAgICAgZm9yIHJvdyBpbiByZXRbInJvd3MiXToKICAgICAg"
    "ICAgICAgc3IgKz0gMQogICAgICAgICAgICB2YWxzID0gWwogICAgICAgICAgICAgICAgc3IsIG1b"
    "Im1vbnRoIl0sIG1bImZ5Il0sIG1bInN0YXRlX2NvZGUiXSwgbVsic3RhdGVfbmFtZSJdLCBtWyJn"
    "c3RpbiJdLAogICAgICAgICAgICAgICAgbVsibGVnYWxfbmFtZSJdLCBtWyJhcm4iXSwKICAgICAg"
    "ICAgICAgICAgIHJvd1sidGFibGVfbm8iXSwgcm93WyJkZXNjcmlwdGlvbiJdLAogICAgICAgICAg"
    "ICAgICAgcm93WyJyZWNvcmRzIl0sCiAgICAgICAgICAgICAgICByb3dbInRheGFibGUiXSwgcm93"
    "WyJpZ3N0Il0sIHJvd1siY2dzdCJdLCByb3dbInNnc3QiXSwgcm93WyJjZXNzIl0sCiAgICAgICAg"
    "ICAgIF0KICAgICAgICAgICAgZm9yIGksIHYgaW4gZW51bWVyYXRlKHZhbHMsIHN0YXJ0PTEpOgog"
    "ICAgICAgICAgICAgICAgY2VsbCA9IHdzLmNlbGwocm93PXIsIGNvbHVtbj1pLCB2YWx1ZT12KQog"
    "ICAgICAgICAgICAgICAgY2VsbC5mb250ID0gUkVHOyBjZWxsLmJvcmRlciA9IEJPUkRFUgogICAg"
    "ICAgICAgICAgICAgaWYgaSBpbiAoMSwgMiwgMywgNCwgOSwgMTEpOgogICAgICAgICAgICAgICAg"
    "ICAgIGNlbGwuYWxpZ25tZW50ID0gQ0VOVEVSCiAgICAgICAgICAgICAgICBlbGlmIGkgaW4gKDUs"
    "IDYsIDcsIDgsIDEwKToKICAgICAgICAgICAgICAgICAgICBjZWxsLmFsaWdubWVudCA9IExFRlQK"
    "ICAgICAgICAgICAgICAgIGVsc2U6CiAgICAgICAgICAgICAgICAgICAgY2VsbC5hbGlnbm1lbnQg"
    "PSBSSUdIVAogICAgICAgICAgICAgICAgICAgIGNlbGwubnVtYmVyX2Zvcm1hdCA9IE5VTV9GTVQK"
    "ICAgICAgICAgICAgciArPSAxCgogICAgZGF0YV9lbmQgPSByIC0gMQoKICAgIHdpZHRocyA9IHsi"
    "QSI6NywiQiI6OSwiQyI6OSwiRCI6NiwiRSI6MTgsIkYiOjIwLCJHIjoyNCwiSCI6MTgsCiAgICAg"
    "ICAgICAgICAgIkkiOjksIkoiOjQ2LCJLIjoxMSwiTCI6MTgsIk0iOjE2LCJOIjoxNiwiTyI6MTgs"
    "IlAiOjEyfQogICAgZm9yIGNvbCwgdyBpbiB3aWR0aHMuaXRlbXMoKToKICAgICAgICB3cy5jb2x1"
    "bW5fZGltZW5zaW9uc1tjb2xdLndpZHRoID0gdwoKICAgIHdzLmZyZWV6ZV9wYW5lcyA9ICJLNCIK"
    "ICAgIGlmIGRhdGFfZW5kID49IGRhdGFfc3RhcnQ6CiAgICAgICAgd3MuYXV0b19maWx0ZXIucmVm"
    "ID0gZiJBe0hEUl9ST1d9OlB7ZGF0YV9lbmR9IgoKICAgICMgLS0tLSBTaGVldCAyOiBTdGF0ZS1N"
    "b250aCBQaXZvdCBTdW1tYXJ5IC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tCiAgICB3czIg"
    "PSB3Yi5jcmVhdGVfc2hlZXQoIlN0YXRlLU1vbnRoIFRvdGFscyIpCiAgICB3czIubWVyZ2VfY2Vs"
    "bHMoIkExOkgxIikKICAgIHdzMlsiQTEiXSA9ICJUT1RBTCBMSUFCSUxJVFkgUEVSIFJFVFVSTiAo"
    "VGFibGVzIDTigJMxMSwgZXhjbHVkZXMgMTIvMTMpIgogICAgd3MyWyJBMSJdLmZvbnQgPSBUSVRM"
    "RTsgd3MyWyJBMSJdLmZpbGwgPSBIRFJfRklMTDsgd3MyWyJBMSJdLmFsaWdubWVudCA9IENFTlRF"
    "UgogICAgd3MyLnJvd19kaW1lbnNpb25zWzFdLmhlaWdodCA9IDI4CgogICAgaDIgPSBbIlNyIE5v"
    "IiwgIk1vbnRoIiwgIlN0YXRlIENvZGUiLCAiU3RhdGUgTmFtZSIsICJHU1RJTiIsCiAgICAgICAg"
    "ICAiVGF4YWJsZSBWYWx1ZSAo4oK5KSIsICJJR1NUICjigrkpIiwgIkNHU1QgKOKCuSkiLCAiU0dT"
    "VC9VVEdTVCAo4oK5KSIsICJDRVNTICjigrkpIl0KICAgIGZvciBpLCBoIGluIGVudW1lcmF0ZSho"
    "Miwgc3RhcnQ9MSk6CiAgICAgICAgYyA9IHdzMi5jZWxsKHJvdz0zLCBjb2x1bW49aSwgdmFsdWU9"
    "aCkKICAgICAgICBjLmZvbnQgPSBXSElURTsgYy5maWxsID0gU1VCX0ZJTEw7IGMuYWxpZ25tZW50"
    "ID0gQ0VOVEVSOyBjLmJvcmRlciA9IEJPUkRFUgogICAgd3MyLnJvd19kaW1lbnNpb25zWzNdLmhl"
    "aWdodCA9IDM2CgogICAgc3IgPSAwOyByciA9IDQKICAgIEVYQ0xVREVfVEFCTEVTID0geyIxMiIs"
    "ICIxMyJ9ICAjIEhTTiBzdW1tYXJ5IGFuZCBEb2N1bWVudHMgPSBkdXBsaWNhdGVzIG9mIHN1cHBs"
    "aWVzIGFib3ZlCiAgICAjIEFsc28gc2tpcCB0aGUgNkEvNkIgZ3JvdXAtVG90YWwgcm93cyBiZWNh"
    "dXNlIHdlIGFscmVhZHkgY291bnQgdGhlaXIgRVhQV1AvRVhQV09QLwogICAgIyBTRVpXUC9TRVpX"
    "T1Agc3ViLXJvd3MgaW5kaXZpZHVhbGx5IOKAlCBpbmNsdWRpbmcgdGhlIFRvdGFsIHdvdWxkIGRv"
    "dWJsZS1jb3VudC4KICAgIEVYQ0xVREVfREVTQ1JJUFRJT05TID0gewogICAgICAgICJFeHBvcnRz"
    "IC0gVG90YWwgKDZBKSIsCiAgICAgICAgIlNFWiAtIFRvdGFsICg2QikiLAogICAgfQogICAgZm9y"
    "IHJldCBpbiByZXR1cm5zOgogICAgICAgIG0gPSByZXRbIm1ldGEiXQogICAgICAgIHRvdGFscyA9"
    "IHsidGF4YWJsZSI6MC4wLCJpZ3N0IjowLjAsImNnc3QiOjAuMCwic2dzdCI6MC4wLCJjZXNzIjow"
    "LjB9CiAgICAgICAgZm9yIHJvdyBpbiByZXRbInJvd3MiXToKICAgICAgICAgICAgaWYgcm93WyJ0"
    "YWJsZV9ubyJdIGluIEVYQ0xVREVfVEFCTEVTOgogICAgICAgICAgICAgICAgY29udGludWUKICAg"
    "ICAgICAgICAgaWYgcm93WyJkZXNjcmlwdGlvbiJdIGluIEVYQ0xVREVfREVTQ1JJUFRJT05TOgog"
    "ICAgICAgICAgICAgICAgY29udGludWUKICAgICAgICAgICAgZm9yIGsgaW4gdG90YWxzOgogICAg"
    "ICAgICAgICAgICAgdG90YWxzW2tdICs9IHJvd1trXQogICAgICAgIHNyICs9IDEKICAgICAgICBv"
    "dXQgPSBbc3IsIG1bIm1vbnRoIl0sIG1bInN0YXRlX2NvZGUiXSwgbVsic3RhdGVfbmFtZSJdLCBt"
    "WyJnc3RpbiJdLAogICAgICAgICAgICAgICB0b3RhbHNbInRheGFibGUiXSwgdG90YWxzWyJpZ3N0"
    "Il0sIHRvdGFsc1siY2dzdCJdLCB0b3RhbHNbInNnc3QiXSwgdG90YWxzWyJjZXNzIl1dCiAgICAg"
    "ICAgZm9yIGksIHYgaW4gZW51bWVyYXRlKG91dCwgc3RhcnQ9MSk6CiAgICAgICAgICAgIGNlbGwg"
    "PSB3czIuY2VsbChyb3c9cnIsIGNvbHVtbj1pLCB2YWx1ZT12KQogICAgICAgICAgICBjZWxsLmZv"
    "bnQgPSBSRUc7IGNlbGwuYm9yZGVyID0gQk9SREVSCiAgICAgICAgICAgIGlmIGkgPD0gNToKICAg"
    "ICAgICAgICAgICAgIGNlbGwuYWxpZ25tZW50ID0gQ0VOVEVSIGlmIGkgIT0gNCBlbHNlIExFRlQK"
    "ICAgICAgICAgICAgZWxzZToKICAgICAgICAgICAgICAgIGNlbGwuYWxpZ25tZW50ID0gUklHSFQ7"
    "IGNlbGwubnVtYmVyX2Zvcm1hdCA9IE5VTV9GTVQKICAgICAgICByciArPSAxCgogICAgd3MyLmNv"
    "bHVtbl9kaW1lbnNpb25zWyJBIl0ud2lkdGggPSA3CiAgICB3czIuY29sdW1uX2RpbWVuc2lvbnNb"
    "IkIiXS53aWR0aCA9IDEwCiAgICB3czIuY29sdW1uX2RpbWVuc2lvbnNbIkMiXS53aWR0aCA9IDEx"
    "CiAgICB3czIuY29sdW1uX2RpbWVuc2lvbnNbIkQiXS53aWR0aCA9IDIyCiAgICB3czIuY29sdW1u"
    "X2RpbWVuc2lvbnNbIkUiXS53aWR0aCA9IDIyCiAgICBmb3IgY29sIGluICgiRiIsIkciLCJIIiwi"
    "SSIsIkoiKToKICAgICAgICB3czIuY29sdW1uX2RpbWVuc2lvbnNbY29sXS53aWR0aCA9IDE4CiAg"
    "ICB3czIuZnJlZXplX3BhbmVzID0gIkE0IgogICAgaWYgcnIgPiA0OgogICAgICAgIHdzMi5hdXRv"
    "X2ZpbHRlci5yZWYgPSBmIkEzOkp7cnItMX0iCgogICAgIyAtLS0tIFNoZWV0IDM6IFByb2Nlc3Np"
    "bmcgTG9nIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0KICAgIHdzMyA9"
    "IHdiLmNyZWF0ZV9zaGVldCgiUHJvY2Vzc2luZyBMb2ciKQogICAgd3MzLmFwcGVuZChbIiMiLCAi"
    "U291cmNlIEZpbGUiLCAiU3RhdHVzIiwgIkdTVElOIiwgIlN0YXRlIiwgIk1vbnRoIiwgIlRvdGFs"
    "IFJvd3MiLCAiTm90ZXMiXSkKICAgIGZvciBjIGluIHJhbmdlKDEsIDkpOgogICAgICAgIGNlbGwg"
    "PSB3czMuY2VsbChyb3c9MSwgY29sdW1uPWMpCiAgICAgICAgY2VsbC5mb250ID0gV0hJVEU7IGNl"
    "bGwuZmlsbCA9IFNVQl9GSUxMOyBjZWxsLmFsaWdubWVudCA9IENFTlRFUjsgY2VsbC5ib3JkZXIg"
    "PSBCT1JERVIKICAgIGZvciBpLCByZXQgaW4gZW51bWVyYXRlKHJldHVybnMsIHN0YXJ0PTEpOgog"
    "ICAgICAgIG0gPSByZXRbIm1ldGEiXQogICAgICAgIG5vdGVzID0gcmV0LmdldCgibm90ZXMiLCAi"
    "IikKICAgICAgICBzdGF0dXMgPSByZXQuZ2V0KCJzdGF0dXMiLCAiT0siKQogICAgICAgIG91dCA9"
    "IFtpLCByZXQuZ2V0KCJzb3VyY2VfZmlsZSIsIiIpLCBzdGF0dXMsCiAgICAgICAgICAgICAgIG0u"
    "Z2V0KCJnc3RpbiIsIiIpLCBtLmdldCgic3RhdGVfbmFtZSIsIiIpLCBtLmdldCgibW9udGgiLCIi"
    "KSwKICAgICAgICAgICAgICAgbGVuKHJldC5nZXQoInJvd3MiLFtdKSksIG5vdGVzXQogICAgICAg"
    "IGZvciBjLCB2IGluIGVudW1lcmF0ZShvdXQsIHN0YXJ0PTEpOgogICAgICAgICAgICBjZWxsID0g"
    "d3MzLmNlbGwocm93PWkrMSwgY29sdW1uPWMsIHZhbHVlPXYpCiAgICAgICAgICAgIGNlbGwuZm9u"
    "dCA9IFJFRzsgY2VsbC5ib3JkZXIgPSBCT1JERVI7IGNlbGwuYWxpZ25tZW50ID0gTEVGVAogICAg"
    "ICAgICAgICBpZiBzdGF0dXMgIT0gIk9LIjoKICAgICAgICAgICAgICAgIGNlbGwuZmlsbCA9IEVS"
    "Ul9GSUxMCiAgICBmb3IgY29sLCB3IGluIFsoIkEiLDYpLCgiQiIsNDIpLCgiQyIsMTApLCgiRCIs"
    "MjIpLCgiRSIsMjIpLCgiRiIsMTApLCgiRyIsMTEpLCgiSCIsNTApXToKICAgICAgICB3czMuY29s"
    "dW1uX2RpbWVuc2lvbnNbY29sXS53aWR0aCA9IHcKICAgIHdzMy5mcmVlemVfcGFuZXMgPSAiQTIi"
    "CgogICAgd2Iuc2F2ZShvdXRwdXRfcGF0aCkKCgojIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLSAjCiMgIE1h"
    "aW4KIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0gIwoKZGVmIHJlc29sdmVfb3V0cHV0X3BhdGgocmF3KToK"
    "ICAgICIiIgogICAgTWFrZSB0aGUgb3V0cHV0IHBhdGggZm9yZ2l2aW5nOgogICAgICAtIElmIHVz"
    "ZXIgcGFzc2VkIGFuIGV4aXN0aW5nIGRpcmVjdG9yeSAgICAgICDihpIgYXBwZW5kIGRlZmF1bHQg"
    "ZmlsZW5hbWUKICAgICAgLSBJZiBwYXRoIGhhcyBubyBleHRlbnNpb24gICAgICAgICAgICAgICAg"
    "ICAg4oaSIGFwcGVuZCAueGxzeAogICAgICAtIElmIGV4dGVuc2lvbiBpcyBub3QgLnhsc3ggICAg"
    "ICAgICAgICAgICAgICDihpIgcmVwbGFjZSB3aXRoIC54bHN4CiAgICAgIC0gSWYgcGFyZW50IGRp"
    "cmVjdG9yeSBkb2Vzbid0IGV4aXN0ICAgICAgICAgIOKGkiBjcmVhdGUgaXQKICAgIFJldHVybnMg"
    "YSBQYXRoIHBvaW50aW5nIHRvIGEgd3JpdGFibGUgLnhsc3ggZmlsZSBsb2NhdGlvbi4KICAgICIi"
    "IgogICAgcCA9IFBhdGgocmF3KQogICAgREVGQVVMVF9OQU1FID0gIkdTVFIxX0NvbnNvbGlkYXRl"
    "ZC54bHN4IgoKICAgIGlmIHAuZXhpc3RzKCkgYW5kIHAuaXNfZGlyKCk6CiAgICAgICAgcCA9IHAg"
    "LyBERUZBVUxUX05BTUUKICAgIGVsaWYgcC5zdWZmaXgubG93ZXIoKSAhPSAiLnhsc3giOgogICAg"
    "ICAgIHAgPSBwLndpdGhfc3VmZml4KCIueGxzeCIpCgogICAgIyBNYWtlIHN1cmUgcGFyZW50IGZv"
    "bGRlciBleGlzdHMKICAgIHAucGFyZW50Lm1rZGlyKHBhcmVudHM9VHJ1ZSwgZXhpc3Rfb2s9VHJ1"
    "ZSkKCiAgICAjIElmIHRoZSByZXNvbHZlZCBmaWxlIGFscmVhZHkgZXhpc3RzIGFuZCBpcyBsb2Nr"
    "ZWQgKG9wZW4gaW4gRXhjZWwpLCB3YXJuIGVhcmx5CiAgICBpZiBwLmV4aXN0cygpOgogICAgICAg"
    "IHRyeToKICAgICAgICAgICAgd2l0aCBvcGVuKHAsICJhYiIpOgogICAgICAgICAgICAgICAgcGFz"
    "cwogICAgICAgIGV4Y2VwdCBQZXJtaXNzaW9uRXJyb3I6CiAgICAgICAgICAgIHJhaXNlIFBlcm1p"
    "c3Npb25FcnJvcigKICAgICAgICAgICAgICAgIGYiT3V0cHV0IGZpbGUgaXMgbG9ja2VkIChwcm9i"
    "YWJseSBvcGVuIGluIEV4Y2VsKTpcbiAge3B9XG4iCiAgICAgICAgICAgICAgICAiQ2xvc2UgaXQg"
    "aW4gRXhjZWwgYW5kIHJlLXJ1bi4iCiAgICAgICAgICAgICkKICAgIHJldHVybiBwCgoKZGVmIHBy"
    "b2Nlc3NfcGRmcyhpbl9mb2xkZXIsIG91dF9wYXRoLCBvbl9wcm9ncmVzcz1Ob25lLCBvbl9sb2c9"
    "Tm9uZSk6CiAgICAiIiIKICAgIENvcmUgcHJvY2Vzc2luZyBwaXBlbGluZSByZXVzYWJsZSBieSBD"
    "TEkgYW5kIEdVSS4KCiAgICBDYWxsYmFja3M6CiAgICAgICAgb25fcHJvZ3Jlc3MoaWR4LCB0b3Rh"
    "bCkgICAgICAgICAg4oaSIGZpcmVzIGFmdGVyIGV2ZXJ5IFBERiAoZm9yIHByb2dyZXNzIGJhcikK"
    "ICAgICAgICBvbl9sb2cobGluZSwgbGV2ZWwpICAgICAgICAgICAgICDihpIgZmlyZXMgZm9yIGV2"
    "ZXJ5IHRleHR1YWwgdXBkYXRlOwogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg"
    "ICAgICAgIGxldmVsIGlzICJvayIgLyAiZmFpbCIgLyAiaW5mbyIKCiAgICBSZXR1cm5zOiAob2tf"
    "Y291bnQsIGZhaWxfY291bnQsIHRvdGFsX2NvdW50LCBvdXRfcGF0aCkKICAgICIiIgogICAgcGRm"
    "cyA9IHNvcnRlZChbcCBmb3IgcCBpbiBpbl9mb2xkZXIucmdsb2IoIioucGRmIildKQogICAgaWYg"
    "bm90IHBkZnM6CiAgICAgICAgcmFpc2UgRmlsZU5vdEZvdW5kRXJyb3IoZiJObyBQREZzIGZvdW5k"
    "IHVuZGVyOiB7aW5fZm9sZGVyfSIpCgogICAgaWYgb25fbG9nOgogICAgICAgIG9uX2xvZyhmIkZv"
    "dW5kIHtsZW4ocGRmcyl9IFBERiBmaWxlKHMpLiAgUHJvY2Vzc2luZy4uLiIsICJpbmZvIikKCiAg"
    "ICByZXR1cm5zID0gW10KICAgIG9rID0gMAogICAgZmFpbCA9IDAKCiAgICBmb3IgaWR4LCBwZGZf"
    "cGF0aCBpbiBlbnVtZXJhdGUocGRmcywgc3RhcnQ9MSk6CiAgICAgICAgdHJ5OgogICAgICAgICAg"
    "ICByZWwgPSBwZGZfcGF0aC5yZWxhdGl2ZV90byhpbl9mb2xkZXIpCiAgICAgICAgZXhjZXB0IFZh"
    "bHVlRXJyb3I6CiAgICAgICAgICAgIHJlbCA9IHBkZl9wYXRoLm5hbWUKCiAgICAgICAgdHJ5Ogog"
    "ICAgICAgICAgICBtZXRhLCByb3dzID0gcGFyc2VfcGRmKHBkZl9wYXRoKQogICAgICAgICAgICBy"
    "ZXR1cm5zLmFwcGVuZCh7CiAgICAgICAgICAgICAgICAibWV0YSI6IG1ldGEsICJyb3dzIjogcm93"
    "cywKICAgICAgICAgICAgICAgICJzb3VyY2VfZmlsZSI6IHN0cihyZWwpLAogICAgICAgICAgICAg"
    "ICAgInN0YXR1cyI6ICJPSyIsICJub3RlcyI6ICIiLAogICAgICAgICAgICB9KQogICAgICAgICAg"
    "ICBvayArPSAxCiAgICAgICAgICAgIGlmIG9uX2xvZzoKICAgICAgICAgICAgICAgIG9uX2xvZygK"
    "ICAgICAgICAgICAgICAgICAgICBmIlt7aWR4Oj40fS97bGVuKHBkZnMpfV0gIE9LICAgICIKICAg"
    "ICAgICAgICAgICAgICAgICBmInttZXRhWydzdGF0ZV9uYW1lJ106PDIyfSB7bWV0YVsnbW9udGgn"
    "XTo8OH0gICIKICAgICAgICAgICAgICAgICAgICBmInttZXRhWydnc3RpbiddfSAgKHtwZGZfcGF0"
    "aC5uYW1lfSkiLAogICAgICAgICAgICAgICAgICAgICJvayIsCiAgICAgICAgICAgICAgICApCiAg"
    "ICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAgICAgICAgICBmYWlsICs9IDEKICAgICAg"
    "ICAgICAgcmV0dXJucy5hcHBlbmQoewogICAgICAgICAgICAgICAgIm1ldGEiOiB7ImdzdGluIjoi"
    "Iiwic3RhdGVfY29kZSI6IiIsInN0YXRlX25hbWUiOiIiLCJtb250aCI6IiIsCiAgICAgICAgICAg"
    "ICAgICAgICAgICAgICAiZnkiOiIiLCJsZWdhbF9uYW1lIjoiIiwiYXJuIjoiIn0sCiAgICAgICAg"
    "ICAgICAgICAicm93cyI6IFtdLAogICAgICAgICAgICAgICAgInNvdXJjZV9maWxlIjogc3RyKHJl"
    "bCksCiAgICAgICAgICAgICAgICAic3RhdHVzIjogIkZBSUwiLAogICAgICAgICAgICAgICAgIm5v"
    "dGVzIjogZiJ7dHlwZShlKS5fX25hbWVfX306IHtlfSIsCiAgICAgICAgICAgIH0pCiAgICAgICAg"
    "ICAgIGlmIG9uX2xvZzoKICAgICAgICAgICAgICAgIG9uX2xvZygKICAgICAgICAgICAgICAgICAg"
    "ICBmIlt7aWR4Oj40fS97bGVuKHBkZnMpfV0gIEZBSUwgIHtwZGZfcGF0aC5uYW1lfSAg4oaSICB7"
    "ZX0iLAogICAgICAgICAgICAgICAgICAgICJmYWlsIiwKICAgICAgICAgICAgICAgICkKCiAgICAg"
    "ICAgaWYgb25fcHJvZ3Jlc3M6CiAgICAgICAgICAgIG9uX3Byb2dyZXNzKGlkeCwgbGVuKHBkZnMp"
    "KQoKICAgIGlmIG9uX2xvZzoKICAgICAgICBvbl9sb2coZiJXcml0aW5nIGNvbnNvbGlkYXRlZCBF"
    "eGNlbDogIHtvdXRfcGF0aH0iLCAiaW5mbyIpCiAgICB3cml0ZV9leGNlbChyZXR1cm5zLCBvdXRf"
    "cGF0aCkKICAgIHJldHVybiBvaywgZmFpbCwgbGVuKHBkZnMpLCBvdXRfcGF0aAoKCiMgLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tICMKIyAgR1VJICh0a2ludGVyIOKAlCBidW5kbGVkIHdpdGggUHl0aG9uIG9u"
    "IFdpbmRvd3MgLyBNYWMpCiMgLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tICMKCg=="
)

_GSTR3B_ENGINE_B64 = (
    "U1RBVEVfQ09ERVMgPSB7CiAgICAiMDEiOiAiSmFtbXUgJiBLYXNobWlyIiwgIjAyIjogIkhpbWFj"
    "aGFsIFByYWRlc2giLCAiMDMiOiAiUHVuamFiIiwKICAgICIwNCI6ICJDaGFuZGlnYXJoIiwgIjA1"
    "IjogIlV0dGFyYWtoYW5kIiwgIjA2IjogIkhhcnlhbmEiLCAiMDciOiAiRGVsaGkiLAogICAgIjA4"
    "IjogIlJhamFzdGhhbiIsICIwOSI6ICJVdHRhciBQcmFkZXNoIiwgIjEwIjogIkJpaGFyIiwgIjEx"
    "IjogIlNpa2tpbSIsCiAgICAiMTIiOiAiQXJ1bmFjaGFsIFByYWRlc2giLCAiMTMiOiAiTmFnYWxh"
    "bmQiLCAiMTQiOiAiTWFuaXB1ciIsICIxNSI6ICJNaXpvcmFtIiwKICAgICIxNiI6ICJUcmlwdXJh"
    "IiwgIjE3IjogIk1lZ2hhbGF5YSIsICIxOCI6ICJBc3NhbSIsICIxOSI6ICJXZXN0IEJlbmdhbCIs"
    "CiAgICAiMjAiOiAiSmhhcmtoYW5kIiwgIjIxIjogIk9kaXNoYSIsICIyMiI6ICJDaGhhdHRpc2dh"
    "cmgiLCAiMjMiOiAiTWFkaHlhIFByYWRlc2giLAogICAgIjI0IjogIkd1amFyYXQiLCAiMjUiOiAi"
    "RGFtYW4gJiBEaXUiLCAiMjYiOiAiRGFkcmEgJiBOYWdhciBIYXZlbGkgYW5kIERhbWFuICYgRGl1"
    "IiwKICAgICIyNyI6ICJNYWhhcmFzaHRyYSIsICIyOCI6ICJBbmRocmEgUHJhZGVzaCAoT2xkKSIs"
    "ICIyOSI6ICJLYXJuYXRha2EiLAogICAgIjMwIjogIkdvYSIsICIzMSI6ICJMYWtzaGFkd2VlcCIs"
    "ICIzMiI6ICJLZXJhbGEiLCAiMzMiOiAiVGFtaWwgTmFkdSIsCiAgICAiMzQiOiAiUHVkdWNoZXJy"
    "eSIsICIzNSI6ICJBbmRhbWFuICYgTmljb2JhciBJc2xhbmRzIiwgIjM2IjogIlRlbGFuZ2FuYSIs"
    "CiAgICAiMzciOiAiQW5kaHJhIFByYWRlc2giLCAiMzgiOiAiTGFkYWtoIiwgIjk3IjogIk90aGVy"
    "IFRlcnJpdG9yeSIsICI5OSI6ICJDZW50cmUgSnVyaXNkaWN0aW9uIiwKfQoKTU9OVEhfQUJCUiA9"
    "IHsKICAgICJKYW51YXJ5IjogIkphbiIsICJGZWJydWFyeSI6ICJGZWIiLCAiTWFyY2giOiAiTWFy"
    "IiwgIkFwcmlsIjogIkFwciIsCiAgICAiTWF5IjogICAgICJNYXkiLCAiSnVuZSI6ICAgICAiSnVu"
    "IiwgIkp1bHkiOiAgIkp1bCIsICJBdWd1c3QiOiAgICJBdWciLAogICAgIlNlcHRlbWJlciI6ICJT"
    "ZXAiLCAiT2N0b2JlciI6ICJPY3QiLCAiTm92ZW1iZXIiOiAiTm92IiwgIkRlY2VtYmVyIjogIkRl"
    "YyIsCn0KCgojIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLSAjCiMgIENlbGwgY2xlYW5lcnMgYW5kIG51bWJl"
    "ciBwYXJzaW5nCiMgLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tICMKCldTX1JFID0gcmUuY29tcGlsZShyIlxz"
    "KyIpCiMgV2F0ZXJtYXJrICJGSUxFRCIgZGlhZ29uYWxseSBvdmVybGF5cyBlYWNoIHBhZ2UuIHBk"
    "ZnBsdW1iZXIgc29tZXRpbWVzIGluamVjdHMKIyBpc29sYXRlZCBzaW5nbGUgdXBwZXJjYXNlIGxl"
    "dHRlcnMgKEYsIEksIEwsIEUsIEQpIGFzIGNlbGwgcHJlZml4ZXMgb3IgaW4gdGhlCiMgbWlkZGxl"
    "IG9mIG51bWVyaWMgdG9rZW5zLiBXZSBzdHJpcCB0aG9zZSBkZWZlbnNpdmVseS4KTEVBRElOR19M"
    "RVRURVJfUkUgPSByZS5jb21waWxlKHIiXltBLVpdXHMrIikKTEVBRElOR19MRVRURVJTX05VTV9S"
    "RSA9IHJlLmNvbXBpbGUociJeW0EtWmEtel0rKD89W1xkXChcLV0pIikKCgpkZWYgY2xlYW5fY2Vs"
    "bChzKToKICAgICIiIk5vcm1hbGl6ZSB3aGl0ZXNwYWNlLCBzdHJpcCB3YXRlcm1hcmsgc2luZ2xl"
    "LWxldHRlciBwcmVmaXhlcy4iIiIKICAgIGlmIHMgaXMgTm9uZToKICAgICAgICByZXR1cm4gIiIK"
    "ICAgIHMgPSBzdHIocykKICAgICMgQ29sbGFwc2UgYWxsIHdoaXRlc3BhY2UgKGluY2wuIG5ld2xp"
    "bmVzIGluc2lkZSBjZWxscyBmcm9tIG11bHRpLWxpbmUgd3JhcHMpCiAgICBzID0gV1NfUkUuc3Vi"
    "KCIgIiwgcykuc3RyaXAoKQogICAgIyBTdHJpcCBsZWFkaW5nICdYICcgKHNpbmdsZSB1cHBlcmNh"
    "c2UgbGV0dGVyICsgc3BhY2UpIOKAlCB3YXRlcm1hcmsgYXJ0aWZhY3QKICAgIHMgPSBMRUFESU5H"
    "X0xFVFRFUl9SRS5zdWIoIiIsIHMpCiAgICByZXR1cm4gcy5zdHJpcCgpCgoKZGVmIHBhcnNlX251"
    "bV9jZWxsKHMpOgogICAgIiIiCiAgICBDb252ZXJ0IGEgbnVtZXJpYyBjZWxsIHRvIGZsb2F0Lgog"
    "ICAgSGFuZGxlczoKICAgICAgLSBlbXB0eSAvICctJyAvIE5vbmUgICAgICAgIOKGkiAwLjAKICAg"
    "ICAgLSAnTDM2NjA5MDIuMDAnICAgICAgICAgICAgICDihpIgMzY2MDkwMi4wMCAgKHN0cmlwIHdh"
    "dGVybWFyayBsZXR0ZXIgcHJlZml4KQogICAgICAtICcyNzE5OTc2OSAuMDAnICAgICAgICAgICAg"
    "IOKGkiAyNzE5OTc2OS4wMCAoc3RyaXAgaW50ZXJuYWwgd2hpdGVzcGFjZSBmcm9tIHdyYXBwZWQg"
    "Y29sKQogICAgICAtICcxMTc3MDIuMCAwJyAgICAgICAgICAgICAgIOKGkiAxMTc3MDIuMDAgICAo"
    "c2FtZSkKICAgICAgLSAnKDUwLDAwMC4wMCknICAgICAgICAgICAgICDihpIgLTUwMDAwLjAwICAg"
    "KHBhcmVuIG5lZ2F0aXZlKQogICAgICAtICctMTIsMzQ1LjAwJyAvICcxMiwzNDUuMDAtJ+KGkiAt"
    "MTIzNDUuMDAgICAobGVhZC90cmFpbCBtaW51cykKICAgICAgLSAnMSwyMyw0NTYuNzgnIG9yICcx"
    "MjM0NTYuNzgnIChJbmRpYW4gb3IgcGxhaW4gY29tbWEgZm9ybWF0KSDihpIgMTIzNDU2Ljc4CiAg"
    "ICBBbnl0aGluZyB0aGF0IHdvbid0IHBhcnNlIGNsZWFubHkgcmV0dXJucyAwLjAuCiAgICAiIiIK"
    "ICAgIGlmIHMgaXMgTm9uZToKICAgICAgICByZXR1cm4gMC4wCiAgICBzID0gY2xlYW5fY2VsbChz"
    "KQogICAgaWYgcyA9PSAiIiBvciBzID09ICItIjoKICAgICAgICByZXR1cm4gMC4wCiAgICAjIFN0"
    "cmlwIEFMTCBpbnRlcm5hbCB3aGl0ZXNwYWNlIOKAlCBoYW5kbGVzIGZyYWdtZW50ZWQgbnVtYmVy"
    "cyBsaWtlICIxMTc3MDIuMCAwIgogICAgcyA9IHJlLnN1YihyIlxzKyIsICIiLCBzKQogICAgIyBT"
    "dHJpcCBsZWFkaW5nIG5vbi1udW1lcmljIHdhdGVybWFyayBsZXR0ZXJzIChlLmcuLCAnTDM2NjA5"
    "MDIuMDAnKQogICAgcyA9IExFQURJTkdfTEVUVEVSU19OVU1fUkUuc3ViKCIiLCBzKQogICAgaWYg"
    "cyA9PSAiIiBvciBzID09ICItIjoKICAgICAgICByZXR1cm4gMC4wCiAgICAjIERldGVjdCBzaWdu"
    "CiAgICBzaWduID0gMS4wCiAgICBpZiBzLnN0YXJ0c3dpdGgoIigiKSBhbmQgcy5lbmRzd2l0aCgi"
    "KSIpOgogICAgICAgIHNpZ24gPSAtMS4wCiAgICAgICAgcyA9IHNbMTotMV0KICAgIGVsaWYgcy5z"
    "dGFydHN3aXRoKCItIik6CiAgICAgICAgc2lnbiA9IC0xLjAKICAgICAgICBzID0gc1sxOl0KICAg"
    "IGVsaWYgcy5lbmRzd2l0aCgiLSIpOgogICAgICAgIHNpZ24gPSAtMS4wCiAgICAgICAgcyA9IHNb"
    "Oi0xXQogICAgIyBSZW1vdmUgdGhvdXNhbmRzIHNlcGFyYXRvcnMgKEluZGlhbiBvciB3ZXN0ZXJu"
    "KQogICAgcyA9IHMucmVwbGFjZSgiLCIsICIiKQogICAgdHJ5OgogICAgICAgIHJldHVybiBzaWdu"
    "ICogZmxvYXQocykKICAgIGV4Y2VwdCBWYWx1ZUVycm9yOgogICAgICAgIHJldHVybiAwLjAKCgpk"
    "ZWYgY2VsbF9zdGFydHNfd2l0aChjZWxsLCBwcmVmaXgpOgogICAgIiIiQ2FzZS1zZW5zaXRpdmUg"
    "cHJlZml4IGNoZWNrIG9uIGNsZWFuZWQgY2VsbC4iIiIKICAgIHJldHVybiBjbGVhbl9jZWxsKGNl"
    "bGwpLnN0YXJ0c3dpdGgocHJlZml4KQoKCmRlZiByb3dfdGV4dChyb3cpOgogICAgIiIiSm9pbiBh"
    "IHRhYmxlIHJvdydzIGNlbGxzIGludG8gb25lIHNwYWNlLXNlcGFyYXRlZCBzdHJpbmcgZm9yIGNv"
    "bnRlbnQgbWF0Y2hpbmcuIiIiCiAgICByZXR1cm4gIiAiLmpvaW4oY2xlYW5fY2VsbChjKSBmb3Ig"
    "YyBpbiByb3cgaWYgYyBpcyBub3QgTm9uZSkKCgpkZWYgdGFibGVfdGV4dCh0YWJsZSk6CiAgICAi"
    "IiJBbGwgdGV4dCBpbiBhIHRhYmxlLCBsb3dlcmNhc2VkIOKAlCBmb3Igc2VjdGlvbiBpZGVudGlm"
    "aWNhdGlvbi4iIiIKICAgIHJldHVybiAiICIuam9pbihyb3dfdGV4dChyKSBmb3IgciBpbiB0YWJs"
    "ZSkubG93ZXIoKQoKCiMgLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tICMKIyAgU2VjdGlvbiBpZGVudGlmaWNh"
    "dGlvbiAmIHBhcnNpbmcKIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0gIwoKZGVmIGlkZW50aWZ5X3NlY3Rp"
    "b24odGFibGUpOgogICAgIiIiSWRlbnRpZnkgd2hpY2ggR1NUUi0zQiBzZWN0aW9uIHRoaXMgZXh0"
    "cmFjdGVkIHRhYmxlIGJlbG9uZ3MgdG8uIiIiCiAgICB0ZXh0ID0gdGFibGVfdGV4dCh0YWJsZSkK"
    "CiAgICBpZiAidGF4IHBhaWQgdGhyb3VnaCBpdGMiIGluIHRleHQgb3IgKCJkZXNjcmlwdGkiIGlu"
    "IHRleHQgYW5kICJwYXlhYmxlIiBpbiB0ZXh0KToKICAgICAgICByZXR1cm4gIjYuMSIKICAgIGlm"
    "ICJzeXN0ZW0gY29tcHV0ZWQiIGluIHRleHQgYW5kICgiaW50ZXJlc3QiIGluIHRleHQgb3IgImxh"
    "dGUgZmVlIiBpbiB0ZXh0KToKICAgICAgICByZXR1cm4gIjUuMSIKICAgIGlmICJpbnRlci0gc3Rh"
    "dGUiIGluIHRleHQgb3IgImludGVyLXN0YXRlIHN1cHBsaWVzIiBpbiB0ZXh0OgogICAgICAgIGlm"
    "ICJjb21wb3NpdGlvbiBzY2hlbWUiIGluIHRleHQgb3IgIm5vbiBnc3Qgc3VwcGx5IiBpbiB0ZXh0"
    "IG9yICJuaWwgcmF0ZWQgc3VwcGx5IiBpbiB0ZXh0OgogICAgICAgICAgICByZXR1cm4gIjUiCiAg"
    "ICBpZiAiKGEpIG91dHdhcmQgdGF4YWJsZSIgaW4gdGV4dDoKICAgICAgICByZXR1cm4gIjMuMSIK"
    "ICAgIGlmICJlbGVjdHJvbmljIGNvbW1lcmNlIG9wZXJhdG9yIiBpbiB0ZXh0IGFuZCAiOSg1KSIg"
    "aW4gdGV4dDoKICAgICAgICByZXR1cm4gIjMuMS4xIgogICAgaWYgInVucmVnaXN0ZXJlZCBwZXJz"
    "b25zIiBpbiB0ZXh0IGFuZCAoImNvbXBvc2l0aW9uIHRheGFibGUiIGluIHRleHQgb3IgInVpbiBo"
    "b2xkZXJzIiBpbiB0ZXh0KToKICAgICAgICByZXR1cm4gIjMuMiIKICAgIGlmIGFueShrIGluIHRl"
    "eHQgZm9yIGsgaW4gKAogICAgICAgICJpdGMgYXZhaWxhYmxlIiwgIml0YyByZXZlcnNlZCIsICJu"
    "ZXQgaXRjIGF2YWlsYWJsZSIsCiAgICAgICAgImltcG9ydCBvZiBnb29kcyIsICJpbXBvcnQgb2Yg"
    "c2VydmljZXMiLCAiYWxsIG90aGVyIGl0YyIsCiAgICAgICAgImluZWxpZ2libGUgaXRjIiwgIml0"
    "YyByZWNsYWltZWQiLAogICAgKSk6CiAgICAgICAgcmV0dXJuICI0IgogICAgaWYgImJyZWFrdXAg"
    "b2YgdGF4IGxpYWJpbGl0eSIgaW4gdGV4dDoKICAgICAgICByZXR1cm4gImJyZWFrdXAiCiAgICAj"
    "IEJyZWFrdXAgc2VjdGlvbiBkYXRhIHJvdyBjYW4gYmUgYSBzdGFuZGFsb25lIHNpbmdsZS1yb3cg"
    "dGFibGUgb24gYSBsYXRlciBwYWdlCiAgICAjIHdpdGggZmlyc3QgY2VsbCBzaGFwZWQgbGlrZSAi"
    "SmFudWFyeSAyMDI2Ii4gRGV0ZWN0IHRoYXQgcGF0dGVybiBleHBsaWNpdGx5LgogICAgaWYgdGFi"
    "bGUgYW5kIGxlbih0YWJsZSkgPD0gMjoKICAgICAgICBmaXJzdCA9IGNsZWFuX2NlbGwodGFibGVb"
    "MF1bMF0pIGlmIHRhYmxlWzBdIGVsc2UgIiIKICAgICAgICBpZiByZS5tYXRjaChyIl5bQS1aXVth"
    "LXpdK1xzK1xkezR9JCIsIGZpcnN0KToKICAgICAgICAgICAgcmV0dXJuICJicmVha3VwIgogICAg"
    "IyBCcmVha3VwIGhlYWRlci1vbmx5IHRhYmxlIChQZXJpb2QgfCBJbnRlZ3JhdGVkIHRheCB8IENl"
    "bnRyYWwgdGF4IHwgLi4uKQogICAgaWYgInBlcmlvZCIgaW4gdGV4dCBhbmQgImludGVncmF0ZWQg"
    "dGF4IiBpbiB0ZXh0IGFuZCBsZW4odGFibGUpIDw9IDI6CiAgICAgICAgcmV0dXJuICJicmVha3Vw"
    "IgogICAgcmV0dXJuIE5vbmUKCgpkZWYgcGFyc2VfM18xKHRhYmxlKToKICAgICIiIlBhcnNlIFRh"
    "YmxlIDMuMSDigJQgNSByb3dzIChhKS0oZSkuIiIiCiAgICBtYXJrZXJzID0gWwogICAgICAgICgi"
    "KGEpIiwgICIzLjEoYSkiLCAiT3V0d2FyZCB0YXhhYmxlIHN1cHBsaWVzIChvdGhlciB0aGFuIHpl"
    "cm8vbmlsL2V4ZW1wdGVkKSIpLAogICAgICAgICgiKGIpIiwgICIzLjEoYikiLCAiT3V0d2FyZCB0"
    "YXhhYmxlIHN1cHBsaWVzICh6ZXJvIHJhdGVkKSIpLAogICAgICAgICgiKGMiLCAgICIzLjEoYyki"
    "LCAiT3RoZXIgb3V0d2FyZCBzdXBwbGllcyAobmlsIHJhdGVkLCBleGVtcHRlZCkiKSwgICMgUERG"
    "IHVzZXMgIihjICkiIHdpdGggc3BhY2UKICAgICAgICAoIihkKSIsICAiMy4xKGQpIiwgIklud2Fy"
    "ZCBzdXBwbGllcyAobGlhYmxlIHRvIHJldmVyc2UgY2hhcmdlKSIpLAogICAgICAgICgiKGUpIiwg"
    "ICIzLjEoZSkiLCAiTm9uLUdTVCBvdXR3YXJkIHN1cHBsaWVzIiksCiAgICBdCiAgICBvdXQgPSBb"
    "XQogICAgZm9yIHJvdyBpbiB0YWJsZToKICAgICAgICBpZiBub3Qgcm93OgogICAgICAgICAgICBj"
    "b250aW51ZQogICAgICAgIGZpcnN0ID0gY2xlYW5fY2VsbChyb3dbMF0pCiAgICAgICAgZm9yIHBy"
    "ZWZpeCwgc2VjdGlvbiwgZGVzYyBpbiBtYXJrZXJzOgogICAgICAgICAgICBpZiBmaXJzdC5zdGFy"
    "dHN3aXRoKHByZWZpeCk6CiAgICAgICAgICAgICAgICAjIEV4cGVjdCA1IG51bWVyaWMgY2VsbHMg"
    "YWZ0ZXIgZGVzY3JpcHRpb24KICAgICAgICAgICAgICAgIHZhbHMgPSBbcGFyc2VfbnVtX2NlbGwo"
    "cm93W2ldKSBpZiBpIDwgbGVuKHJvdykgZWxzZSAwLjAgZm9yIGkgaW4gcmFuZ2UoMSwgNildCiAg"
    "ICAgICAgICAgICAgICBvdXQuYXBwZW5kKHsKICAgICAgICAgICAgICAgICAgICAic2VjdGlvbiI6"
    "IHNlY3Rpb24sICJkZXNjcmlwdGlvbiI6IGRlc2MsCiAgICAgICAgICAgICAgICAgICAgInRheGFi"
    "bGUiOiB2YWxzWzBdLCAiaWdzdCI6IHZhbHNbMV0sCiAgICAgICAgICAgICAgICAgICAgImNnc3Qi"
    "OiB2YWxzWzJdLCAic2dzdCI6IHZhbHNbM10sICJjZXNzIjogdmFsc1s0XSwKICAgICAgICAgICAg"
    "ICAgIH0pCiAgICAgICAgICAgICAgICBicmVhawogICAgcmV0dXJuIG91dAoKCmRlZiBwYXJzZV8z"
    "XzFfMSh0YWJsZSk6CiAgICAiIiJQYXJzZSBUYWJsZSAzLjEuMSDigJQgMiByb3dzOiAoaSkgYW5k"
    "IChpaSkuIiIiCiAgICBtYXJrZXJzID0gWwogICAgICAgICgiKGkpIiwgICIzLjEuMShpKSIsICAi"
    "VGF4YWJsZSBzdXBwbGllcyBvbiB3aGljaCBFQ08gcGF5cyB0YXggdS9zIDkoNSkiKSwKICAgICAg"
    "ICAoIihpaSkiLCAiMy4xLjEoaWkpIiwgIlRheGFibGUgc3VwcGxpZXMgbWFkZSBieSBSUCB0aHJv"
    "dWdoIEVDTyB1L3MgOSg1KSIpLAogICAgXQogICAgb3V0ID0gW10KICAgIGZvciByb3cgaW4gdGFi"
    "bGU6CiAgICAgICAgaWYgbm90IHJvdzoKICAgICAgICAgICAgY29udGludWUKICAgICAgICBmaXJz"
    "dCA9IGNsZWFuX2NlbGwocm93WzBdKQogICAgICAgIGZvciBwcmVmaXgsIHNlY3Rpb24sIGRlc2Mg"
    "aW4gbWFya2VyczoKICAgICAgICAgICAgIyBOZWVkIHRvIGJlIGNhcmVmdWw6ICIoaWkpIiBtdXN0"
    "IG1hdGNoIGJlZm9yZSAiKGkpIiBiZWNhdXNlICIoaSkiIGlzIGEgcHJlZml4IG9mICIoaWkpIgog"
    "ICAgICAgICAgICAjIE9yZGVyIHRoZW0gY2FyZWZ1bGx5IGJlbG93CiAgICAgICAgICAgIHBhc3MK"
    "ICAgICAgICAjIFJlLXRlc3QgaW4gY29ycmVjdCBwcmVjZWRlbmNlIG9yZGVyCiAgICAgICAgaWYg"
    "Zmlyc3Quc3RhcnRzd2l0aCgiKGlpKSIpIG9yICIgKGlpKSAiIGluICgiICIgKyBmaXJzdCArICIg"
    "Iik6CiAgICAgICAgICAgIHZhbHMgPSBbcGFyc2VfbnVtX2NlbGwocm93W2ldKSBpZiBpIDwgbGVu"
    "KHJvdykgZWxzZSAwLjAgZm9yIGkgaW4gcmFuZ2UoMSwgNildCiAgICAgICAgICAgIG91dC5hcHBl"
    "bmQoewogICAgICAgICAgICAgICAgInNlY3Rpb24iOiAiMy4xLjEoaWkpIiwgImRlc2NyaXB0aW9u"
    "IjogIlRheGFibGUgc3VwcGxpZXMgbWFkZSBieSBSUCB0aHJvdWdoIEVDTyB1L3MgOSg1KSIsCiAg"
    "ICAgICAgICAgICAgICAidGF4YWJsZSI6IHZhbHNbMF0sICJpZ3N0IjogdmFsc1sxXSwKICAgICAg"
    "ICAgICAgICAgICJjZ3N0IjogdmFsc1syXSwgInNnc3QiOiB2YWxzWzNdLCAiY2VzcyI6IHZhbHNb"
    "NF0sCiAgICAgICAgICAgIH0pCiAgICAgICAgZWxpZiBmaXJzdC5zdGFydHN3aXRoKCIoaSkiKSBv"
    "ciAiIChpKSAiIGluICgiICIgKyBmaXJzdCArICIgIik6CiAgICAgICAgICAgIHZhbHMgPSBbcGFy"
    "c2VfbnVtX2NlbGwocm93W2ldKSBpZiBpIDwgbGVuKHJvdykgZWxzZSAwLjAgZm9yIGkgaW4gcmFu"
    "Z2UoMSwgNildCiAgICAgICAgICAgIG91dC5hcHBlbmQoewogICAgICAgICAgICAgICAgInNlY3Rp"
    "b24iOiAiMy4xLjEoaSkiLCAiZGVzY3JpcHRpb24iOiAiVGF4YWJsZSBzdXBwbGllcyBvbiB3aGlj"
    "aCBFQ08gcGF5cyB0YXggdS9zIDkoNSkiLAogICAgICAgICAgICAgICAgInRheGFibGUiOiB2YWxz"
    "WzBdLCAiaWdzdCI6IHZhbHNbMV0sCiAgICAgICAgICAgICAgICAiY2dzdCI6IHZhbHNbMl0sICJz"
    "Z3N0IjogdmFsc1szXSwgImNlc3MiOiB2YWxzWzRdLAogICAgICAgICAgICB9KQogICAgcmV0dXJu"
    "IG91dAoKCmRlZiBwYXJzZV8zXzIodGFibGUpOgogICAgIiIiUGFyc2UgVGFibGUgMy4yIOKAlCBp"
    "bnRlci1zdGF0ZSBicmVha2Rvd24uIDMgcm93cywgMiBudW1lcmljIGNvbHMgKFRheGFibGUgKyBJ"
    "R1NUKS4iIiIKICAgIG1hcmtlcnMgPSBbCiAgICAgICAgKCJ1bnJlZ2lzdGVyZWQgcGVyc29ucyIs"
    "ICAgICAgICAiMy4yIChVbnJlZ2lzdGVyZWQpIiwgICJJbnRlci1zdGF0ZSBzdXBwbGllcyAtIFVu"
    "cmVnaXN0ZXJlZCBQZXJzb25zIiksCiAgICAgICAgKCJjb21wb3NpdGlvbiB0YXhhYmxlIiwgICAg"
    "ICAgICAiMy4yIChDb21wb3NpdGlvbikiLCAgICJJbnRlci1zdGF0ZSBzdXBwbGllcyAtIENvbXBv"
    "c2l0aW9uIFRheGFibGUgUGVyc29ucyIpLAogICAgICAgICgidWluIGhvbGRlcnMiLCAgICAgICAg"
    "ICAgICAgICAgIjMuMiAoVUlOKSIsICAgICAgICAgICAiSW50ZXItc3RhdGUgc3VwcGxpZXMgLSBV"
    "SU4gaG9sZGVycyIpLAogICAgXQogICAgb3V0ID0gW10KICAgIGZvciByb3cgaW4gdGFibGU6CiAg"
    "ICAgICAgaWYgbm90IHJvdzoKICAgICAgICAgICAgY29udGludWUKICAgICAgICBmaXJzdCA9IGNs"
    "ZWFuX2NlbGwocm93WzBdKS5sb3dlcigpCiAgICAgICAgZm9yIGtleSwgc2VjdGlvbiwgZGVzYyBp"
    "biBtYXJrZXJzOgogICAgICAgICAgICBpZiBrZXkgaW4gZmlyc3Q6CiAgICAgICAgICAgICAgICB0"
    "YXhhYmxlID0gcGFyc2VfbnVtX2NlbGwocm93WzFdKSBpZiBsZW4ocm93KSA+IDEgZWxzZSAwLjAK"
    "ICAgICAgICAgICAgICAgIGlnc3QgPSBwYXJzZV9udW1fY2VsbChyb3dbMl0pIGlmIGxlbihyb3cp"
    "ID4gMiBlbHNlIDAuMAogICAgICAgICAgICAgICAgb3V0LmFwcGVuZCh7CiAgICAgICAgICAgICAg"
    "ICAgICAgInNlY3Rpb24iOiBzZWN0aW9uLCAiZGVzY3JpcHRpb24iOiBkZXNjLAogICAgICAgICAg"
    "ICAgICAgICAgICJ0YXhhYmxlIjogdGF4YWJsZSwgImlnc3QiOiBpZ3N0LAogICAgICAgICAgICAg"
    "ICAgICAgICJjZ3N0IjogMC4wLCAic2dzdCI6IDAuMCwgImNlc3MiOiAwLjAsCiAgICAgICAgICAg"
    "ICAgICB9KQogICAgICAgICAgICAgICAgYnJlYWsKICAgIHJldHVybiBvdXQKCgpkZWYgcGFyc2Vf"
    "NCh0YWJsZSk6CiAgICAiIiJQYXJzZSBUYWJsZSA0IElUQy4gNCBudW1lcmljIGNvbHMgKElHU1Qs"
    "IENHU1QsIFNHU1QsIENlc3MpIHBlciByb3cuIE5vIHRheGFibGUgdmFsdWUuIiIiCiAgICAjIDQo"
    "QSkgc3ViLXJvd3MgYXJlICgxKS0oNSkgQUZURVIgIkEuIElUQyBBdmFpbGFibGUiIGhlYWRlcgog"
    "ICAgIyA0KEIpIHN1Yi1yb3dzIGFyZSAoMSktKDIpIEFGVEVSICJCLiBJVEMgUmV2ZXJzZWQiIGhl"
    "YWRlcgogICAgIyA0KEMpIGlzICJDLiBOZXQgSVRDIGF2YWlsYWJsZSIKICAgICMgNChEKSBpcyAi"
    "KEQpIE90aGVyIERldGFpbHMiIHRoZW4gc3ViLXJvd3MgKDEpLSgyKQogICAgb3V0ID0gW10KICAg"
    "ICMgRGVmYXVsdCB0byAnQScgYmVjYXVzZSBUYWJsZSA0IGFsd2F5cyBzdGFydHMgd2l0aCBTZWN0"
    "aW9uIEEg4oCUIGltcG9ydGFudCBmb3IKICAgICMgbXVsdGktcGFnZSB0YWJsZXMgd2hlcmUgcGFn"
    "ZSAyIGJlZ2lucyBkaXJlY3RseSB3aXRoIHJvdyAiKDQpIiB3aXRob3V0IGFueQogICAgIyBwcmVj"
    "ZWRpbmcgJ0EuIElUQyBBdmFpbGFibGUnIGhlYWRlciBvbiB0aGF0IHBhZ2UuCiAgICBjdXJyZW50"
    "X3NlY3Rpb24gPSAiQSIKCiAgICBhX2Rlc2NzID0gewogICAgICAgICIoMSkiOiAoIjQoQSkoMSki"
    "LCAiSW1wb3J0IG9mIGdvb2RzIiksCiAgICAgICAgIigyKSI6ICgiNChBKSgyKSIsICJJbXBvcnQg"
    "b2Ygc2VydmljZXMiKSwKICAgICAgICAiKDMpIjogKCI0KEEpKDMpIiwgIklud2FyZCBzdXBwbGll"
    "cyBsaWFibGUgdG8gUkNNIChvdGhlciB0aGFuIDEgJiAyKSIpLAogICAgICAgICIoNCkiOiAoIjQo"
    "QSkoNCkiLCAiSW53YXJkIHN1cHBsaWVzIGZyb20gSVNEIiksCiAgICAgICAgIig1KSI6ICgiNChB"
    "KSg1KSIsICJBbGwgb3RoZXIgSVRDIiksCiAgICB9CiAgICBiX2Rlc2NzID0gewogICAgICAgICIo"
    "MSkiOiAoIjQoQikoMSkiLCAiQXMgcGVyIHJ1bGVzIDM4LCA0MiAmIDQzIGFuZCBzZWN0aW9uIDE3"
    "KDUpIiksCiAgICAgICAgIigyKSI6ICgiNChCKSgyKSIsICJPdGhlcnMiKSwKICAgIH0KICAgIGRf"
    "ZGVzY3MgPSB7CiAgICAgICAgIigxKSI6ICgiNChEKSgxKSIsICJJVEMgcmVjbGFpbWVkIHdoaWNo"
    "IHdhcyByZXZlcnNlZCB1bmRlciBUYWJsZSA0KEIpKDIpIGVhcmxpZXIiKSwKICAgICAgICAiKDIp"
    "IjogKCI0KEQpKDIpIiwgIkluZWxpZ2libGUgSVRDIHVuZGVyIHNlY3Rpb24gMTYoNCkgJiBQb1Mg"
    "cmVzdHJpY3Rpb24iKSwKICAgIH0KCiAgICBmb3Igcm93IGluIHRhYmxlOgogICAgICAgIGlmIG5v"
    "dCByb3c6CiAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgZmlyc3QgPSBjbGVhbl9jZWxsKHJv"
    "d1swXSkKICAgICAgICBmaXJzdF9sb3dlciA9IGZpcnN0Lmxvd2VyKCkKCiAgICAgICAgIyBTZWN0"
    "aW9uIGhlYWRlcnMg4oCUIHN3aXRjaCBjb250ZXh0CiAgICAgICAgaWYgZmlyc3RfbG93ZXIuc3Rh"
    "cnRzd2l0aCgiYS4gaXRjIGF2YWlsYWJsZSIpOgogICAgICAgICAgICBjdXJyZW50X3NlY3Rpb24g"
    "PSAiQSIKICAgICAgICAgICAgY29udGludWUKICAgICAgICBpZiBmaXJzdF9sb3dlci5zdGFydHN3"
    "aXRoKCJiLiBpdGMgcmV2ZXJzZWQiKToKICAgICAgICAgICAgY3VycmVudF9zZWN0aW9uID0gIkIi"
    "CiAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgaWYgZmlyc3RfbG93ZXIuc3RhcnRzd2l0aCgi"
    "Yy4gbmV0IGl0YyBhdmFpbGFibGUiKToKICAgICAgICAgICAgIyA0KEMpIGlzIGEgc2luZ2xlIGRh"
    "dGEgcm93IHdpdGggNCBudW1lcmljIHZhbHVlcwogICAgICAgICAgICB2YWxzID0gW3BhcnNlX251"
    "bV9jZWxsKHJvd1tpXSkgaWYgaSA8IGxlbihyb3cpIGVsc2UgMC4wIGZvciBpIGluIHJhbmdlKDEs"
    "IDUpXQogICAgICAgICAgICBvdXQuYXBwZW5kKHsKICAgICAgICAgICAgICAgICJzZWN0aW9uIjog"
    "IjQoQykiLCAiZGVzY3JpcHRpb24iOiAiTmV0IElUQyBhdmFpbGFibGUgKEEgLSBCKSIsCiAgICAg"
    "ICAgICAgICAgICAidGF4YWJsZSI6IE5vbmUsICJpZ3N0IjogdmFsc1swXSwKICAgICAgICAgICAg"
    "ICAgICJjZ3N0IjogdmFsc1sxXSwgInNnc3QiOiB2YWxzWzJdLCAiY2VzcyI6IHZhbHNbM10sCiAg"
    "ICAgICAgICAgIH0pCiAgICAgICAgICAgIGN1cnJlbnRfc2VjdGlvbiA9IE5vbmUKICAgICAgICAg"
    "ICAgY29udGludWUKICAgICAgICBpZiBmaXJzdF9sb3dlci5zdGFydHN3aXRoKCIoZCkgb3RoZXIg"
    "ZGV0YWlscyIpIG9yIGZpcnN0X2xvd2VyLnN0YXJ0c3dpdGgoImQuIG90aGVyIGRldGFpbHMiKToK"
    "ICAgICAgICAgICAgIyA0KEQpIGlzIGEgaGVhZGVyIHRoYXQgQUxTTyBoYXMgdG90YWwgdmFsdWVz"
    "IG9uIHRoZSBzYW1lIHJvdwogICAgICAgICAgICB2YWxzID0gW3BhcnNlX251bV9jZWxsKHJvd1tp"
    "XSkgaWYgaSA8IGxlbihyb3cpIGVsc2UgMC4wIGZvciBpIGluIHJhbmdlKDEsIDUpXQogICAgICAg"
    "ICAgICBvdXQuYXBwZW5kKHsKICAgICAgICAgICAgICAgICJzZWN0aW9uIjogIjQoRCkiLCAiZGVz"
    "Y3JpcHRpb24iOiAiT3RoZXIgRGV0YWlscyAodG90YWwpIiwKICAgICAgICAgICAgICAgICJ0YXhh"
    "YmxlIjogTm9uZSwgImlnc3QiOiB2YWxzWzBdLAogICAgICAgICAgICAgICAgImNnc3QiOiB2YWxz"
    "WzFdLCAic2dzdCI6IHZhbHNbMl0sICJjZXNzIjogdmFsc1szXSwKICAgICAgICAgICAgfSkKICAg"
    "ICAgICAgICAgY3VycmVudF9zZWN0aW9uID0gIkQiCiAgICAgICAgICAgIGNvbnRpbnVlCgogICAg"
    "ICAgICMgU3ViLXJvd3Mg4oCUIGRlcGVuZCBvbiBjdXJyZW50X3NlY3Rpb24KICAgICAgICBpZiBj"
    "dXJyZW50X3NlY3Rpb24gaW4gKCJBIiwgIkIiLCAiRCIpOgogICAgICAgICAgICAjIERldGVjdCAo"
    "MSktKDUpIG1hcmtlcnMKICAgICAgICAgICAgbWFya2VyX21hdGNoID0gTm9uZQogICAgICAgICAg"
    "ICBmb3Iga2V5IGluICgiKDEpIiwgIigyKSIsICIoMykiLCAiKDQpIiwgIig1KSIpOgogICAgICAg"
    "ICAgICAgICAgaWYgZmlyc3Quc3RhcnRzd2l0aChrZXkpOgogICAgICAgICAgICAgICAgICAgIG1h"
    "cmtlcl9tYXRjaCA9IGtleQogICAgICAgICAgICAgICAgICAgIGJyZWFrCiAgICAgICAgICAgIGlm"
    "IG1hcmtlcl9tYXRjaDoKICAgICAgICAgICAgICAgIGRlc2NfbWFwID0geyJBIjogYV9kZXNjcywg"
    "IkIiOiBiX2Rlc2NzLCAiRCI6IGRfZGVzY3N9W2N1cnJlbnRfc2VjdGlvbl0KICAgICAgICAgICAg"
    "ICAgIGlmIG1hcmtlcl9tYXRjaCBpbiBkZXNjX21hcDoKICAgICAgICAgICAgICAgICAgICBzZWN0"
    "aW9uLCBkZXNjID0gZGVzY19tYXBbbWFya2VyX21hdGNoXQogICAgICAgICAgICAgICAgICAgIHZh"
    "bHMgPSBbcGFyc2VfbnVtX2NlbGwocm93W2ldKSBpZiBpIDwgbGVuKHJvdykgZWxzZSAwLjAgZm9y"
    "IGkgaW4gcmFuZ2UoMSwgNSldCiAgICAgICAgICAgICAgICAgICAgb3V0LmFwcGVuZCh7CiAgICAg"
    "ICAgICAgICAgICAgICAgICAgICJzZWN0aW9uIjogc2VjdGlvbiwgImRlc2NyaXB0aW9uIjogZGVz"
    "YywKICAgICAgICAgICAgICAgICAgICAgICAgInRheGFibGUiOiBOb25lLCAiaWdzdCI6IHZhbHNb"
    "MF0sCiAgICAgICAgICAgICAgICAgICAgICAgICJjZ3N0IjogdmFsc1sxXSwgInNnc3QiOiB2YWxz"
    "WzJdLCAiY2VzcyI6IHZhbHNbM10sCiAgICAgICAgICAgICAgICAgICAgfSkKICAgIHJldHVybiBv"
    "dXQKCgpkZWYgcGFyc2VfNSh0YWJsZSk6CiAgICAiIiIKICAgIFBhcnNlIFRhYmxlIDUg4oCUIGV4"
    "ZW1wdC9uaWwtcmF0ZWQgYW5kIG5vbi1HU1QgaW53YXJkIHN1cHBsaWVzLgogICAgMiBjb2x1bW5z"
    "OiBJbnRlci1TdGF0ZSwgSW50cmEtU3RhdGUuCiAgICBFYWNoIFBERiByb3cgYmVjb21lcyBUV08g"
    "b3V0cHV0IHJvd3MgKEludGVyICsgSW50cmEpLgogICAgIiIiCiAgICBvdXQgPSBbXQogICAgdGFy"
    "Z2V0cyA9IFsKICAgICAgICAoImNvbXBvc2l0aW9uIHNjaGVtZSIsICAgIjUgKENvbXBvc2l0aW9u"
    "L0V4ZW1wdC9OaWwpIiwgIkNvbXBvc2l0aW9uIC8gRXhlbXB0IC8gTmlsIHJhdGVkIHN1cHBseSAo"
    "aW53YXJkKSIpLAogICAgICAgICgibm9uIGdzdCBzdXBwbHkiLCAgICAgICAiNSAoTm9uLUdTVCki"
    "LCAgICAgICAgICAgICAgICAgIk5vbi1HU1Qgc3VwcGx5IChpbndhcmQpIiksCiAgICBdCiAgICBm"
    "b3Igcm93IGluIHRhYmxlOgogICAgICAgIGlmIG5vdCByb3c6CiAgICAgICAgICAgIGNvbnRpbnVl"
    "CiAgICAgICAgZmlyc3QgPSBjbGVhbl9jZWxsKHJvd1swXSkubG93ZXIoKQogICAgICAgIGZvciBr"
    "ZXksIHNlY3Rpb24sIGRlc2NfYmFzZSBpbiB0YXJnZXRzOgogICAgICAgICAgICBpZiBrZXkgaW4g"
    "Zmlyc3Q6CiAgICAgICAgICAgICAgICBpbnRlciA9IHBhcnNlX251bV9jZWxsKHJvd1sxXSkgaWYg"
    "bGVuKHJvdykgPiAxIGVsc2UgMC4wCiAgICAgICAgICAgICAgICBpbnRyYSA9IHBhcnNlX251bV9j"
    "ZWxsKHJvd1syXSkgaWYgbGVuKHJvdykgPiAyIGVsc2UgMC4wCiAgICAgICAgICAgICAgICBvdXQu"
    "YXBwZW5kKHsKICAgICAgICAgICAgICAgICAgICAic2VjdGlvbiI6IHNlY3Rpb24sICJkZXNjcmlw"
    "dGlvbiI6IGRlc2NfYmFzZSArICIg4oCUIEludGVyLVN0YXRlIiwKICAgICAgICAgICAgICAgICAg"
    "ICAidGF4YWJsZSI6IGludGVyLCAiaWdzdCI6IDAuMCwgImNnc3QiOiAwLjAsICJzZ3N0IjogMC4w"
    "LCAiY2VzcyI6IDAuMCwKICAgICAgICAgICAgICAgIH0pCiAgICAgICAgICAgICAgICBvdXQuYXBw"
    "ZW5kKHsKICAgICAgICAgICAgICAgICAgICAic2VjdGlvbiI6IHNlY3Rpb24sICJkZXNjcmlwdGlv"
    "biI6IGRlc2NfYmFzZSArICIg4oCUIEludHJhLVN0YXRlIiwKICAgICAgICAgICAgICAgICAgICAi"
    "dGF4YWJsZSI6IGludHJhLCAiaWdzdCI6IDAuMCwgImNnc3QiOiAwLjAsICJzZ3N0IjogMC4wLCAi"
    "Y2VzcyI6IDAuMCwKICAgICAgICAgICAgICAgIH0pCiAgICAgICAgICAgICAgICBicmVhawogICAg"
    "cmV0dXJuIG91dAoKCmRlZiBwYXJzZV81XzEodGFibGUpOgogICAgIiIiUGFyc2UgVGFibGUgNS4x"
    "IOKAlCBJbnRlcmVzdCBhbmQgTGF0ZSBmZWUuIDQgdGF4IGNvbHVtbnMuIiIiCiAgICBvdXQgPSBb"
    "XQogICAgbWFya2VycyA9IFsKICAgICAgICAoInN5c3RlbSBjb21wdXRlZCIsICI1LjEgKFN5c3Rl"
    "bSBJbnRlcmVzdCkiLCAiU3lzdGVtIGNvbXB1dGVkIEludGVyZXN0IiksCiAgICAgICAgKCJpbnRl"
    "cmVzdCBwYWlkIiwgICAgIjUuMSAoSW50ZXJlc3QgUGFpZCkiLCAgICJJbnRlcmVzdCBQYWlkIiks"
    "CiAgICAgICAgKCJsYXRlIGZlZSIsICAgICAgICAgIjUuMSAoTGF0ZSBGZWUpIiwgICAgICAgICJM"
    "YXRlIGZlZSIpLAogICAgXQogICAgZm9yIHJvdyBpbiB0YWJsZToKICAgICAgICBpZiBub3Qgcm93"
    "OgogICAgICAgICAgICBjb250aW51ZQogICAgICAgIGZpcnN0ID0gY2xlYW5fY2VsbChyb3dbMF0p"
    "Lmxvd2VyKCkKICAgICAgICBmb3Iga2V5LCBzZWN0aW9uLCBkZXNjIGluIG1hcmtlcnM6CiAgICAg"
    "ICAgICAgIGlmIGtleSBpbiBmaXJzdDoKICAgICAgICAgICAgICAgIHZhbHMgPSBbcGFyc2VfbnVt"
    "X2NlbGwocm93W2ldKSBpZiBpIDwgbGVuKHJvdykgZWxzZSAwLjAgZm9yIGkgaW4gcmFuZ2UoMSwg"
    "NSldCiAgICAgICAgICAgICAgICBvdXQuYXBwZW5kKHsKICAgICAgICAgICAgICAgICAgICAic2Vj"
    "dGlvbiI6IHNlY3Rpb24sICJkZXNjcmlwdGlvbiI6IGRlc2MsCiAgICAgICAgICAgICAgICAgICAg"
    "InRheGFibGUiOiBOb25lLCAiaWdzdCI6IHZhbHNbMF0sCiAgICAgICAgICAgICAgICAgICAgImNn"
    "c3QiOiB2YWxzWzFdLCAic2dzdCI6IHZhbHNbMl0sICJjZXNzIjogdmFsc1szXSwKICAgICAgICAg"
    "ICAgICAgIH0pCiAgICAgICAgICAgICAgICBicmVhawogICAgcmV0dXJuIG91dAoKCmRlZiBwYXJz"
    "ZV82XzEodGFibGUpOgogICAgIiIiCiAgICBQYXJzZSBUYWJsZSA2LjEg4oCUIFBheW1lbnQgb2Yg"
    "dGF4LiAxMCBudW1lcmljIGNvbHVtbnMgcGVyIHRheC10eXBlIHJvdy4KICAgIFJldHVybnMgcm93"
    "cyB3aXRoIGEgZGlmZmVyZW50IHNjaGVtYSAodXNlZCBpbiBhIHNlcGFyYXRlIHNoZWV0KS4KICAg"
    "IENvbHVtbiBsYXlvdXQgKGFmdGVyIERlc2NyaXB0aW9uIGNvbCk6CiAgICAgICAwOiBUYXggcGF5"
    "YWJsZQogICAgICAgMTogQWRqdXN0bWVudCBvZiBuZWdhdGl2ZSBsaWFiaWxpdHkgKHByZXZpb3Vz"
    "IHBlcmlvZCkKICAgICAgIDI6IE5ldCBUYXggUGF5YWJsZQogICAgICAgMy02OiBUYXggcGFpZCB0"
    "aHJvdWdoIElUQyAoSUdTVCwgQ0dTVCwgU0dTVCwgQ2VzcykKICAgICAgIDc6IFRheCBwYWlkIGlu"
    "IGNhc2gKICAgICAgIDg6IEludGVyZXN0IHBhaWQgaW4gY2FzaAogICAgICAgOTogTGF0ZSBmZWUg"
    "cGFpZCBpbiBjYXNoCiAgICAiIiIKICAgIG91dCA9IFtdCiAgICBjdXJyZW50X3N1YnNlY3Rpb24g"
    "PSBOb25lICAjICdBJyBvciAnQicKCiAgICB0YXhfdHlwZXMgPSBbImludGVncmF0ZWQgdGF4Iiwg"
    "ImNlbnRyYWwgdGF4IiwgInN0YXRlL3V0IHRheCIsICJjZXNzIl0KCiAgICBmb3Igcm93IGluIHRh"
    "YmxlOgogICAgICAgIGlmIG5vdCByb3c6CiAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgZmly"
    "c3RfbG93ZXIgPSBjbGVhbl9jZWxsKHJvd1swXSkubG93ZXIoKQogICAgICAgIGpvaW5lZF9sb3dl"
    "ciA9IHJvd190ZXh0KHJvdykubG93ZXIoKQoKICAgICAgICAjIERldGVjdCAoQSkgYW5kIChCKSBz"
    "dWItc2VjdGlvbiBoZWFkZXJzCiAgICAgICAgaWYgIihhKSIgaW4gZmlyc3RfbG93ZXIgYW5kICJv"
    "dGhlciB0aGFuIHJldmVyc2UiIGluIGpvaW5lZF9sb3dlcjoKICAgICAgICAgICAgY3VycmVudF9z"
    "dWJzZWN0aW9uID0gIkEiCiAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAgaWYgIihiKSIgaW4g"
    "Zmlyc3RfbG93ZXIgYW5kICgicmV2ZXJzZSBjaGFyZ2UiIGluIGpvaW5lZF9sb3dlciBvciAiOSg1"
    "KSIgaW4gam9pbmVkX2xvd2VyKToKICAgICAgICAgICAgY3VycmVudF9zdWJzZWN0aW9uID0gIkIi"
    "CiAgICAgICAgICAgIGNvbnRpbnVlCgogICAgICAgIGlmIGN1cnJlbnRfc3Vic2VjdGlvbiBpcyBO"
    "b25lOgogICAgICAgICAgICBjb250aW51ZQoKICAgICAgICAjIE1hdGNoIHRheC10eXBlIHJvd3MK"
    "ICAgICAgICB0YXhfdHlwZSA9IE5vbmUKICAgICAgICBmb3IgdHQgaW4gdGF4X3R5cGVzOgogICAg"
    "ICAgICAgICBpZiB0dCBpbiBmaXJzdF9sb3dlcjoKICAgICAgICAgICAgICAgIHRheF90eXBlID0g"
    "dHQudGl0bGUoKS5yZXBsYWNlKCIvVXQiLCAiL1VUIikKICAgICAgICAgICAgICAgIGJyZWFrCiAg"
    "ICAgICAgaWYgdGF4X3R5cGUgaXMgTm9uZToKICAgICAgICAgICAgY29udGludWUKCiAgICAgICAg"
    "IyBFeHRyYWN0IDEwIG51bWVyaWMgdmFsdWVzIGZyb20gY29sdW1ucyAxLi4xMAogICAgICAgIG51"
    "bXMgPSBbcGFyc2VfbnVtX2NlbGwocm93W2ldKSBpZiBpIDwgbGVuKHJvdykgZWxzZSAwLjAgZm9y"
    "IGkgaW4gcmFuZ2UoMSwgMTEpXQogICAgICAgIG91dC5hcHBlbmQoewogICAgICAgICAgICAic3Vi"
    "c2VjdGlvbiI6IGYiNi4xKHtjdXJyZW50X3N1YnNlY3Rpb259KSIsCiAgICAgICAgICAgICJzdWJz"
    "ZWN0aW9uX2Rlc2MiOiAiT3RoZXIgdGhhbiByZXZlcnNlIGNoYXJnZSIgaWYgY3VycmVudF9zdWJz"
    "ZWN0aW9uID09ICJBIgogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgZWxzZSAiUmV2ZXJz"
    "ZSBjaGFyZ2UgYW5kIHN1cHBsaWVzIHUvcyA5KDUpIiwKICAgICAgICAgICAgInRheF90eXBlIjog"
    "dGF4X3R5cGUsCiAgICAgICAgICAgICJ0YXhfcGF5YWJsZSI6ICAgbnVtc1swXSwKICAgICAgICAg"
    "ICAgImFkanVzdG1lbnQiOiAgICBudW1zWzFdLAogICAgICAgICAgICAibmV0X3RheF9wYXlhYmxl"
    "IjogbnVtc1syXSwKICAgICAgICAgICAgIml0Y19pZ3N0IjogICAgICBudW1zWzNdLAogICAgICAg"
    "ICAgICAiaXRjX2Nnc3QiOiAgICAgIG51bXNbNF0sCiAgICAgICAgICAgICJpdGNfc2dzdCI6ICAg"
    "ICAgbnVtc1s1XSwKICAgICAgICAgICAgIml0Y19jZXNzIjogICAgICBudW1zWzZdLAogICAgICAg"
    "ICAgICAidGF4X2luX2Nhc2giOiAgIG51bXNbN10sCiAgICAgICAgICAgICJpbnRlcmVzdF9jYXNo"
    "IjogbnVtc1s4XSwKICAgICAgICAgICAgImxhdGVfZmVlX2Nhc2giOiBudW1zWzldLAogICAgICAg"
    "IH0pCiAgICByZXR1cm4gb3V0CgoKZGVmIHBhcnNlX2JyZWFrdXAodGFibGUsIGhlYWRlcl90YWJs"
    "ZT1Ob25lKToKICAgICIiIgogICAgUGFyc2UgJ0JyZWFrdXAgb2YgdGF4IGxpYWJpbGl0eSBkZWNs"
    "YXJlZCcgKHBlcmlvZC13aXNlKS4KICAgIFNpbmdsZSBkYXRhIHJvdzogPHBlcmlvZD4gPGlnc3Q+"
    "IDxjZ3N0PiA8c2dzdD4gPGNlc3M+CiAgICAiIiIKICAgIG91dCA9IFtdCiAgICBmb3Igcm93IGlu"
    "IHRhYmxlOgogICAgICAgIGlmIG5vdCByb3c6CiAgICAgICAgICAgIGNvbnRpbnVlCiAgICAgICAg"
    "Zmlyc3QgPSBjbGVhbl9jZWxsKHJvd1swXSkKICAgICAgICAjIFNraXAgaGVhZGVyIHJvdwogICAg"
    "ICAgIGlmIGZpcnN0Lmxvd2VyKCkgPT0gInBlcmlvZCI6CiAgICAgICAgICAgIGNvbnRpbnVlCiAg"
    "ICAgICAgIyBEYXRhIHJvdyDigJQgaGFzIGEgbW9udGgveWVhciB0ZXh0IGluIGZpcnN0IGNvbAog"
    "ICAgICAgICMgZS5nLiwgJ0phbnVhcnkgMjAyNicKICAgICAgICBpZiByZS5zZWFyY2gociJbQS1a"
    "YS16XStccytcZHs0fSIsIGZpcnN0KToKICAgICAgICAgICAgdmFscyA9IFtwYXJzZV9udW1fY2Vs"
    "bChyb3dbaV0pIGlmIGkgPCBsZW4ocm93KSBlbHNlIDAuMCBmb3IgaSBpbiByYW5nZSgxLCA1KV0K"
    "ICAgICAgICAgICAgb3V0LmFwcGVuZCh7CiAgICAgICAgICAgICAgICAic2VjdGlvbiI6ICJCcmVh"
    "a3VwIiwgImRlc2NyaXB0aW9uIjogZiJUYXggbGlhYmlsaXR5IGJyZWFrdXAg4oCUIHtmaXJzdH0i"
    "LAogICAgICAgICAgICAgICAgInRheGFibGUiOiBOb25lLCAiaWdzdCI6IHZhbHNbMF0sCiAgICAg"
    "ICAgICAgICAgICAiY2dzdCI6IHZhbHNbMV0sICJzZ3N0IjogdmFsc1syXSwgImNlc3MiOiB2YWxz"
    "WzNdLAogICAgICAgICAgICB9KQogICAgcmV0dXJuIG91dAoKCiMgLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "ICMKIyAgTWV0YSBleHRyYWN0aW9uCiMgLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tICMKCmRlZiBwYXJzZV9t"
    "ZXRhKHRleHQpOgogICAgZ3N0aW5fbSAgPSByZS5zZWFyY2gociJHU1RJTlxzK29mXHMrdGhlXHMr"
    "c3VwcGxpZXJccysoWzAtOV17Mn1bQS1aMC05XXsxM30pIiwgdGV4dCkKICAgIGlmIG5vdCBnc3Rp"
    "bl9tOgogICAgICAgIGdzdGluX20gPSByZS5zZWFyY2gociJcYihbMC05XXsyfVtBLVpdezV9XGR7"
    "NH1bQS1aXVtBLVowLTldWltBLVowLTldKVxiIiwgdGV4dCkKICAgIHBlcmlvZF9tID0gcmUuc2Vh"
    "cmNoKHIiUGVyaW9kXHMrKFtBLVphLXpdKykiLCB0ZXh0KQogICAgZnlfbSAgICAgPSByZS5zZWFy"
    "Y2gociJZZWFyXHMrKFxkezR9LVxkezJ9KSIsIHRleHQpCiAgICBsZWdhbF9tICA9IHJlLnNlYXJj"
    "aChyIkxlZ2FsXHMrbmFtZVxzK29mXHMrdGhlXHMrcmVnaXN0ZXJlZFxzK3BlcnNvblxzKyguKyki"
    "LCB0ZXh0KQogICAgYXJuX20gICAgPSByZS5zZWFyY2gociIyXChjXClcLlxzKkFSTlxzKyhcUysp"
    "IiwgdGV4dCkKICAgIGFybl9kYXRlX20gPSByZS5zZWFyY2gociIyXChkXClcLlxzKkRhdGVccytv"
    "ZlxzK0FSTlxzKyhcUyspIiwgdGV4dCkKCiAgICBnc3RpbiA9IGdzdGluX20uZ3JvdXAoMSkgaWYg"
    "Z3N0aW5fbSBlbHNlICIiCiAgICBzdGF0ZV9jb2RlID0gZ3N0aW5bOjJdIGlmIGdzdGluIGVsc2Ug"
    "IiIKICAgIHN0YXRlX25hbWUgPSBTVEFURV9DT0RFUy5nZXQoc3RhdGVfY29kZSwgIlVua25vd24i"
    "KQogICAgcGVyaW9kID0gcGVyaW9kX20uZ3JvdXAoMSkgaWYgcGVyaW9kX20gZWxzZSAiIgogICAg"
    "ZnkgPSBmeV9tLmdyb3VwKDEpIGlmIGZ5X20gZWxzZSAiIgogICAgIyBJbmRpYW4gRlk6IEFwcuKA"
    "k0RlYyA9IGZpcnN0IHllYXIsIEphbuKAk01hciA9IHNlY29uZCB5ZWFyCiAgICB5eSA9ICIiCiAg"
    "ICBpZiBmeSBhbmQgIi0iIGluIGZ5OgogICAgICAgIHl5MSA9IGZ5LnNwbGl0KCItIilbMF1bLTI6"
    "XQogICAgICAgIHl5MiA9IGZ5LnNwbGl0KCItIilbMV0KICAgICAgICB5eSA9IHl5MiBpZiBwZXJp"
    "b2QgaW4gKCJKYW51YXJ5IiwgIkZlYnJ1YXJ5IiwgIk1hcmNoIikgZWxzZSB5eTEKICAgIG1vbnRo"
    "ID0gZiJ7TU9OVEhfQUJCUi5nZXQocGVyaW9kLCBwZXJpb2RbOjNdKX0te3l5fSIgaWYgcGVyaW9k"
    "IGVsc2UgIiIKCiAgICByZXR1cm4gewogICAgICAgICJnc3RpbiI6ICAgICAgZ3N0aW4sCiAgICAg"
    "ICAgInN0YXRlX2NvZGUiOiBzdGF0ZV9jb2RlLAogICAgICAgICJzdGF0ZV9uYW1lIjogc3RhdGVf"
    "bmFtZSwKICAgICAgICAibW9udGgiOiAgICAgIG1vbnRoLAogICAgICAgICJmeSI6ICAgICAgICAg"
    "ZnksCiAgICAgICAgInRheF9wZXJpb2QiOiBwZXJpb2QsCiAgICAgICAgImxlZ2FsX25hbWUiOiBj"
    "bGVhbl9jZWxsKGxlZ2FsX20uZ3JvdXAoMSkpIGlmIGxlZ2FsX20gZWxzZSAiIiwKICAgICAgICAi"
    "YXJuIjogICAgICAgIGFybl9tLmdyb3VwKDEpIGlmIGFybl9tIGVsc2UgIiIsCiAgICAgICAgImFy"
    "bl9kYXRlIjogICBhcm5fZGF0ZV9tLmdyb3VwKDEpIGlmIGFybl9kYXRlX20gZWxzZSAiIiwKICAg"
    "IH0KCgojIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLSAjCiMgIE1haW4gcGVyLVBERiBwYXJzZXIKIyAtLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0gIwoKZGVmIHBhcnNlX3BkZihwZGZfcGF0aCk6CiAgICAiIiJSZXR1cm4g"
    "KG1ldGEsIGNvbnNvbGlkYXRlZF9yb3dzLCBwYXltZW50X3Jvd3MpLiBSYWlzZXMgb24gaGFyZCBm"
    "YWlsdXJlLiIiIgogICAgZnVsbF90ZXh0ID0gIiIKICAgIGFsbF90YWJsZXMgPSBbXQogICAgd2l0"
    "aCBwZGZwbHVtYmVyLm9wZW4ocGRmX3BhdGgpIGFzIHBkZjoKICAgICAgICBmb3IgcGFnZSBpbiBw"
    "ZGYucGFnZXM6CiAgICAgICAgICAgIHQgPSBwYWdlLmV4dHJhY3RfdGV4dCgpIG9yICIiCiAgICAg"
    "ICAgICAgIGZ1bGxfdGV4dCArPSB0ICsgIlxuIgogICAgICAgICAgICBmb3IgdGFibGUgaW4gcGFn"
    "ZS5leHRyYWN0X3RhYmxlcygpIG9yIFtdOgogICAgICAgICAgICAgICAgY2xlYW5lZCA9IFtbY2xl"
    "YW5fY2VsbChjKSBpZiBjIGVsc2UgIiIgZm9yIGMgaW4gcm93XSBmb3Igcm93IGluIHRhYmxlXQog"
    "ICAgICAgICAgICAgICAgYWxsX3RhYmxlcy5hcHBlbmQoY2xlYW5lZCkKCiAgICBpZiBub3QgZnVs"
    "bF90ZXh0LnN0cmlwKCk6CiAgICAgICAgcmFpc2UgVmFsdWVFcnJvcigiUERGIGhhcyBubyBleHRy"
    "YWN0YWJsZSB0ZXh0IChzY2FubmVkIC8gaW1hZ2Utb25seSkuIikKCiAgICBtZXRhID0gcGFyc2Vf"
    "bWV0YShmdWxsX3RleHQpCiAgICBpZiBub3QgbWV0YVsiZ3N0aW4iXToKICAgICAgICByYWlzZSBW"
    "YWx1ZUVycm9yKCJDb3VsZCBub3QgbG9jYXRlIEdTVElOIGluIFBERi4iKQoKICAgIGNvbnNvbGlk"
    "YXRlZCA9IFtdICAgIyByb3dzIGZvciB0aGUgbWFpbiBDb25zb2xpZGF0ZWQgc2hlZXQKICAgIHBh"
    "eW1lbnQgICAgICA9IFtdICAgIyByb3dzIGZvciB0aGUgNi4xIFRheCBQYXltZW50IHNoZWV0CiAg"
    "ICBzZWVuX3NlY3Rpb25zID0gc2V0KCkgICMgYXZvaWQgZHVwbGljYXRlIHBhcnNlcyBpZiB0YWJs"
    "ZSBhcHBlYXJzIHR3aWNlCgogICAgZm9yIHRhYmxlIGluIGFsbF90YWJsZXM6CiAgICAgICAgc2Vj"
    "dGlvbl90eXBlID0gaWRlbnRpZnlfc2VjdGlvbih0YWJsZSkKICAgICAgICBpZiBzZWN0aW9uX3R5"
    "cGUgaXMgTm9uZToKICAgICAgICAgICAgY29udGludWUKCiAgICAgICAgaWYgc2VjdGlvbl90eXBl"
    "ID09ICIzLjEiOgogICAgICAgICAgICBmb3IgciBpbiBwYXJzZV8zXzEodGFibGUpOgogICAgICAg"
    "ICAgICAgICAgY29uc29saWRhdGVkLmFwcGVuZChyKQogICAgICAgIGVsaWYgc2VjdGlvbl90eXBl"
    "ID09ICIzLjEuMSI6CiAgICAgICAgICAgIGZvciByIGluIHBhcnNlXzNfMV8xKHRhYmxlKToKICAg"
    "ICAgICAgICAgICAgIGNvbnNvbGlkYXRlZC5hcHBlbmQocikKICAgICAgICBlbGlmIHNlY3Rpb25f"
    "dHlwZSA9PSAiMy4yIjoKICAgICAgICAgICAgZm9yIHIgaW4gcGFyc2VfM18yKHRhYmxlKToKICAg"
    "ICAgICAgICAgICAgIGNvbnNvbGlkYXRlZC5hcHBlbmQocikKICAgICAgICBlbGlmIHNlY3Rpb25f"
    "dHlwZSA9PSAiNCI6CiAgICAgICAgICAgIGZvciByIGluIHBhcnNlXzQodGFibGUpOgogICAgICAg"
    "ICAgICAgICAga2V5ID0gKHJbInNlY3Rpb24iXSwgclsiZGVzY3JpcHRpb24iXSkKICAgICAgICAg"
    "ICAgICAgIGlmIGtleSBpbiBzZWVuX3NlY3Rpb25zOgogICAgICAgICAgICAgICAgICAgIGNvbnRp"
    "bnVlCiAgICAgICAgICAgICAgICBzZWVuX3NlY3Rpb25zLmFkZChrZXkpCiAgICAgICAgICAgICAg"
    "ICBjb25zb2xpZGF0ZWQuYXBwZW5kKHIpCiAgICAgICAgZWxpZiBzZWN0aW9uX3R5cGUgPT0gIjUi"
    "OgogICAgICAgICAgICBmb3IgciBpbiBwYXJzZV81KHRhYmxlKToKICAgICAgICAgICAgICAgIGNv"
    "bnNvbGlkYXRlZC5hcHBlbmQocikKICAgICAgICBlbGlmIHNlY3Rpb25fdHlwZSA9PSAiNS4xIjoK"
    "ICAgICAgICAgICAgZm9yIHIgaW4gcGFyc2VfNV8xKHRhYmxlKToKICAgICAgICAgICAgICAgIGNv"
    "bnNvbGlkYXRlZC5hcHBlbmQocikKICAgICAgICBlbGlmIHNlY3Rpb25fdHlwZSA9PSAiNi4xIjoK"
    "ICAgICAgICAgICAgZm9yIHIgaW4gcGFyc2VfNl8xKHRhYmxlKToKICAgICAgICAgICAgICAgIHBh"
    "eW1lbnQuYXBwZW5kKHIpCiAgICAgICAgZWxpZiBzZWN0aW9uX3R5cGUgPT0gImJyZWFrdXAiOgog"
    "ICAgICAgICAgICBmb3IgciBpbiBwYXJzZV9icmVha3VwKHRhYmxlKToKICAgICAgICAgICAgICAg"
    "IGtleSA9IChyWyJzZWN0aW9uIl0sIHJbImRlc2NyaXB0aW9uIl0pCiAgICAgICAgICAgICAgICBp"
    "ZiBrZXkgaW4gc2Vlbl9zZWN0aW9uczoKICAgICAgICAgICAgICAgICAgICBjb250aW51ZQogICAg"
    "ICAgICAgICAgICAgc2Vlbl9zZWN0aW9ucy5hZGQoa2V5KQogICAgICAgICAgICAgICAgY29uc29s"
    "aWRhdGVkLmFwcGVuZChyKQoKICAgIHJldHVybiBtZXRhLCBjb25zb2xpZGF0ZWQsIHBheW1lbnQK"
    "CgojIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLSAjCiMgIEV4Y2VsIHdyaXRlcgojIC0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLSAjCgpIRFJfRklMTCAgID0gUGF0dGVybkZpbGwoInNvbGlkIiwgc3RhcnRfY29sb3I9IjFG"
    "NEU3OCIpClNVQl9GSUxMICAgPSBQYXR0ZXJuRmlsbCgic29saWQiLCBzdGFydF9jb2xvcj0iMkU3"
    "NUI2IikKVE9UQUxfRklMTCA9IFBhdHRlcm5GaWxsKCJzb2xpZCIsIHN0YXJ0X2NvbG9yPSJGRkU2"
    "OTkiKQpHUk9VUF9GSUxMID0gUGF0dGVybkZpbGwoInNvbGlkIiwgc3RhcnRfY29sb3I9IkRERUJG"
    "NyIpCkVSUl9GSUxMICAgPSBQYXR0ZXJuRmlsbCgic29saWQiLCBzdGFydF9jb2xvcj0iRkZDN0NF"
    "IikKCldISVRFID0gRm9udChuYW1lPSJBcmlhbCIsIGJvbGQ9VHJ1ZSwgY29sb3I9IkZGRkZGRiIs"
    "IHNpemU9MTEpCkJPTEQgID0gRm9udChuYW1lPSJBcmlhbCIsIGJvbGQ9VHJ1ZSwgc2l6ZT0xMCkK"
    "UkVHICAgPSBGb250KG5hbWU9IkFyaWFsIiwgc2l6ZT0xMCkKVElUTEUgPSBGb250KG5hbWU9IkFy"
    "aWFsIiwgYm9sZD1UcnVlLCBzaXplPTE0LCBjb2xvcj0iRkZGRkZGIikKCnRoaW4gPSBTaWRlKGJv"
    "cmRlcl9zdHlsZT0idGhpbiIsIGNvbG9yPSJCNEI0QjQiKQpCT1JERVIgPSBCb3JkZXIobGVmdD10"
    "aGluLCByaWdodD10aGluLCB0b3A9dGhpbiwgYm90dG9tPXRoaW4pCkNFTlRFUiA9IEFsaWdubWVu"
    "dChob3Jpem9udGFsPSJjZW50ZXIiLCB2ZXJ0aWNhbD0iY2VudGVyIiwgd3JhcF90ZXh0PVRydWUp"
    "CkxFRlQgICA9IEFsaWdubWVudChob3Jpem9udGFsPSJsZWZ0IiwgICB2ZXJ0aWNhbD0iY2VudGVy"
    "Iiwgd3JhcF90ZXh0PVRydWUpClJJR0hUICA9IEFsaWdubWVudChob3Jpem9udGFsPSJyaWdodCIs"
    "ICB2ZXJ0aWNhbD0iY2VudGVyIikKCk5VTV9GTVQgPSAnIywjIzAuMDA7KCMsIyMwLjAwKTsiLSIn"
    "CgoKZGVmIF9zdHlsZV9oZWFkZXIod3MsIHJvdywgbl9jb2xzKToKICAgIGZvciBjIGluIHJhbmdl"
    "KDEsIG5fY29scyArIDEpOgogICAgICAgIGNlbGwgPSB3cy5jZWxsKHJvdz1yb3csIGNvbHVtbj1j"
    "KQogICAgICAgIGNlbGwuZm9udCA9IFdISVRFOyBjZWxsLmZpbGwgPSBTVUJfRklMTAogICAgICAg"
    "IGNlbGwuYWxpZ25tZW50ID0gQ0VOVEVSOyBjZWxsLmJvcmRlciA9IEJPUkRFUgoKCmRlZiB3cml0"
    "ZV9leGNlbChyZXR1cm5zLCBvdXRwdXRfcGF0aCk6CiAgICAiIiIKICAgIHJldHVybnM6IGxpc3Qg"
    "b2YgZGljdHM6CiAgICAgICAgeyAibWV0YSI6IHsuLi59LCAiY29uc29saWRhdGVkIjogWy4uLl0s"
    "ICJwYXltZW50IjogWy4uLl0sCiAgICAgICAgICAic291cmNlX2ZpbGUiOiAiLi4uIiwgInN0YXR1"
    "cyI6ICJPSyJ8IkZBSUwiLCAibm90ZXMiOiAiLi4uIiB9CiAgICAiIiIKICAgIHdiID0gV29ya2Jv"
    "b2soKQoKICAgICMgPT09PT09PT09PSBTaGVldCAxOiBDb25zb2xpZGF0ZWQgPT09PT09PT09PT09"
    "PT09PT09PT09PT09PT09CiAgICB3cyA9IHdiLmFjdGl2ZQogICAgd3MudGl0bGUgPSAiQ29uc29s"
    "aWRhdGVkIgoKICAgIHdzLm1lcmdlX2NlbGxzKCJBMTpOMSIpCiAgICB3c1siQTEiXSA9ICJHU1RS"
    "LTNCIENPTlNPTElEQVRFRCDigJQgU1RBVEUtV0lTRSAvIE1PTlRILVdJU0UgLyBTRUNUSU9OLVdJ"
    "U0UiCiAgICB3c1siQTEiXS5mb250ID0gVElUTEU7IHdzWyJBMSJdLmZpbGwgPSBIRFJfRklMTDsg"
    "d3NbIkExIl0uYWxpZ25tZW50ID0gQ0VOVEVSCiAgICB3cy5yb3dfZGltZW5zaW9uc1sxXS5oZWln"
    "aHQgPSAyOAoKICAgIGhlYWRlcnMgPSBbCiAgICAgICAgIlNyIE5vIiwgIk1vbnRoIiwgIkZZIiwg"
    "IlN0YXRlIENvZGUiLCAiU3RhdGUgTmFtZSIsICJHU1RJTiIsCiAgICAgICAgIkxlZ2FsIE5hbWUi"
    "LCAiQVJOIiwKICAgICAgICAiU2VjdGlvbiIsICJEZXNjcmlwdGlvbiIsCiAgICAgICAgIlRheGFi"
    "bGUgVmFsdWUgKOKCuSkiLCAiSUdTVCAo4oK5KSIsICJDR1NUICjigrkpIiwgIlNHU1QvVVRHU1Qg"
    "KOKCuSkiLCAiQ0VTUyAo4oK5KSIKICAgIF0KICAgIEhEUiA9IDMKICAgIGZvciBpLCBoIGluIGVu"
    "dW1lcmF0ZShoZWFkZXJzLCBzdGFydD0xKToKICAgICAgICBjID0gd3MuY2VsbChyb3c9SERSLCBj"
    "b2x1bW49aSwgdmFsdWU9aCkKICAgICAgICBjLmZvbnQgPSBXSElURTsgYy5maWxsID0gU1VCX0ZJ"
    "TEw7IGMuYWxpZ25tZW50ID0gQ0VOVEVSOyBjLmJvcmRlciA9IEJPUkRFUgogICAgd3Mucm93X2Rp"
    "bWVuc2lvbnNbSERSXS5oZWlnaHQgPSAzNgoKICAgIHNyID0gMAogICAgciA9IEhEUiArIDEKICAg"
    "IGRhdGFfc3RhcnQgPSByCiAgICBmb3IgcmV0IGluIHJldHVybnM6CiAgICAgICAgbSA9IHJldFsi"
    "bWV0YSJdCiAgICAgICAgZm9yIHJvdyBpbiByZXQuZ2V0KCJjb25zb2xpZGF0ZWQiLCBbXSk6CiAg"
    "ICAgICAgICAgIHNyICs9IDEKICAgICAgICAgICAgdmFscyA9IFsKICAgICAgICAgICAgICAgIHNy"
    "LCBtWyJtb250aCJdLCBtWyJmeSJdLCBtWyJzdGF0ZV9jb2RlIl0sIG1bInN0YXRlX25hbWUiXSwg"
    "bVsiZ3N0aW4iXSwKICAgICAgICAgICAgICAgIG1bImxlZ2FsX25hbWUiXSwgbVsiYXJuIl0sCiAg"
    "ICAgICAgICAgICAgICByb3dbInNlY3Rpb24iXSwgcm93WyJkZXNjcmlwdGlvbiJdLAogICAgICAg"
    "ICAgICAgICAgcm93LmdldCgidGF4YWJsZSIpLCByb3cuZ2V0KCJpZ3N0IiwgMC4wKSwgcm93Lmdl"
    "dCgiY2dzdCIsIDAuMCksCiAgICAgICAgICAgICAgICByb3cuZ2V0KCJzZ3N0IiwgMC4wKSwgcm93"
    "LmdldCgiY2VzcyIsIDAuMCksCiAgICAgICAgICAgIF0KICAgICAgICAgICAgZm9yIGksIHYgaW4g"
    "ZW51bWVyYXRlKHZhbHMsIHN0YXJ0PTEpOgogICAgICAgICAgICAgICAgY2VsbCA9IHdzLmNlbGwo"
    "cm93PXIsIGNvbHVtbj1pLCB2YWx1ZT12KQogICAgICAgICAgICAgICAgY2VsbC5mb250ID0gUkVH"
    "OyBjZWxsLmJvcmRlciA9IEJPUkRFUgogICAgICAgICAgICAgICAgaWYgaSBpbiAoMSwgMiwgMywg"
    "NCwgOSk6CiAgICAgICAgICAgICAgICAgICAgY2VsbC5hbGlnbm1lbnQgPSBDRU5URVIKICAgICAg"
    "ICAgICAgICAgIGVsaWYgaSBpbiAoNSwgNiwgNywgOCwgMTApOgogICAgICAgICAgICAgICAgICAg"
    "IGNlbGwuYWxpZ25tZW50ID0gTEVGVAogICAgICAgICAgICAgICAgZWxzZToKICAgICAgICAgICAg"
    "ICAgICAgICBjZWxsLmFsaWdubWVudCA9IFJJR0hUCiAgICAgICAgICAgICAgICAgICAgY2VsbC5u"
    "dW1iZXJfZm9ybWF0ID0gTlVNX0ZNVAogICAgICAgICAgICByICs9IDEKICAgIGRhdGFfZW5kID0g"
    "ciAtIDEKCiAgICB3aWR0aHMgPSB7IkEiOjcsIkIiOjksIkMiOjksIkQiOjYsIkUiOjE4LCJGIjoy"
    "MCwiRyI6MjIsIkgiOjE4LAogICAgICAgICAgICAgICJJIjoxNCwiSiI6NTIsIksiOjE4LCJMIjox"
    "OCwiTSI6MTYsIk4iOjE2LCJPIjoxOH0KICAgIGZvciBjb2wsIHcgaW4gd2lkdGhzLml0ZW1zKCk6"
    "CiAgICAgICAgd3MuY29sdW1uX2RpbWVuc2lvbnNbY29sXS53aWR0aCA9IHcKICAgIHdzLmZyZWV6"
    "ZV9wYW5lcyA9ICJLNCIKICAgIGlmIGRhdGFfZW5kID49IGRhdGFfc3RhcnQ6CiAgICAgICAgd3Mu"
    "YXV0b19maWx0ZXIucmVmID0gZiJBe0hEUn06T3tkYXRhX2VuZH0iCgogICAgIyA9PT09PT09PT09"
    "IFNoZWV0IDI6IFRheCBQYXltZW50ICg2LjEpID09PT09PT09PT09PT09PT09PT09PT0KICAgIHdz"
    "MiA9IHdiLmNyZWF0ZV9zaGVldCgiVGF4IFBheW1lbnQgKDYuMSkiKQogICAgd3MyLm1lcmdlX2Nl"
    "bGxzKCJBMTpRMSIpCiAgICB3czJbIkExIl0gPSAiR1NUUi0zQiAgVEFCTEUgNi4xICDigJQgUEFZ"
    "TUVOVCBPRiBUQVgiCiAgICB3czJbIkExIl0uZm9udCA9IFRJVExFOyB3czJbIkExIl0uZmlsbCA9"
    "IEhEUl9GSUxMOyB3czJbIkExIl0uYWxpZ25tZW50ID0gQ0VOVEVSCiAgICB3czIucm93X2RpbWVu"
    "c2lvbnNbMV0uaGVpZ2h0ID0gMjgKCiAgICBoMiA9IFsKICAgICAgICAiU3IgTm8iLCAiTW9udGgi"
    "LCAiRlkiLCAiU3RhdGUgQ29kZSIsICJTdGF0ZSBOYW1lIiwgIkdTVElOIiwKICAgICAgICAiU3Vi"
    "LXNlY3Rpb24iLCAiU3ViLXNlY3Rpb24gRGVzYy4iLCAiVGF4IFR5cGUiLAogICAgICAgICJUYXgg"
    "UGF5YWJsZSAo4oK5KSIsICJBZGp1c3RtZW50ICjigrkpIiwgIk5ldCBUYXggUGF5YWJsZSAo4oK5"
    "KSIsCiAgICAgICAgIlBhaWQgdmlhIElUQyAtIElHU1QgKOKCuSkiLCAiUGFpZCB2aWEgSVRDIC0g"
    "Q0dTVCAo4oK5KSIsCiAgICAgICAgIlBhaWQgdmlhIElUQyAtIFNHU1QvVVQgKOKCuSkiLCAiUGFp"
    "ZCB2aWEgSVRDIC0gQ2VzcyAo4oK5KSIsCiAgICAgICAgIlRheCBQYWlkIGluIENhc2ggKOKCuSki"
    "LCAiSW50ZXJlc3QgaW4gQ2FzaCAo4oK5KSIsICJMYXRlIEZlZSBpbiBDYXNoICjigrkpIiwKICAg"
    "IF0KICAgIGZvciBpLCBoIGluIGVudW1lcmF0ZShoMiwgc3RhcnQ9MSk6CiAgICAgICAgYyA9IHdz"
    "Mi5jZWxsKHJvdz0zLCBjb2x1bW49aSwgdmFsdWU9aCkKICAgICAgICBjLmZvbnQgPSBXSElURTsg"
    "Yy5maWxsID0gU1VCX0ZJTEw7IGMuYWxpZ25tZW50ID0gQ0VOVEVSOyBjLmJvcmRlciA9IEJPUkRF"
    "UgogICAgd3MyLnJvd19kaW1lbnNpb25zWzNdLmhlaWdodCA9IDQyCgogICAgc3IgPSAwOyByciA9"
    "IDQKICAgIGZvciByZXQgaW4gcmV0dXJuczoKICAgICAgICBtID0gcmV0WyJtZXRhIl0KICAgICAg"
    "ICBmb3Igcm93IGluIHJldC5nZXQoInBheW1lbnQiLCBbXSk6CiAgICAgICAgICAgIHNyICs9IDEK"
    "ICAgICAgICAgICAgdmFscyA9IFsKICAgICAgICAgICAgICAgIHNyLCBtWyJtb250aCJdLCBtWyJm"
    "eSJdLCBtWyJzdGF0ZV9jb2RlIl0sIG1bInN0YXRlX25hbWUiXSwgbVsiZ3N0aW4iXSwKICAgICAg"
    "ICAgICAgICAgIHJvd1sic3Vic2VjdGlvbiJdLCByb3dbInN1YnNlY3Rpb25fZGVzYyJdLCByb3db"
    "InRheF90eXBlIl0sCiAgICAgICAgICAgICAgICByb3dbInRheF9wYXlhYmxlIl0sIHJvd1siYWRq"
    "dXN0bWVudCJdLCByb3dbIm5ldF90YXhfcGF5YWJsZSJdLAogICAgICAgICAgICAgICAgcm93WyJp"
    "dGNfaWdzdCJdLCByb3dbIml0Y19jZ3N0Il0sIHJvd1siaXRjX3Nnc3QiXSwgcm93WyJpdGNfY2Vz"
    "cyJdLAogICAgICAgICAgICAgICAgcm93WyJ0YXhfaW5fY2FzaCJdLCByb3dbImludGVyZXN0X2Nh"
    "c2giXSwgcm93WyJsYXRlX2ZlZV9jYXNoIl0sCiAgICAgICAgICAgIF0KICAgICAgICAgICAgZm9y"
    "IGksIHYgaW4gZW51bWVyYXRlKHZhbHMsIHN0YXJ0PTEpOgogICAgICAgICAgICAgICAgY2VsbCA9"
    "IHdzMi5jZWxsKHJvdz1yciwgY29sdW1uPWksIHZhbHVlPXYpCiAgICAgICAgICAgICAgICBjZWxs"
    "LmZvbnQgPSBSRUc7IGNlbGwuYm9yZGVyID0gQk9SREVSCiAgICAgICAgICAgICAgICBpZiBpIGlu"
    "ICgxLCAyLCAzLCA0LCA3KToKICAgICAgICAgICAgICAgICAgICBjZWxsLmFsaWdubWVudCA9IENF"
    "TlRFUgogICAgICAgICAgICAgICAgZWxpZiBpIGluICg1LCA2LCA4LCA5KToKICAgICAgICAgICAg"
    "ICAgICAgICBjZWxsLmFsaWdubWVudCA9IExFRlQKICAgICAgICAgICAgICAgIGVsc2U6CiAgICAg"
    "ICAgICAgICAgICAgICAgY2VsbC5hbGlnbm1lbnQgPSBSSUdIVAogICAgICAgICAgICAgICAgICAg"
    "IGNlbGwubnVtYmVyX2Zvcm1hdCA9IE5VTV9GTVQKICAgICAgICAgICAgcnIgKz0gMQoKICAgIGZv"
    "ciBjb2wsIHcgaW4gWygiQSIsNiksKCJCIiw5KSwoIkMiLDkpLCgiRCIsNiksKCJFIiwxOCksKCJG"
    "IiwyMCksCiAgICAgICAgICAgICAgICAgICAoIkciLDEwKSwoIkgiLDMwKSwoIkkiLDE2KSwKICAg"
    "ICAgICAgICAgICAgICAgICgiSiIsMTYpLCgiSyIsMTQpLCgiTCIsMTgpLAogICAgICAgICAgICAg"
    "ICAgICAgKCJNIiwxOCksKCJOIiwxOCksKCJPIiwxOCksKCJQIiwxNiksCiAgICAgICAgICAgICAg"
    "ICAgICAoIlEiLDE4KSwoIlIiLDE4KSwoIlMiLDE4KV06CiAgICAgICAgd3MyLmNvbHVtbl9kaW1l"
    "bnNpb25zW2NvbF0ud2lkdGggPSB3CiAgICB3czIuZnJlZXplX3BhbmVzID0gIko0IgogICAgaWYg"
    "cnIgPiA0OgogICAgICAgIHdzMi5hdXRvX2ZpbHRlci5yZWYgPSBmIkEzOlN7cnItMX0iCgogICAg"
    "IyA9PT09PT09PT09IFNoZWV0IDM6IFN0YXRlLU1vbnRoIFN1bW1hcnkgPT09PT09PT09PT09PT09"
    "PT09PT0KICAgIHdzMyA9IHdiLmNyZWF0ZV9zaGVldCgiU3RhdGUtTW9udGggU3VtbWFyeSIpCiAg"
    "ICB3czMubWVyZ2VfY2VsbHMoIkExOk4xIikKICAgIHdzM1siQTEiXSA9ICJLRVkgVE9UQUxTIFBF"
    "UiBSRVRVUk4iCiAgICB3czNbIkExIl0uZm9udCA9IFRJVExFOyB3czNbIkExIl0uZmlsbCA9IEhE"
    "Ul9GSUxMOyB3czNbIkExIl0uYWxpZ25tZW50ID0gQ0VOVEVSCiAgICB3czMucm93X2RpbWVuc2lv"
    "bnNbMV0uaGVpZ2h0ID0gMjgKCiAgICBoMyA9IFsKICAgICAgICAiU3IgTm8iLCAiTW9udGgiLCAi"
    "U3RhdGUgQ29kZSIsICJTdGF0ZSBOYW1lIiwgIkdTVElOIiwKICAgICAgICAiMy4xKGEpIFRheGFi"
    "bGUgKOKCuSkiLCAiMy4xKGEpIElHU1QgKOKCuSkiLCAiMy4xKGEpIENHU1QgKOKCuSkiLCAiMy4x"
    "KGEpIFNHU1QgKOKCuSkiLAogICAgICAgICJUb3RhbCBJVEMgQXZhaWwuIOKAlCBBICjigrkpIiwK"
    "ICAgICAgICAiVG90YWwgSVRDIFJldmVyc2VkIOKAlCBCICjigrkpIiwKICAgICAgICAiTmV0IElU"
    "QyDigJQgQyAo4oK5KSIsCiAgICAgICAgIkNhc2ggUGFpZCAoNi4xIOKAlCBJR1NUK0NHU1QrU0dT"
    "VCtDZXNzKSAo4oK5KSIsCiAgICAgICAgIlRvdGFsIExpYWJpbGl0eSBCcmVha3VwICjigrkpIiwK"
    "ICAgIF0KICAgIGZvciBpLCBoIGluIGVudW1lcmF0ZShoMywgc3RhcnQ9MSk6CiAgICAgICAgYyA9"
    "IHdzMy5jZWxsKHJvdz0zLCBjb2x1bW49aSwgdmFsdWU9aCkKICAgICAgICBjLmZvbnQgPSBXSElU"
    "RTsgYy5maWxsID0gU1VCX0ZJTEw7IGMuYWxpZ25tZW50ID0gQ0VOVEVSOyBjLmJvcmRlciA9IEJP"
    "UkRFUgogICAgd3MzLnJvd19kaW1lbnNpb25zWzNdLmhlaWdodCA9IDUwCgogICAgc3IgPSAwOyBy"
    "ciA9IDQKICAgIGZvciByZXQgaW4gcmV0dXJuczoKICAgICAgICBtID0gcmV0WyJtZXRhIl0KICAg"
    "ICAgICBjb25zID0gcmV0LmdldCgiY29uc29saWRhdGVkIiwgW10pCiAgICAgICAgcGF5ID0gcmV0"
    "LmdldCgicGF5bWVudCIsIFtdKQoKICAgICAgICBkZWYgZmluZChzZWN0aW9uX3ByZWZpeCk6CiAg"
    "ICAgICAgICAgIHJldHVybiBuZXh0KChyb3cgZm9yIHJvdyBpbiBjb25zIGlmIHJvd1sic2VjdGlv"
    "biJdID09IHNlY3Rpb25fcHJlZml4KSwgTm9uZSkKCiAgICAgICAgcl8zXzFfYSA9IGZpbmQoIjMu"
    "MShhKSIpCiAgICAgICAgIyBTdW0gNChBKSgxLi41KQogICAgICAgIGFfaWdzdCA9IHN1bShyb3db"
    "Imlnc3QiXSBmb3Igcm93IGluIGNvbnMgaWYgcm93WyJzZWN0aW9uIl0uc3RhcnRzd2l0aCgiNChB"
    "KSgiKSkKICAgICAgICBhX2Nnc3QgPSBzdW0ocm93WyJjZ3N0Il0gZm9yIHJvdyBpbiBjb25zIGlm"
    "IHJvd1sic2VjdGlvbiJdLnN0YXJ0c3dpdGgoIjQoQSkoIikpCiAgICAgICAgYV9zZ3N0ID0gc3Vt"
    "KHJvd1sic2dzdCJdIGZvciByb3cgaW4gY29ucyBpZiByb3dbInNlY3Rpb24iXS5zdGFydHN3aXRo"
    "KCI0KEEpKCIpKQogICAgICAgIGFfY2VzcyA9IHN1bShyb3dbImNlc3MiXSBmb3Igcm93IGluIGNv"
    "bnMgaWYgcm93WyJzZWN0aW9uIl0uc3RhcnRzd2l0aCgiNChBKSgiKSkKICAgICAgICB0b3RhbF9h"
    "ID0gYV9pZ3N0ICsgYV9jZ3N0ICsgYV9zZ3N0ICsgYV9jZXNzCgogICAgICAgICMgU3VtIDQoQiko"
    "MS4uMikKICAgICAgICBiX2lnc3QgPSBzdW0ocm93WyJpZ3N0Il0gZm9yIHJvdyBpbiBjb25zIGlm"
    "IHJvd1sic2VjdGlvbiJdLnN0YXJ0c3dpdGgoIjQoQikoIikpCiAgICAgICAgYl9jZ3N0ID0gc3Vt"
    "KHJvd1siY2dzdCJdIGZvciByb3cgaW4gY29ucyBpZiByb3dbInNlY3Rpb24iXS5zdGFydHN3aXRo"
    "KCI0KEIpKCIpKQogICAgICAgIGJfc2dzdCA9IHN1bShyb3dbInNnc3QiXSBmb3Igcm93IGluIGNv"
    "bnMgaWYgcm93WyJzZWN0aW9uIl0uc3RhcnRzd2l0aCgiNChCKSgiKSkKICAgICAgICBiX2Nlc3Mg"
    "PSBzdW0ocm93WyJjZXNzIl0gZm9yIHJvdyBpbiBjb25zIGlmIHJvd1sic2VjdGlvbiJdLnN0YXJ0"
    "c3dpdGgoIjQoQikoIikpCiAgICAgICAgdG90YWxfYiA9IGJfaWdzdCArIGJfY2dzdCArIGJfc2dz"
    "dCArIGJfY2VzcwoKICAgICAgICByXzRfYyA9IGZpbmQoIjQoQykiKQogICAgICAgIGNfdG90YWwg"
    "PSAwLjAKICAgICAgICBpZiByXzRfYzoKICAgICAgICAgICAgY190b3RhbCA9IChyXzRfYy5nZXQo"
    "Imlnc3QiKSBvciAwKSArIChyXzRfYy5nZXQoImNnc3QiKSBvciAwKSArIChyXzRfYy5nZXQoInNn"
    "c3QiKSBvciAwKSArIChyXzRfYy5nZXQoImNlc3MiKSBvciAwKQoKICAgICAgICAjIENhc2ggcGFp"
    "ZCB0b3RhbAogICAgICAgIGNhc2hfcGFpZCA9IHN1bShyb3dbInRheF9pbl9jYXNoIl0gZm9yIHJv"
    "dyBpbiBwYXkpCgogICAgICAgICMgQnJlYWt1cCB0b3RhbAogICAgICAgIGJyZWFrdXBfdG90YWwg"
    "PSAwLjAKICAgICAgICBmb3Igcm93IGluIGNvbnM6CiAgICAgICAgICAgIGlmIHJvd1sic2VjdGlv"
    "biJdID09ICJCcmVha3VwIjoKICAgICAgICAgICAgICAgIGJyZWFrdXBfdG90YWwgPSAocm93Lmdl"
    "dCgiaWdzdCIpIG9yIDApICsgKHJvdy5nZXQoImNnc3QiKSBvciAwKSArIChyb3cuZ2V0KCJzZ3N0"
    "Iikgb3IgMCkgKyAocm93LmdldCgiY2VzcyIpIG9yIDApCiAgICAgICAgICAgICAgICBicmVhawoK"
    "ICAgICAgICBzciArPSAxCiAgICAgICAgdmFscyA9IFsKICAgICAgICAgICAgc3IsIG1bIm1vbnRo"
    "Il0sIG1bInN0YXRlX2NvZGUiXSwgbVsic3RhdGVfbmFtZSJdLCBtWyJnc3RpbiJdLAogICAgICAg"
    "ICAgICAocl8zXzFfYSBvciB7fSkuZ2V0KCJ0YXhhYmxlIiwgMC4wKSwKICAgICAgICAgICAgKHJf"
    "M18xX2Egb3Ige30pLmdldCgiaWdzdCIsIDAuMCksCiAgICAgICAgICAgIChyXzNfMV9hIG9yIHt9"
    "KS5nZXQoImNnc3QiLCAwLjApLAogICAgICAgICAgICAocl8zXzFfYSBvciB7fSkuZ2V0KCJzZ3N0"
    "IiwgMC4wKSwKICAgICAgICAgICAgdG90YWxfYSwgdG90YWxfYiwgY190b3RhbCwgY2FzaF9wYWlk"
    "LCBicmVha3VwX3RvdGFsLAogICAgICAgIF0KICAgICAgICBmb3IgaSwgdiBpbiBlbnVtZXJhdGUo"
    "dmFscywgc3RhcnQ9MSk6CiAgICAgICAgICAgIGNlbGwgPSB3czMuY2VsbChyb3c9cnIsIGNvbHVt"
    "bj1pLCB2YWx1ZT12KQogICAgICAgICAgICBjZWxsLmZvbnQgPSBSRUc7IGNlbGwuYm9yZGVyID0g"
    "Qk9SREVSCiAgICAgICAgICAgIGlmIGkgPD0gNDoKICAgICAgICAgICAgICAgIGNlbGwuYWxpZ25t"
    "ZW50ID0gQ0VOVEVSIGlmIGkgIT0gNCBlbHNlIExFRlQKICAgICAgICAgICAgZWxpZiBpID09IDU6"
    "CiAgICAgICAgICAgICAgICBjZWxsLmFsaWdubWVudCA9IExFRlQKICAgICAgICAgICAgZWxzZToK"
    "ICAgICAgICAgICAgICAgIGNlbGwuYWxpZ25tZW50ID0gUklHSFQKICAgICAgICAgICAgICAgIGNl"
    "bGwubnVtYmVyX2Zvcm1hdCA9IE5VTV9GTVQKICAgICAgICByciArPSAxCgogICAgZm9yIGNvbCwg"
    "dyBpbiBbKCJBIiw2KSwoIkIiLDkpLCgiQyIsNiksKCJEIiwxOCksKCJFIiwyMiksCiAgICAgICAg"
    "ICAgICAgICAgICAoIkYiLDE4KSwoIkciLDE2KSwoIkgiLDE2KSwoIkkiLDE2KSwKICAgICAgICAg"
    "ICAgICAgICAgICgiSiIsMjApLCgiSyIsMjApLCgiTCIsMTgpLCgiTSIsMjYpLCgiTiIsMjIpXToK"
    "ICAgICAgICB3czMuY29sdW1uX2RpbWVuc2lvbnNbY29sXS53aWR0aCA9IHcKICAgIHdzMy5mcmVl"
    "emVfcGFuZXMgPSAiQTQiCiAgICBpZiByciA+IDQ6CiAgICAgICAgd3MzLmF1dG9fZmlsdGVyLnJl"
    "ZiA9IGYiQTM6Tntyci0xfSIKCiAgICAjID09PT09PT09PT0gU2hlZXQgNDogUHJvY2Vzc2luZyBM"
    "b2cgPT09PT09PT09PT09PT09PT09PT09PT09PQogICAgd3M0ID0gd2IuY3JlYXRlX3NoZWV0KCJQ"
    "cm9jZXNzaW5nIExvZyIpCiAgICBsb2dfaGVhZGVycyA9IFsiIyIsICJTb3VyY2UgRmlsZSIsICJT"
    "dGF0dXMiLCAiR1NUSU4iLCAiU3RhdGUiLCAiTW9udGgiLAogICAgICAgICAgICAgICAgICAgIkNv"
    "bnNvbGlkYXRlZCBSb3dzIiwgIlBheW1lbnQgUm93cyIsICJOb3RlcyJdCiAgICB3czQuYXBwZW5k"
    "KGxvZ19oZWFkZXJzKQogICAgZm9yIGMgaW4gcmFuZ2UoMSwgbGVuKGxvZ19oZWFkZXJzKSArIDEp"
    "OgogICAgICAgIGNlbGwgPSB3czQuY2VsbChyb3c9MSwgY29sdW1uPWMpCiAgICAgICAgY2VsbC5m"
    "b250ID0gV0hJVEU7IGNlbGwuZmlsbCA9IFNVQl9GSUxMOyBjZWxsLmFsaWdubWVudCA9IENFTlRF"
    "UjsgY2VsbC5ib3JkZXIgPSBCT1JERVIKICAgIGZvciBpLCByZXQgaW4gZW51bWVyYXRlKHJldHVy"
    "bnMsIHN0YXJ0PTEpOgogICAgICAgIG0gPSByZXRbIm1ldGEiXQogICAgICAgIHN0YXR1cyA9IHJl"
    "dC5nZXQoInN0YXR1cyIsICJPSyIpCiAgICAgICAgb3V0ID0gW2ksIHJldC5nZXQoInNvdXJjZV9m"
    "aWxlIiwiIiksIHN0YXR1cywKICAgICAgICAgICAgICAgbS5nZXQoImdzdGluIiwiIiksIG0uZ2V0"
    "KCJzdGF0ZV9uYW1lIiwiIiksIG0uZ2V0KCJtb250aCIsIiIpLAogICAgICAgICAgICAgICBsZW4o"
    "cmV0LmdldCgiY29uc29saWRhdGVkIiwgW10pKSwgbGVuKHJldC5nZXQoInBheW1lbnQiLCBbXSkp"
    "LAogICAgICAgICAgICAgICByZXQuZ2V0KCJub3RlcyIsIiIpXQogICAgICAgIGZvciBjLCB2IGlu"
    "IGVudW1lcmF0ZShvdXQsIHN0YXJ0PTEpOgogICAgICAgICAgICBjZWxsID0gd3M0LmNlbGwocm93"
    "PWkrMSwgY29sdW1uPWMsIHZhbHVlPXYpCiAgICAgICAgICAgIGNlbGwuZm9udCA9IFJFRzsgY2Vs"
    "bC5ib3JkZXIgPSBCT1JERVI7IGNlbGwuYWxpZ25tZW50ID0gTEVGVAogICAgICAgICAgICBpZiBz"
    "dGF0dXMgIT0gIk9LIjoKICAgICAgICAgICAgICAgIGNlbGwuZmlsbCA9IEVSUl9GSUxMCiAgICBm"
    "b3IgY29sLCB3IGluIFsoIkEiLDYpLCgiQiIsNDIpLCgiQyIsMTApLCgiRCIsMjIpLCgiRSIsMjIp"
    "LCgiRiIsMTApLAogICAgICAgICAgICAgICAgICAgKCJHIiwxOCksKCJIIiwxMyksKCJJIiw1MCld"
    "OgogICAgICAgIHdzNC5jb2x1bW5fZGltZW5zaW9uc1tjb2xdLndpZHRoID0gdwogICAgd3M0LmZy"
    "ZWV6ZV9wYW5lcyA9ICJBMiIKCiAgICB3Yi5zYXZlKG91dHB1dF9wYXRoKQoKCiMgLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tICMKIyAgT3V0cHV0IHBhdGggaGVscGVyCiMgLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tICMK"
    "CmRlZiByZXNvbHZlX291dHB1dF9wYXRoKHJhdyk6CiAgICBwID0gUGF0aChyYXcpCiAgICBERUZB"
    "VUxUX05BTUUgPSAiR1NUUjNCX0NvbnNvbGlkYXRlZC54bHN4IgogICAgaWYgcC5leGlzdHMoKSBh"
    "bmQgcC5pc19kaXIoKToKICAgICAgICBwID0gcCAvIERFRkFVTFRfTkFNRQogICAgZWxpZiBwLnN1"
    "ZmZpeC5sb3dlcigpICE9ICIueGxzeCI6CiAgICAgICAgcCA9IHAud2l0aF9zdWZmaXgoIi54bHN4"
    "IikKICAgIHAucGFyZW50Lm1rZGlyKHBhcmVudHM9VHJ1ZSwgZXhpc3Rfb2s9VHJ1ZSkKICAgIGlm"
    "IHAuZXhpc3RzKCk6CiAgICAgICAgdHJ5OgogICAgICAgICAgICB3aXRoIG9wZW4ocCwgImFiIik6"
    "CiAgICAgICAgICAgICAgICBwYXNzCiAgICAgICAgZXhjZXB0IFBlcm1pc3Npb25FcnJvcjoKICAg"
    "ICAgICAgICAgcmFpc2UgUGVybWlzc2lvbkVycm9yKAogICAgICAgICAgICAgICAgZiJPdXRwdXQg"
    "ZmlsZSBpcyBsb2NrZWQgKHByb2JhYmx5IG9wZW4gaW4gRXhjZWwpOlxuICB7cH1cbiIKICAgICAg"
    "ICAgICAgICAgICJDbG9zZSBpdCBpbiBFeGNlbCBhbmQgcmUtcnVuLiIKICAgICAgICAgICAgKQog"
    "ICAgcmV0dXJuIHAKCgojIC0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLSAjCiMgIFByb2Nlc3NpbmcgcGlwZWxp"
    "bmUKIyAtLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0gIwoKZGVmIHByb2Nlc3NfcGRmcyhpbl9mb2xkZXIsIG91"
    "dF9wYXRoLCBvbl9wcm9ncmVzcz1Ob25lLCBvbl9sb2c9Tm9uZSk6CiAgICBwZGZzID0gc29ydGVk"
    "KFtwIGZvciBwIGluIGluX2ZvbGRlci5yZ2xvYigiKi5wZGYiKV0pCiAgICBpZiBub3QgcGRmczoK"
    "ICAgICAgICByYWlzZSBGaWxlTm90Rm91bmRFcnJvcihmIk5vIFBERnMgZm91bmQgdW5kZXI6IHtp"
    "bl9mb2xkZXJ9IikKCiAgICBpZiBvbl9sb2c6CiAgICAgICAgb25fbG9nKGYiRm91bmQge2xlbihw"
    "ZGZzKX0gUERGIGZpbGUocykuICBQcm9jZXNzaW5nLi4uIiwgImluZm8iKQoKICAgIHJldHVybnMg"
    "PSBbXQogICAgb2sgPSAwOyBmYWlsID0gMAoKICAgIGZvciBpZHgsIHBkZl9wYXRoIGluIGVudW1l"
    "cmF0ZShwZGZzLCBzdGFydD0xKToKICAgICAgICB0cnk6CiAgICAgICAgICAgIHJlbCA9IHBkZl9w"
    "YXRoLnJlbGF0aXZlX3RvKGluX2ZvbGRlcikKICAgICAgICBleGNlcHQgVmFsdWVFcnJvcjoKICAg"
    "ICAgICAgICAgcmVsID0gcGRmX3BhdGgubmFtZQoKICAgICAgICB0cnk6CiAgICAgICAgICAgIG1l"
    "dGEsIGNvbnNvbGlkYXRlZCwgcGF5bWVudCA9IHBhcnNlX3BkZihwZGZfcGF0aCkKICAgICAgICAg"
    "ICAgcmV0dXJucy5hcHBlbmQoewogICAgICAgICAgICAgICAgIm1ldGEiOiBtZXRhLCAiY29uc29s"
    "aWRhdGVkIjogY29uc29saWRhdGVkLCAicGF5bWVudCI6IHBheW1lbnQsCiAgICAgICAgICAgICAg"
    "ICAic291cmNlX2ZpbGUiOiBzdHIocmVsKSwgInN0YXR1cyI6ICJPSyIsICJub3RlcyI6ICIiLAog"
    "ICAgICAgICAgICB9KQogICAgICAgICAgICBvayArPSAxCiAgICAgICAgICAgIGlmIG9uX2xvZzoK"
    "ICAgICAgICAgICAgICAgIG9uX2xvZygKICAgICAgICAgICAgICAgICAgICBmIlt7aWR4Oj40fS97"
    "bGVuKHBkZnMpfV0gIE9LICAgICIKICAgICAgICAgICAgICAgICAgICBmInttZXRhWydzdGF0ZV9u"
    "YW1lJ106PDIyfSB7bWV0YVsnbW9udGgnXTo8OH0gICIKICAgICAgICAgICAgICAgICAgICBmIntt"
    "ZXRhWydnc3RpbiddfSAgKHtwZGZfcGF0aC5uYW1lfSkiLAogICAgICAgICAgICAgICAgICAgICJv"
    "ayIsCiAgICAgICAgICAgICAgICApCiAgICAgICAgZXhjZXB0IEV4Y2VwdGlvbiBhcyBlOgogICAg"
    "ICAgICAgICBmYWlsICs9IDEKICAgICAgICAgICAgcmV0dXJucy5hcHBlbmQoewogICAgICAgICAg"
    "ICAgICAgIm1ldGEiOiB7ImdzdGluIjoiIiwic3RhdGVfY29kZSI6IiIsInN0YXRlX25hbWUiOiIi"
    "LCJtb250aCI6IiIsCiAgICAgICAgICAgICAgICAgICAgICAgICAiZnkiOiIiLCJsZWdhbF9uYW1l"
    "IjoiIiwiYXJuIjoiIn0sCiAgICAgICAgICAgICAgICAiY29uc29saWRhdGVkIjogW10sICJwYXlt"
    "ZW50IjogW10sCiAgICAgICAgICAgICAgICAic291cmNlX2ZpbGUiOiBzdHIocmVsKSwKICAgICAg"
    "ICAgICAgICAgICJzdGF0dXMiOiAiRkFJTCIsCiAgICAgICAgICAgICAgICAibm90ZXMiOiBmInt0"
    "eXBlKGUpLl9fbmFtZV9ffToge2V9IiwKICAgICAgICAgICAgfSkKICAgICAgICAgICAgaWYgb25f"
    "bG9nOgogICAgICAgICAgICAgICAgb25fbG9nKGYiW3tpZHg6PjR9L3tsZW4ocGRmcyl9XSAgRkFJ"
    "TCAge3BkZl9wYXRoLm5hbWV9ICDihpIgIHtlfSIsICJmYWlsIikKCiAgICAgICAgaWYgb25fcHJv"
    "Z3Jlc3M6CiAgICAgICAgICAgIG9uX3Byb2dyZXNzKGlkeCwgbGVuKHBkZnMpKQoKICAgIGlmIG9u"
    "X2xvZzoKICAgICAgICBvbl9sb2coZiJXcml0aW5nIGNvbnNvbGlkYXRlZCBFeGNlbDogIHtvdXRf"
    "cGF0aH0iLCAiaW5mbyIpCiAgICB3cml0ZV9leGNlbChyZXR1cm5zLCBvdXRfcGF0aCkKICAgIHJl"
    "dHVybiBvaywgZmFpbCwgbGVuKHBkZnMpLCBvdXRfcGF0aAoKCiMgLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "ICMKIyAgR1VJCiMgLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0t"
    "LS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tLS0tICMKCg=="
)

_GSTR1_NS = {"__name__": "_gstr1_engine"}
_GSTR3B_NS = {"__name__": "_gstr3b_engine"}

def _load_extractor_engines():
    """Decode and exec the embedded engines into their isolated namespaces.
    Called lazily — only when the user clicks Run on those tabs."""
    if not HAS_PDFPLUMBER:
        raise RuntimeError(
            "pdfplumber not installed. In CMD:  pip install pdfplumber openpyxl")
    if "parse_pdf" not in _GSTR1_NS:
        src1 = _base64.b64decode("".join(_GSTR1_ENGINE_B64)).decode("utf-8")
        # Provide pre-imported modules so engines don't re-import
        _GSTR1_NS["os"] = os
        _GSTR1_NS["re"] = re
        _GSTR1_NS["sys"] = sys
        _GSTR1_NS["traceback"] = traceback
        from pathlib import Path as _P
        _GSTR1_NS["Path"] = _P
        import pdfplumber as _pp
        _GSTR1_NS["pdfplumber"] = _pp
        from openpyxl import Workbook as _WB
        from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A, Border as _B, Side as _S
        from openpyxl.utils import get_column_letter as _gcl
        _GSTR1_NS["Workbook"] = _WB
        _GSTR1_NS["Font"] = _F; _GSTR1_NS["PatternFill"] = _PF
        _GSTR1_NS["Alignment"] = _A; _GSTR1_NS["Border"] = _B; _GSTR1_NS["Side"] = _S
        _GSTR1_NS["get_column_letter"] = _gcl
        exec(src1, _GSTR1_NS)

    if "parse_pdf" not in _GSTR3B_NS:
        src3 = _base64.b64decode("".join(_GSTR3B_ENGINE_B64)).decode("utf-8")
        _GSTR3B_NS["os"] = os
        _GSTR3B_NS["re"] = re
        _GSTR3B_NS["sys"] = sys
        from pathlib import Path as _P
        _GSTR3B_NS["Path"] = _P
        import pdfplumber as _pp
        _GSTR3B_NS["pdfplumber"] = _pp
        from openpyxl import Workbook as _WB
        from openpyxl.styles import Font as _F, PatternFill as _PF, Alignment as _A, Border as _B, Side as _S
        _GSTR3B_NS["Workbook"] = _WB
        _GSTR3B_NS["Font"] = _F; _GSTR3B_NS["PatternFill"] = _PF
        _GSTR3B_NS["Alignment"] = _A; _GSTR3B_NS["Border"] = _B; _GSTR3B_NS["Side"] = _S
        exec(src3, _GSTR3B_NS)


# ════════════════════════════════════════════════════════════════
#  TAX COMPARISON CONSOLIDATOR ENGINE
# ════════════════════════════════════════════════════════════════
# Reference for state codes
_STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram",
    "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
    "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh",
    "24": "Gujarat", "25": "Daman & Diu", "26": "Dadra & NH and Daman & Diu",
    "27": "Maharashtra", "28": "Andhra Pradesh (Old)", "29": "Karnataka",
    "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman & Nicobar Islands", "36": "Telangana",
    "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
    "99": "Centre Jurisdiction",
}

# Each comparison file from GSTN portal contains these 8 sheets with known structure.
# (header_row_for_columns, data_start_row, data_end_marker_col_A='Total', max_cols)
_TC_SHEET_LAYOUT = {
    "Tax Liability Summary": {
        "main_header_row":  8,    # column-group labels (rowspan with row 7)
        "sub_header_row":   9,    # IGST/CGST/SGST/CESS sub-labels
        "data_start_row":   10,
        "expected_cols":    27,
    },
    "Comparison Summary": {
        "main_header_row":  8,
        "sub_header_row":   9,
        "data_start_row":   10,
        "expected_cols":    15,
    },
    "Tax liability": {
        "main_header_row":  5,
        "sub_header_row":   6,
        "data_start_row":   7,
        "expected_cols":    37,
    },
    "Reverse charge": {
        "main_header_row":  5,
        "sub_header_row":   6,
        "data_start_row":   7,
        "expected_cols":    37,
    },
    "Export and SEZ": {
        "main_header_row":  5,
        "sub_header_row":   6,
        "data_start_row":   7,
        "expected_cols":    11,
    },
    "ITC (Other than IMPG)": {
        "main_header_row":  5,
        "sub_header_row":   6,
        "data_start_row":   7,
        "expected_cols":    37,
    },
    "ITC (IMPG)": {
        "main_header_row":  5,
        "sub_header_row":   6,
        "data_start_row":   7,
        "expected_cols":    11,
    },
    "RCM_LIABILITY_ITC": {
        "main_header_row":  5,
        "sub_header_row":   6,
        "data_start_row":   7,
        "expected_cols":    37,
    },
}


def parse_tax_comparison_file(filepath):
    """
    Parse a single GSTN portal 'Tax liability & ITC comparison' xlsx.
    Returns: {
        "meta": {gstin, fy, state_code, state_name, legal_name, trade_name, report_dt, source_file},
        "sheets": {sheet_name: {"main_headers": [...], "sub_headers": [...], "rows": [{month, values:[...]}, ...]}}
    }
    """
    import openpyxl, re
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)

    # --- Extract meta from filename + first sheet ---
    fname = os.path.basename(filepath)
    m = re.search(r"(\d{4}-\d{2})_(\d{2}[A-Z0-9]{13})_", fname)
    fy_from_name = m.group(1) if m else ""
    gstin_from_name = m.group(2) if m else ""

    meta = {
        "source_file": fname,
        "gstin": gstin_from_name, "fy": fy_from_name,
        "state_code": gstin_from_name[:2] if gstin_from_name else "",
        "state_name": _STATE_CODES.get(gstin_from_name[:2], "Unknown") if gstin_from_name else "Unknown",
        "legal_name": "", "trade_name": "", "report_dt": "",
    }

    # Pull legal/trade names from any sheet (typically rows 4-5)
    for sn in wb.sheetnames:
        ws = wb[sn]
        for r in range(1, min(ws.max_row + 1, 8)):
            for c in range(1, min(ws.max_column + 1, 15)):
                v = ws.cell(r, c).value
                if not isinstance(v, str): continue
                t = v.strip()
                if t.startswith("GSTIN:"):
                    val = t.split(":", 1)[1].strip()
                    if val and not meta["gstin"]: meta["gstin"] = val
                elif t.startswith("Legal name:"):
                    if not meta["legal_name"]:
                        meta["legal_name"] = t.split(":", 1)[1].strip()
                elif t.startswith("Trade name"):
                    if not meta["trade_name"]:
                        meta["trade_name"] = t.split(":", 1)[1].strip()
                elif t.startswith("Financial Year:"):
                    if not meta["fy"]:
                        meta["fy"] = t.split(":", 1)[1].strip()
                elif t.startswith("Report generated"):
                    if not meta["report_dt"]:
                        meta["report_dt"] = t.split(":", 1)[1].strip() \
                            if ":" in t else t
        if meta["legal_name"]: break

    # Refill state info if GSTIN was only inside file
    if meta["gstin"] and not meta["state_code"]:
        meta["state_code"] = meta["gstin"][:2]
        meta["state_name"] = _STATE_CODES.get(meta["state_code"], "Unknown")

    # --- Extract each sheet ---
    sheets = {}
    for sn, layout in _TC_SHEET_LAYOUT.items():
        if sn not in wb.sheetnames:
            sheets[sn] = {"missing": True}
            continue
        ws = wb[sn]
        # Collect headers — first read main_header row (with merged-cell expansion),
        # then sub_header row. We capture column-by-column to preserve order.
        mh = []
        sh = []
        max_c = min(ws.max_column, layout["expected_cols"])
        for c in range(1, max_c + 1):
            # main header — resolve merged cells: if the cell is in a merged range,
            # take the top-left value of that range
            mh.append(_resolve_merged_value(ws, layout["main_header_row"], c))
            sh.append(_resolve_merged_value(ws, layout["sub_header_row"], c))
        # Data rows — read until "Total" found in column A
        rows = []
        r = layout["data_start_row"]
        while r <= ws.max_row:
            month = ws.cell(r, 1).value
            if month is None or (isinstance(month, str) and not month.strip()):
                r += 1
                continue
            values = []
            for c in range(2, max_c + 1):
                values.append(ws.cell(r, c).value)
            rows.append({
                "month":    str(month).strip(),
                "is_total": (str(month).strip().lower() == "total"),
                "values":   values,
            })
            if str(month).strip().lower() == "total":
                break
            r += 1
        sheets[sn] = {
            "main_headers": mh,
            "sub_headers":  sh,
            "rows":         rows,
        }
    wb.close()
    return {"meta": meta, "sheets": sheets}


def _resolve_merged_value(ws, row, col):
    """Return the cell value; if the cell is inside a merged range, return the
    top-left cell's value of that range."""
    cell = ws.cell(row, col)
    val = cell.value
    if val is not None and (not isinstance(val, str) or val.strip()):
        return val
    # Check merged ranges
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
            return ws.cell(mr.min_row, mr.min_col).value
    return None


# Pretty short labels for output sheets (Excel limit 31 chars)
_SHORT_NAMES = {
    "Tax Liability Summary":  "Tax Liability Summary",
    "Comparison Summary":     "Comparison Summary",
    "Tax liability":          "1. Tax Liability",
    "Reverse charge":         "2. Reverse Charge",
    "Export and SEZ":         "3. Export & SEZ",
    "ITC (Other than IMPG)":  "4. ITC (Other than IMPG)",
    "ITC (IMPG)":             "5. ITC (IMPG)",
    "RCM_LIABILITY_ITC":      "6. RCM Liability & ITC",
}


def write_consolidated_comparison(all_data, out_path, mode):
    """
    all_data: list of dicts returned by parse_tax_comparison_file
    mode: 'single' | 'multi' | 'both'
    """
    import openpyxl
    from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    HDR_FILL = PatternFill("solid", start_color="1F4E78")
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    META_FILL = PatternFill("solid", start_color="FFE699")
    TOTAL_FILL = PatternFill("solid", start_color="FFF2CC")
    GROUP_FILL = PatternFill("solid", start_color="DDEBF7")

    WHITE_B = XLFont(name="Calibri", bold=True, color="FFFFFF", size=10)
    BOLD = XLFont(name="Calibri", bold=True, size=10)
    REG = XLFont(name="Calibri", size=10)
    TITLE = XLFont(name="Calibri", bold=True, color="FFFFFF", size=14)

    thin = Side(border_style="thin", color="B4B4B4")
    BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT = Alignment(horizontal="right", vertical="center")

    NUM_FMT = '#,##0.00;(#,##0.00);"-"'

    # ── 1. Cover / Index sheet ──────────────────────────────
    ws_idx = wb.create_sheet("Cover")
    ws_idx.merge_cells("A1:F1")
    ws_idx["A1"] = "GST Tax Liability & ITC Comparison — Consolidated"
    ws_idx["A1"].font = TITLE
    ws_idx["A1"].fill = HDR_FILL
    ws_idx["A1"].alignment = CENTER
    ws_idx.row_dimensions[1].height = 30

    ws_idx["A3"] = "Total states/GSTINs consolidated:"
    ws_idx["B3"] = len(all_data)
    ws_idx["A3"].font = BOLD
    ws_idx["B3"].font = BOLD

    headers = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name", "Source File"]
    for c, h in enumerate(headers, 1):
        cell = ws_idx.cell(row=5, column=c, value=h)
        cell.font = WHITE_B
        cell.fill = SUB_FILL
        cell.alignment = CENTER
        cell.border = BORDER_ALL
    ws_idx.row_dimensions[5].height = 22

    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        for c, val in enumerate([i, m["state_code"], m["state_name"], m["gstin"],
                                 m["legal_name"], m["source_file"]], 1):
            cell = ws_idx.cell(row=5 + i, column=c, value=val)
            cell.font = REG
            cell.alignment = LEFT if c >= 3 else CENTER
            cell.border = BORDER_ALL

    for col, w in zip("ABCDEF", [7, 10, 22, 22, 28, 42]):
        ws_idx.column_dimensions[col].width = w
    ws_idx.freeze_panes = "A6"

    # ── 2. Multi-sheet output (one sheet per original section) ──
    if mode in ("multi", "both"):
        for sheet_key in _TC_SHEET_LAYOUT.keys():
            ws_name = _SHORT_NAMES[sheet_key][:31]
            ws = wb.create_sheet(ws_name)

            # Title
            ws["A1"] = f"{sheet_key}  —  All States Consolidated"
            ws["A1"].font = TITLE
            ws["A1"].fill = HDR_FILL
            ws["A1"].alignment = CENTER

            # Determine column count from first available data set
            ncols = 0
            for data in all_data:
                ss = data["sheets"].get(sheet_key)
                if ss and not ss.get("missing"):
                    ncols = max(ncols, len(ss["main_headers"]))
            if ncols == 0:
                ws["A3"] = f"(Sheet '{sheet_key}' was missing from all input files.)"
                ws["A3"].font = REG
                continue

            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=ncols + 4)
            ws.row_dimensions[1].height = 28

            # Multi-row header:  S.No | State Code | State Name | GSTIN | <orig cols>
            meta_cols = ["S.No.", "State Code", "State Name", "GSTIN"]

            # Row 3: main headers
            for c, lbl in enumerate(meta_cols, 1):
                cell = ws.cell(row=3, column=c, value=lbl)
                cell.font = WHITE_B; cell.fill = SUB_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
            # use the FIRST data's main_headers as canonical
            canon_main = []
            canon_sub = []
            for data in all_data:
                ss = data["sheets"].get(sheet_key)
                if ss and not ss.get("missing"):
                    canon_main = ss["main_headers"]
                    canon_sub = ss["sub_headers"]
                    break
            for c in range(1, ncols + 1):
                mv = canon_main[c-1] if c-1 < len(canon_main) else None
                sv = canon_sub[c-1] if c-1 < len(canon_sub) else None
                cell = ws.cell(row=3, column=4 + c, value=mv)
                cell.font = WHITE_B; cell.fill = SUB_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
                cell = ws.cell(row=4, column=4 + c, value=sv)
                cell.font = BOLD; cell.fill = GROUP_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL

            for c in range(1, 5):
                ws.cell(row=4, column=c).fill = SUB_FILL
                ws.cell(row=4, column=c).border = BORDER_ALL
            ws.row_dimensions[3].height = 38
            ws.row_dimensions[4].height = 22

            # Add a "Month" column on the left side of each row
            # Restructure: S.No | State Code | State Name | GSTIN | Month | <orig cols from col 2>
            # Actually the original col 1 (Tax Period) IS the month. Re-layout below.
            # Easier: put Month as 5th meta col, then values start from col 6.
            ws.cell(row=3, column=5, value="Tax Period").font = WHITE_B
            ws.cell(row=3, column=5).fill = SUB_FILL
            ws.cell(row=3, column=5).alignment = CENTER
            ws.cell(row=3, column=5).border = BORDER_ALL
            ws.cell(row=4, column=5).fill = SUB_FILL
            ws.cell(row=4, column=5).border = BORDER_ALL
            # Shift main headers right by 1: re-write cols 6..ncols+4
            for c in range(1, ncols):
                mv = canon_main[c] if c < len(canon_main) else None
                sv = canon_sub[c] if c < len(canon_sub) else None
                cell = ws.cell(row=3, column=5 + c, value=mv)
                cell.font = WHITE_B; cell.fill = SUB_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
                cell = ws.cell(row=4, column=5 + c, value=sv)
                cell.font = BOLD; cell.fill = GROUP_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
            # Clear out the last col we double-wrote
            cur_last = 4 + ncols
            if cur_last >= 5 + (ncols - 1):
                # The col index 4+ncols overwrites the bottom of original layout — clear:
                ws.cell(row=3, column=cur_last).value = None
                ws.cell(row=4, column=cur_last).value = None
                ws.cell(row=3, column=cur_last).fill = PatternFill()
                ws.cell(row=4, column=cur_last).fill = PatternFill()
                ws.cell(row=3, column=cur_last).border = Border()
                ws.cell(row=4, column=cur_last).border = Border()

            # Write data rows
            r = 5
            sn = 0
            for data in all_data:
                ss = data["sheets"].get(sheet_key)
                if not ss or ss.get("missing"):
                    continue
                m = data["meta"]
                for row_data in ss["rows"]:
                    sn += 1
                    is_total = row_data["is_total"]
                    fill = TOTAL_FILL if is_total else None
                    font = BOLD if is_total else REG

                    meta_vals = [sn, m["state_code"], m["state_name"],
                                 m["gstin"], row_data["month"]]
                    for c, v in enumerate(meta_vals, 1):
                        cell = ws.cell(row=r, column=c, value=v)
                        cell.font = font
                        cell.border = BORDER_ALL
                        cell.alignment = LEFT if c in (3, 4) else CENTER
                        if c == 1:
                            cell.number_format = "0"  # integer S.No.
                        if fill: cell.fill = fill

                    for c, v in enumerate(row_data["values"], 1):
                        # Skip the first value which corresponds to column 2 in orig
                        # (already we shifted: original col 2 → output col 6 = 5 + 1)
                        if c >= ncols: break
                        cell = ws.cell(row=r, column=5 + c, value=v)
                        cell.font = font
                        cell.border = BORDER_ALL
                        if isinstance(v, (int, float)):
                            cell.alignment = RIGHT
                            cell.number_format = NUM_FMT
                        else:
                            cell.alignment = LEFT
                        if fill: cell.fill = fill
                    r += 1

            # Column widths
            ws.column_dimensions["A"].width = 7
            ws.column_dimensions["B"].width = 7
            ws.column_dimensions["C"].width = 20
            ws.column_dimensions["D"].width = 18
            ws.column_dimensions["E"].width = 10
            for c in range(6, 5 + ncols):
                ws.column_dimensions[get_column_letter(c)].width = 16
            ws.freeze_panes = "F5"
            ws.auto_filter.ref = f"A3:{get_column_letter(4 + ncols)}{r - 1}"

    # ── 3. Single-sheet long format ─────────────────────────
    if mode in ("single", "both"):
        ws = wb.create_sheet("Long Format (All)")
        ws.merge_cells("A1:N1")
        ws["A1"] = "All Sections × All States × All Months — Long Format"
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 28

        hdrs = ["S.No.", "State Code", "State Name", "GSTIN", "FY",
                "Section", "Tax Period", "Is Total?",
                "Column (Main Header)", "Sub Header (Tax Type)",
                "Column #", "Value"]
        for c, h in enumerate(hdrs, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[3].height = 32

        r = 4
        sn = 0
        for data in all_data:
            m = data["meta"]
            for sheet_key, ss in data["sheets"].items():
                if not ss or ss.get("missing"): continue
                for row_data in ss["rows"]:
                    for c, v in enumerate(row_data["values"], 1):
                        # Skip empty/blank values to reduce row count
                        if v is None or (isinstance(v, str) and not v.strip()):
                            continue
                        if isinstance(v, str) and v.strip() in ("-", "—"):
                            continue
                        sn += 1
                        main_h = ss["main_headers"][c] if c < len(ss["main_headers"]) else None
                        sub_h  = ss["sub_headers"][c]  if c < len(ss["sub_headers"]) else None
                        row_vals = [sn, m["state_code"], m["state_name"],
                                    m["gstin"], m["fy"], _SHORT_NAMES.get(sheet_key, sheet_key),
                                    row_data["month"],
                                    "Yes" if row_data["is_total"] else "",
                                    main_h, sub_h, c + 1, v]
                        for cc, vv in enumerate(row_vals, 1):
                            cell = ws.cell(row=r, column=cc, value=vv)
                            cell.font = REG; cell.border = BORDER_ALL
                            if cc == 1:
                                cell.alignment = CENTER
                                cell.number_format = "0"
                            elif cc in (2, 5, 7, 8, 11):
                                cell.alignment = CENTER
                            elif cc == 12 and isinstance(vv, (int, float)):
                                cell.alignment = RIGHT
                                cell.number_format = NUM_FMT
                            else:
                                cell.alignment = LEFT
                        r += 1

        for col, w in zip("ABCDEFGHIJKL", [7, 7, 20, 18, 8, 24, 10, 8, 50, 18, 8, 16]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A4"
        if r > 4:
            ws.auto_filter.ref = f"A3:L{r-1}"

    # ── 4. State-Summary Pivot ──────────────────────────────
    if mode == "both":
        ws = wb.create_sheet("State Summary")
        ws.merge_cells("A1:K1")
        ws["A1"] = ("State-wise Annual Totals  —  Tax Liability & ITC Summary")
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 28

        hdrs = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name",
                "Liability (GSTR-1) Total", "Liability (GSTR-3B) Total",
                "Shortfall/Excess",
                "ITC Claimed Total", "ITC (GSTR-2B) Total",
                "ITC Shortfall/Excess"]
        for c, h in enumerate(hdrs, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[3].height = 36

        r = 4
        for i, data in enumerate(all_data, 1):
            m = data["meta"]
            # Pull from "Comparison Summary" total row if available
            cs = data["sheets"].get("Comparison Summary")
            tot = None
            if cs and not cs.get("missing"):
                for row_data in cs["rows"]:
                    if row_data["is_total"]:
                        tot = row_data["values"]
                        break
            # values map (0-indexed) for Comparison Summary:
            # 0=GSTR-1, 1=GSTR-3B, 2=Shortfall/Excess, 3=cum, 4=cum%,
            # 5=ITC 3B, 6=ITC 2B, 7=Shortfall, 8=cum, 9=cum%, 10..=reversed cols
            if tot:
                def gv(i):
                    if i < len(tot):
                        v = tot[i]
                        if isinstance(v, (int, float)): return v
                    return None
                gstr1_tot = gv(0); gstr3b_tot = gv(1); short_liab = gv(2)
                itc_3b_tot = gv(5); itc_2b_tot = gv(6); short_itc = gv(7)
            else:
                gstr1_tot = gstr3b_tot = short_liab = itc_3b_tot = itc_2b_tot = short_itc = None

            row_vals = [i, m["state_code"], m["state_name"], m["gstin"], m["legal_name"],
                        gstr1_tot, gstr3b_tot, short_liab,
                        itc_3b_tot, itc_2b_tot, short_itc]
            for c, v in enumerate(row_vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = REG; cell.border = BORDER_ALL
                if c == 1:
                    cell.alignment = CENTER
                    cell.number_format = "0"
                elif c == 2:
                    cell.alignment = CENTER
                elif c in (3, 4, 5):
                    cell.alignment = LEFT
                else:
                    cell.alignment = RIGHT
                    if isinstance(v, (int, float)):
                        cell.number_format = NUM_FMT
            r += 1

        # Grand total row
        ws.cell(row=r, column=1).fill = TOTAL_FILL
        for c in range(1, 12):
            cell = ws.cell(row=r, column=c)
            cell.fill = TOTAL_FILL; cell.border = BORDER_ALL; cell.font = BOLD
        ws.cell(row=r, column=5, value="GRAND TOTAL").alignment = LEFT
        from openpyxl.utils import get_column_letter as gcl
        for c in range(6, 12):
            col_letter = gcl(c)
            ws.cell(row=r, column=c,
                    value=f"=SUM({col_letter}4:{col_letter}{r-1})")
            ws.cell(row=r, column=c).number_format = NUM_FMT
            ws.cell(row=r, column=c).alignment = RIGHT

        for col, w in zip("ABCDEFGHIJK",
                          [7, 7, 20, 18, 28, 22, 22, 18, 22, 22, 18]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A4"
        if r > 4:
            ws.auto_filter.ref = f"A3:K{r}"

    wb.save(out_path)


# ════════════════════════════════════════════════════════════════
#  GSTR-2B CONSOLIDATOR ENGINE
# ════════════════════════════════════════════════════════════════
# Consolidates GSTR-2B Excel files from multiple states/months into ONE
# file, preserving original GSTN layout (per-sheet stacked), flipping
# Credit Note rows to negative for proper netting in downstream analysis.

# Sheets to skip entirely
_G2B_SKIP_SHEETS = {"Read me"}

# Summary sheets — different structure (vertical "card" layout)
_G2B_SUMMARY_SHEETS = [
    "ITC Available", "ITC not available", "ITC Reversal", "ITC Rejected",
]

# Transactional sheets — list captures the typical full set
_G2B_TXN_SHEETS = [
    "B2B", "B2BA", "B2B-CDNR", "B2B-CDNRA",
    "ECO", "ECOA",
    "ISD", "ISDA",
    "IMPG", "IMPGA", "IMPGSEZ", "IMPGSEZA",
    "B2B (ITC Reversal)", "B2BA (ITC Reversal)",
    "B2B-DNR", "B2B-DNRA",
    "B2B(Rejected)", "B2BA(Rejected)",
    "B2B-CDNR(Rejected)", "B2B-CDNRA(Rejected)",
    "ECO(Rejected)", "ECOA(Rejected)",
    "ISD(Rejected)", "ISDA(Rejected)",
]

# Sheets where we should look for "Credit Note" in the Note type col
# and flip taxable + tax columns to negative
_G2B_CDN_SHEETS = {
    "B2B-CDNR", "B2B-CDNRA", "B2B-CDNR(Rejected)", "B2B-CDNRA(Rejected)",
}


def _g2b_extract_meta(wb, filepath):
    """Pull GSTIN, FY, period, legal name from 'Read me' sheet."""
    meta = {
        "source_file": os.path.basename(filepath),
        "gstin": "", "fy": "", "period": "",
        "state_code": "", "state_name": "",
        "legal_name": "", "trade_name": "", "gen_date": "",
    }
    if "Read me" in wb.sheetnames:
        ws = wb["Read me"]
        # Standard layout: row 4-9, label in col A, value in col C
        for r in range(1, min(ws.max_row + 1, 15)):
            label = ws.cell(r, 1).value
            value = ws.cell(r, 3).value
            if not label or value is None: continue
            label_lower = str(label).strip().lower()
            value_str = str(value).strip()
            if "financial year" in label_lower:
                meta["fy"] = value_str
            elif "tax period" in label_lower:
                meta["period"] = value_str
            elif label_lower == "gstin":
                meta["gstin"] = value_str
            elif "legal name" in label_lower:
                meta["legal_name"] = value_str
            elif "trade name" in label_lower:
                meta["trade_name"] = value_str
            elif "date of generation" in label_lower:
                meta["gen_date"] = value_str

    # Fallback to filename parsing if Read me missing
    if not meta["gstin"]:
        base = os.path.basename(filepath)
        # GSTIN structure: 2 digits + 5 letters + 4 digits + 1 letter + 1
        # alphanum + 'Z' + 1 alphanum.  Cannot use \b boundaries because
        # underscore is a word char (so '\b' fails before/after '_GSTIN_').
        # Use explicit non-alphanum lookarounds instead.
        m = re.search(
            r"(?<![A-Z0-9])(\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9])"
            r"(?![A-Z0-9])", base)
        if not m:
            # Looser fallback: any 15-char sequence starting with 2 digits.
            # Validate it has 'Z' at position 13 (standard GSTIN check).
            for cand in re.findall(r"(?<![A-Z0-9])(\d{2}[A-Z0-9]{13})(?![A-Z0-9])",
                                    base):
                if len(cand) == 15 and cand[13] == 'Z':
                    m = re.match(r"(.+)", cand)  # wrap as match
                    break
        if m: meta["gstin"] = m.group(1)
    if not meta["period"]:
        # Match 6-digit MMYYYY pattern: either at start, or between underscores.
        # Examples this must handle:
        #   1780133457797_022026_04AAACH4041D1ZC_GSTR2B_14032026.xlsx  (has _MMYYYY_)
        #   012026_06AAACH4041D1Z8_GSTR2B_14022026_1.xlsx              (starts with MMYYYY_)
        # Use the GSTIN's neighbouring digits as anchor to disambiguate from
        # other 6-digit numbers (like the 8-digit generation date).
        base = os.path.basename(filepath)
        # Best signal: 6 digits immediately followed by _GSTIN (2 digits + 13 alphanum)
        m = re.search(r"(?:^|_)(\d{6})_(?=\d{2}[A-Z0-9]{13}_)", base)
        if not m:
            # Fallback: 6 digits between underscores (old behavior)
            m = re.search(r"_(\d{6})_", base)
        if not m:
            # Last resort: any 6-digit MM-YYYY at start of basename
            m = re.match(r"^(\d{6})_", base)
        if m:
            mm, yy = m.group(1)[:2], m.group(1)[2:]
            # Validate: MM should be 01-12, YYYY should be 20XX
            if mm.isdigit() and yy.isdigit() and 1 <= int(mm) <= 12 \
               and 2000 <= int(yy) <= 2099:
                months = {"01":"January","02":"February","03":"March","04":"April",
                          "05":"May","06":"June","07":"July","08":"August",
                          "09":"September","10":"October","11":"November","12":"December"}
                meta["period"] = months.get(mm, mm)
                # Also derive FY if not set
                if not meta.get("fy"):
                    mm_int = int(mm); yy_int = int(yy)
                    if mm_int <= 3:
                        meta["fy"] = f"{yy_int - 1}-{yy_int % 100:02d}"
                    else:
                        meta["fy"] = f"{yy_int}-{(yy_int + 1) % 100:02d}"

    # Derive state from GSTIN
    if meta["gstin"]:
        meta["state_code"] = meta["gstin"][:2]
        meta["state_name"] = _STATE_CODES.get(meta["state_code"], "Unknown")

    # Build short Month abbreviation like 'Feb-26'
    if meta["period"] and meta["fy"]:
        months_abbr = {"January":"Jan","February":"Feb","March":"Mar","April":"Apr",
                       "May":"May","June":"Jun","July":"Jul","August":"Aug",
                       "September":"Sep","October":"Oct","November":"Nov","December":"Dec"}
        abbr = months_abbr.get(meta["period"], meta["period"][:3])
        # FY format "2025-26" → for Jan/Feb/Mar use "26", else "25"
        if "-" in meta["fy"]:
            yy1, yy2 = meta["fy"].split("-")[0][-2:], meta["fy"].split("-")[1]
            yy = yy2 if meta["period"] in ("January","February","March") else yy1
            meta["month_abbr"] = f"{abbr}-{yy}"
        else:
            meta["month_abbr"] = abbr
    else:
        meta["month_abbr"] = meta["period"]

    return meta


def _g2b_find_data_start_row(ws):
    """Detect first data row in a GSTR-2B transaction sheet.

    Layout varies:
      • 2-row header: row 5 = main headers, row 6 = sub-headers, row 7+ = data
      • 3-row header (amendment sheets): row 5 = group banner
        ('Original Details' / 'Revised Details'), row 6 = main headers,
        row 7 = sub-headers, row 8+ = data

    Returns the row number where actual data begins (or where it would begin
    if the sheet has no rows, so headers are still correctly placed).
    """
    HEADER_WORDS_A = {
        "gstin", "trade", "icegate", "port code", "place of supply",
        "supply attract", "isd document", "amendments to", "taxable inward",
        "debit/credit", "documents reported", "import of goods", "itc reversed",
        "itc rejected", "amendment", "credit notes", "isd credits",
        "amendments isd", "debit notes", "original details", "revised details",
        "original invoice number", "input tax",
    }
    HEADER_SUBSTR_ANY_CELL = (
        "integrated tax", "central tax", "state/ut tax", "cess(",
        "input tax distribution", "tax amount", "taxable value",
        "invoice details", "credit note/debit note", "bill of entry",
        "amount of tax", "credit note", "debit note details", "tax amount",
        "amount declared by taxpayer", "type of amendment",
    )

    def is_header_row(r):
        """Return True if this row looks like a header row anywhere across cells."""
        all_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        # Count non-empty cells
        non_empty = [v for v in all_vals if v is not None and
                     (not isinstance(v, str) or v.strip())]
        if not non_empty:
            return False  # truly blank — not a header
        # If every non-empty cell is a string and matches header keywords, it's a header
        all_str = all(isinstance(v, str) for v in non_empty)
        if not all_str:
            return False  # numeric data present → data row
        # Check if any cell contains a sub-header substring
        for v in non_empty:
            vl = v.strip().lower()
            if any(kw in vl for kw in HEADER_SUBSTR_ANY_CELL):
                return True
        # First-column starts-with check
        a = ws.cell(r, 1).value
        if a:
            a_str = str(a).strip().lower()
            if any(a_str.startswith(w) for w in HEADER_WORDS_A):
                return True
        return False

    # Walk rows 5 through 12 to find first row that is NOT a header
    for r in range(5, min(ws.max_row + 1, 13)):
        a = ws.cell(r, 1).value
        # GSTIN pattern or date pattern in col A → definitely data
        if a is not None:
            a_str = str(a).strip()
            if re.match(r"^\d{2}[A-Z0-9]{13}$", a_str):
                return r
            if re.match(r"^\d{2}[-/]\d{2}[-/]\d{4}$", a_str):
                return r
        if is_header_row(r):
            continue
        # Empty col A but other cells have content — could be data row (e.g. IMPG amendments)
        any_non_empty = any(
            ws.cell(r, c).value is not None and
            (not isinstance(ws.cell(r, c).value, str) or ws.cell(r, c).value.strip())
            for c in range(1, ws.max_column + 1)
        )
        if any_non_empty:
            return r
    # Empty sheet: figure out where data WOULD start by detecting last header row
    last_header = 4
    for r in range(5, min(ws.max_row + 1, 13)):
        if is_header_row(r):
            last_header = r
        else:
            break
    return last_header + 1


def _g2b_collect_headers(ws, data_start_row):
    """Return list of merged header strings, one per column.
    Handles both layouts:
      • 2-row header: rows [data_start-2, data_start-1] = (main, sub)
      • 3-row header: rows [data_start-3, data_start-2, data_start-1] =
                       (group banner, main, sub)
    We combine main + sub as 'Main — Sub' for display.
    """
    ncols = ws.max_column

    def _resolve(ws, row, col):
        if row < 1: return ""
        v = ws.cell(row, col).value
        if v is not None and (not isinstance(v, str) or v.strip()):
            return str(v).strip()
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                top_val = ws.cell(mr.min_row, mr.min_col).value
                if top_val is not None:
                    return str(top_val).strip()
        return ""

    # Detect 3-row header by checking if row data_start-3 contains 'Original Details'
    # or 'Revised Details' as a group banner
    has_group_banner = False
    if data_start_row >= 4:
        for c in range(1, min(ncols + 1, 10)):
            v = _resolve(ws, data_start_row - 3, c).lower()
            if v in ("original details", "revised details"):
                has_group_banner = True
                break

    if has_group_banner:
        # 3-row: group_row = data_start-3, main_row = data_start-2, sub_row = data_start-1
        main_row = data_start_row - 2
        sub_row  = data_start_row - 1
    else:
        # 2-row: main_row = data_start-2, sub_row = data_start-1
        main_row = data_start_row - 2
        sub_row  = data_start_row - 1

    headers = []
    for c in range(1, ncols + 1):
        m = _resolve(ws, main_row, c) if main_row >= 1 else ""
        s = _resolve(ws, sub_row, c)  if sub_row  >= 1 else ""
        # If sub-row value equals main-row value (merged through), use just main
        if m and s and m != s:
            combined = f"{m} — {s}"
        elif s:
            combined = s
        elif m:
            combined = m
        else:
            combined = f"Col {c}"
        headers.append(combined)
    return headers


def _g2b_parse_summary_sheet(ws):
    """Parse one of the four summary sheets. Returns list of dicts with
    section, sno, heading, gstr3b_table, igst, cgst, sgst, cess, advisory, level."""
    out = []
    if ws.max_row < 6: return out
    # Find header row by scanning for 'S.no.' in col A
    hdr_row = None
    for r in range(1, min(ws.max_row + 1, 15)):
        v = ws.cell(r, 1).value
        if v and str(v).strip().lower().startswith("s.no"):
            hdr_row = r
            break
    if hdr_row is None: return out
    data_start = hdr_row + 1

    current_part = ""        # 'Part A' / 'Part B'
    current_part_desc = ""
    current_section = ""     # Roman 'I', 'II', etc.
    current_section_desc = ""
    current_3b_ref = ""

    for r in range(data_start, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        igst = ws.cell(r, 4).value
        cgst = ws.cell(r, 5).value
        sgst = ws.cell(r, 6).value
        cess = ws.cell(r, 7).value
        advisory = ws.cell(r, 8).value

        if a is None and b is None: continue
        a_str = str(a).strip() if a else ""
        b_str = str(b).strip() if b else ""

        # Skip rows that announce a textual heading ("Credit which may be...")
        if a_str and not b_str and not isinstance(igst, (int, float)):
            # heading row with no values
            continue

        # Detect "Part A" / "Part B"
        if a_str.lower().startswith("part "):
            current_part = a_str
            current_part_desc = b_str
            continue

        # Detect roman section heading (I/II/III/IV)
        if a_str in ("I", "II", "III", "IV", "V"):
            current_section = a_str
            current_section_desc = b_str
            current_3b_ref = str(c).strip() if c else ""
            # This is a section TOTAL row — record it as level='section'
            out.append({
                "level": "Section",
                "part": current_part, "part_desc": current_part_desc,
                "section": current_section, "heading": current_section_desc,
                "gstr3b_table": current_3b_ref,
                "igst": igst if isinstance(igst, (int, float)) else None,
                "cgst": cgst if isinstance(cgst, (int, float)) else None,
                "sgst": sgst if isinstance(sgst, (int, float)) else None,
                "cess": cess if isinstance(cess, (int, float)) else None,
                "advisory": str(advisory).strip() if advisory else "",
            })
            continue

        # Detect "Details" row — sub-row under current section
        if a_str.lower() == "details":
            out.append({
                "level": "Detail",
                "part": current_part, "part_desc": current_part_desc,
                "section": current_section, "heading": b_str,
                "gstr3b_table": str(c).strip() if c else current_3b_ref,
                "igst": igst if isinstance(igst, (int, float)) else None,
                "cgst": cgst if isinstance(cgst, (int, float)) else None,
                "sgst": sgst if isinstance(sgst, (int, float)) else None,
                "cess": cess if isinstance(cess, (int, float)) else None,
                "advisory": str(advisory).strip() if advisory else "",
            })
            continue

        # Otherwise: continuation of detail rows (no S.no. label, only heading in B)
        if b_str and isinstance(igst, (int, float)):
            out.append({
                "level": "Detail",
                "part": current_part, "part_desc": current_part_desc,
                "section": current_section, "heading": b_str,
                "gstr3b_table": str(c).strip() if c else current_3b_ref,
                "igst": igst, "cgst": cgst, "sgst": sgst, "cess": cess,
                "advisory": str(advisory).strip() if advisory else "",
            })
    return out


def parse_gstr2b_file(filepath):
    """Parse one GSTR-2B Excel. Returns dict with meta + per-sheet data."""
    import openpyxl
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
    meta = _g2b_extract_meta(wb, filepath)

    result = {"meta": meta, "summary": {}, "transactions": {}}

    # Parse summary sheets
    for sn in _G2B_SUMMARY_SHEETS:
        if sn in wb.sheetnames:
            result["summary"][sn] = _g2b_parse_summary_sheet(wb[sn])

    # Parse transaction sheets
    for sn in wb.sheetnames:
        if sn in _G2B_SKIP_SHEETS or sn in _G2B_SUMMARY_SHEETS:
            continue
        ws = wb[sn]
        # Find data start
        ds = _g2b_find_data_start_row(ws)
        # Pull headers
        headers = _g2b_collect_headers(ws, ds)
        # Read data rows
        rows = []
        for r in range(ds, ws.max_row + 1):
            row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            # Skip entirely blank rows
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in row_vals):
                continue
            rows.append(row_vals)
        result["transactions"][sn] = {
            "headers": headers,
            "rows": rows,
            "data_start_row": ds,
            "n_cols": ws.max_column,
        }

    wb.close()
    return result


def _g2b_group_split_files(all_data):
    """Merge "split" GSTR-2B files (big states downloaded as multiple parts)
    into a single logical entry per (GSTIN, Period, GenerationDate).

    GSTN portal splits 2B for large GSTINs into:
       <period>_<GSTIN>_GSTR2B_<gendate>.xlsx        ← part 1 (summary + start)
       <period>_<GSTIN>_GSTR2B_<gendate>_1.xlsx      ← part 2 (continuation)
       <period>_<GSTIN>_GSTR2B_<gendate>_2.xlsx      ← part 3 (more), etc.

    Without merging, these appear as DUPLICATE state-month entries in the
    Cover/Console and transaction rows are split across two output blocks.

    Returns dict:
        merged_list  — final list of entries (one per state-month-gendate)
        split_groups — list of dicts describing each multi-part merge
        redownloads  — list of dicts where same (GSTIN, Period) had MULTIPLE
                       generation dates (likely re-downloads — kept latest)
        no_meta_files — list of source files whose (GSTIN, Period) couldn't
                        be determined, kept as standalone (review needed)
    """
    import re

    def extract_gendate(source_file):
        """Pull the 8-digit DDMMYYYY generation date out of filename."""
        if not source_file: return ""
        m = re.search(r"_GSTR2B_(\d{8})", source_file, re.IGNORECASE)
        return m.group(1) if m else ""

    # Bucket by (gstin, month_abbr) — month_abbr like "Apr-25" / "Apr-26"
    # DISAMBIGUATES THE YEAR. Using just "period" ('April') would collide
    # April-2025 with April-2026 and incorrectly treat one as a re-download
    # of the other.
    by_period = {}
    no_meta_files = []
    for d in all_data:
        m = d["meta"]
        gstin = m.get("gstin", "")
        # Prefer month_abbr (year-aware) as key, fall back to period+fy
        month_key = m.get("month_abbr", "")
        if not month_key:
            month_key = f"{m.get('period', '')}_{m.get('fy', '')}"
        if gstin and month_key.strip("_"):
            by_period.setdefault((gstin, month_key), []).append(d)
        else:
            no_meta_files.append(d)

    merged_list = []
    split_groups = []
    redownloads = []

    for (gstin, month_key), entries in by_period.items():
        # Within each (GSTIN, Period), distinguish by generation date.
        # If multiple gen dates exist, that's a re-download — keep LATEST only.
        by_gendate = {}
        for d in entries:
            gd = extract_gendate(d["meta"].get("source_file", ""))
            by_gendate.setdefault(gd, []).append(d)

        # If only ONE gen date (or only "" gen date), proceed normally
        # If multiple gen dates, pick the latest (DDMMYYYY → re-arrange to YYYYMMDD)
        if len(by_gendate) > 1 and any(by_gendate.keys()):
            def gd_sortkey(gd):
                if len(gd) != 8: return ""
                return gd[4:8] + gd[2:4] + gd[0:2]
            sorted_gds = sorted(by_gendate.keys(), key=gd_sortkey, reverse=True)
            latest_gd = sorted_gds[0]
            discarded = []
            for gd in sorted_gds[1:]:
                for d in by_gendate[gd]:
                    discarded.append(d["meta"].get("source_file", "?"))
            kept = by_gendate[latest_gd]
            redownloads.append({
                "state": kept[0]["meta"].get("state_name", "?"),
                "period": month_key,
                "gstin": gstin,
                "all_gen_dates": sorted_gds,
                "kept_gen_date": latest_gd,
                "discarded_files": discarded,
            })
            files_to_merge = kept
        else:
            # Single gen date — all files are parts of one split
            files_to_merge = entries

        if len(files_to_merge) == 1:
            merged_list.append(files_to_merge[0])
        else:
            merged = _g2b_merge_parts(files_to_merge)
            merged_list.append(merged)
            total_rows = sum(len(t.get("rows", []))
                             for t in merged["transactions"].values())
            split_groups.append({
                "state": merged["meta"].get("state_name", "?"),
                "period": month_key,
                "gstin": gstin,
                "parts": len(files_to_merge),
                "total_rows": total_rows,
                "sources": [d["meta"].get("source_file", "?")
                            for d in files_to_merge],
            })

    # Append files with no parseable meta (cannot be grouped)
    merged_list.extend(no_meta_files)

    return {
        "merged_list": merged_list,
        "split_groups": split_groups,
        "redownloads": redownloads,
        "no_meta_files": [d["meta"].get("source_file", "?")
                          for d in no_meta_files],
    }


def _g2b_merge_parts(parts):
    """Merge multiple parsed files (split parts of one state-month) into one
    combined entry. Strategy:
      • Meta: take from first part, fill blanks from later parts
      • Summary: per sheet, take first non-empty version (parts typically share
        the same summary block)
      • Transactions: per sheet, concatenate rows from ALL parts
    """
    # Sort parts so 'part1' (no _N suffix) comes first, then _1, _2, ...
    # GSTN filenames look like:  ..._GSTR2B_<DDMMYYYY>.xlsx (part 1, no suffix)
    #                       OR:  ..._GSTR2B_<DDMMYYYY>_<N>.xlsx (part N+1)
    # We must NOT confuse the 8-digit generation date for a part number.
    def sort_key(d):
        sf = (d["meta"].get("source_file") or "").lower()
        import re
        # Look for the specific GSTN suffix pattern: _GSTR2B_<8 digits>_<part>.xlsx
        m = re.search(r"_gstr2b_\d{8}_(\d+)\.xlsx?$", sf)
        if m:
            return (int(m.group(1)), sf)
        # No part suffix → this is the original/part-1
        return (0, sf)
    parts = sorted(parts, key=sort_key)

    # Meta — start from first, fill blanks from later
    merged_meta = dict(parts[0]["meta"])
    for p in parts[1:]:
        for k, v in p["meta"].items():
            if v and not merged_meta.get(k):
                merged_meta[k] = v
    # Source file list (helpful for audit)
    src_files = [p["meta"].get("source_file", "?") for p in parts]
    merged_meta["source_file"] = "  +  ".join(src_files)

    # Summary — per sheet, prefer first non-empty version
    merged_summary = {}
    for p in parts:
        for sn, summ in p.get("summary", {}).items():
            if sn not in merged_summary or not merged_summary[sn]:
                merged_summary[sn] = summ

    # Transactions — concatenate rows across parts (per sheet)
    merged_tx = {}
    for p in parts:
        for sn, tx in p.get("transactions", {}).items():
            if sn not in merged_tx:
                merged_tx[sn] = {
                    "headers": tx.get("headers", []),
                    "rows": list(tx.get("rows", [])),
                    "data_start_row": tx.get("data_start_row", 1),
                    "n_cols": tx.get("n_cols", 0),
                }
            else:
                merged_tx[sn]["rows"].extend(tx.get("rows", []))
                # Adopt headers from a later part if the earlier one had none
                if not merged_tx[sn]["headers"] and tx.get("headers"):
                    merged_tx[sn]["headers"] = tx["headers"]

    return {
        "meta": merged_meta,
        "summary": merged_summary,
        "transactions": merged_tx,
    }


def _g2b_flip_cdn_row(sheet_name, headers, row):
    """If this row is a Credit Note row, return a new row with taxable + tax columns
    flipped to negative. Otherwise return the row unchanged.

    Detection:
      - For sheets in _G2B_CDN_SHEETS: column whose sub-header == 'Note type'
        and value == 'Credit Note'  →  flip
    """
    if sheet_name not in _G2B_CDN_SHEETS:
        return row, False
    # Find Note type column
    note_type_col = None
    for i, h in enumerate(headers):
        h_lower = (h or "").lower()
        if "note type" in h_lower:
            note_type_col = i
            break
    if note_type_col is None or note_type_col >= len(row):
        return row, False
    nt = row[note_type_col]
    if not (isinstance(nt, str) and nt.strip().lower() == "credit note"):
        return row, False

    # It's a credit note — flip these column types to negative
    new_row = list(row)
    FLIP_KEYWORDS = (
        "taxable value", "integrated tax", "central tax", "state/ut tax", "cess",
        "note value",
    )
    for i, h in enumerate(headers):
        if i >= len(new_row): break
        h_lower = (h or "").lower()
        if any(kw in h_lower for kw in FLIP_KEYWORDS):
            v = new_row[i]
            if isinstance(v, (int, float)) and v > 0:
                new_row[i] = -v
    return new_row, True


def _g2b_classify_sheet(sheet_name):
    """Classify a sheet into (sheet_type, category) for the Console.

    sheet_type:  B2B / CDNR / DNR / ECO / ISD / IMPG / IMPGSEZ
    category:    Original / Amendment / Rejected / Reversal / Rejected-Amendment / Reversal-Amendment
    """
    name = sheet_name
    # Determine category from suffix/qualifier
    if "(Rejected)" in name:
        cat = "Rejected-Amendment" if name.endswith("A(Rejected)") else "Rejected"
    elif "(ITC Reversal)" in name:
        cat = "Reversal-Amendment" if "BA" in name.split("(")[0] else "Reversal"
    elif name.endswith("A") or name.endswith("RA") or name.endswith("ZA"):
        cat = "Amendment"
    else:
        cat = "Original"

    # Determine sheet type
    base = name.replace("(Rejected)", "").replace("(ITC Reversal)", "").strip()
    if base.startswith("B2B-CDNR"):
        stype = "CDNR"
    elif base.startswith("B2B-DNR"):
        stype = "DNR"
    elif base.startswith("ECO"):
        stype = "ECO"
    elif base.startswith("ISD"):
        stype = "ISD"
    elif base.startswith("IMPGSEZ"):
        stype = "IMPGSEZ"
    elif base.startswith("IMPG"):
        stype = "IMPG"
    elif base.startswith("B2B"):
        stype = "B2B"
    else:
        stype = base
    return stype, cat


def _g2b_match_col(headers, *keyword_groups):
    """Find first column index whose header matches any keyword group.
    Each group is a tuple of (must_have_all_lowercase_keywords).
    Returns column index (0-based) or None."""
    for i, h in enumerate(headers):
        hl = (h or "").lower()
        if not hl: continue
        for grp in keyword_groups:
            if all(kw in hl for kw in grp):
                return i
        # Single-keyword shortcut
    return None


def _g2b_map_to_console(sheet_name, headers, row, is_credit_note):
    """Map a single transactional row from any GSTR-2B sheet to the unified
    Console schema. Returns dict with all Console columns."""
    stype, cat = _g2b_classify_sheet(sheet_name)

    def G(*groups, default=None):
        """Find value by header keyword groups."""
        idx = None
        # Try each group in order
        for grp in groups:
            for i, h in enumerate(headers):
                hl = (h or "").lower()
                if not hl: continue
                if all(kw in hl for kw in grp):
                    idx = i
                    break
            if idx is not None: break
        if idx is None or idx >= len(row): return default
        v = row[idx]
        if v is None: return default
        if isinstance(v, str) and not v.strip(): return default
        return v

    # Counterparty GSTIN — varies by sheet
    if stype == "ISD":
        gstin_cp = G(("gstin", "isd",))
    elif stype == "ECO":
        gstin_cp = G(("gstin", "eco",))
    elif stype in ("IMPG",):
        gstin_cp = None  # imports from overseas have no supplier GSTIN
    elif stype == "IMPGSEZ":
        gstin_cp = G(("gstin", "supplier"))
    else:
        gstin_cp = G(("gstin", "supplier"))

    trade_name = G(("trade",), ("legal",))

    # Document number, date, value, type — varies by sheet
    if stype == "CDNR":
        doc_num   = G(("note number",), ("note", "number"))
        doc_type  = G(("note type",))
        doc_supply = G(("note supply",))
        doc_date  = G(("note date",))
        doc_value = G(("note value",))
    elif stype == "DNR":
        doc_num   = G(("debit note", "note number"), ("note number",))
        doc_type  = G(("note type",))
        doc_supply = G(("note supply",))
        doc_date  = G(("note date",))
        doc_value = G(("note value",))
    elif stype == "ECO":
        doc_num   = G(("document number",))
        doc_type  = G(("document type",))
        doc_supply = None
        doc_date  = G(("document date",))
        doc_value = G(("document value",))
    elif stype == "ISD":
        doc_num   = G(("isd document number",), ("document number",))
        doc_type  = G(("isd document type",), ("document type",))
        doc_supply = G(("original invoice number",))
        doc_date  = G(("isd document date",), ("document date",))
        doc_value = None
    elif stype in ("IMPG", "IMPGSEZ"):
        doc_num   = G(("bill of entry", "number"), ("number",))
        doc_type  = G(("port code",))   # use port code as identifier
        doc_supply = G(("icegate", "date"))
        doc_date  = G(("bill of entry", "date"), ("date",))
        doc_value = None
    else:  # B2B and variants
        doc_num   = G(("invoice number",))
        doc_type  = G(("invoice type",))
        doc_supply = None
        doc_date  = G(("invoice date",))
        doc_value = G(("invoice value",))

    place = G(("place of supply",))
    rcm = G(("reverse charge",))

    # Tax + taxable
    taxable = G(("taxable value",), ("taxable",))
    igst    = G(("integrated",))
    cgst    = G(("central tax",))
    sgst    = G(("state/ut",), ("state",), ("ut tax",))
    cess    = G(("cess",))

    # Other fields
    itc_avail = G(("itc availability",), ("eligibility",))
    reason    = G(("reason",))
    source    = G(("source",))
    # Tax Rate column — only B2B-family sheets actually have this column.
    # The loose ('rate',) fallback was matching headers like 'Integrated Tax
    # (Rate)' on IMPG/IMPGSEZ sheets and putting IGST amounts in the Rate
    # column. Restrict matches to actual "Applicable % of Tax Rate" / "Rate(%)"
    # patterns, and force None for IMPG/IMPGSEZ where no rate column exists.
    if stype in ("IMPG", "IMPGSEZ"):
        rate = None
    else:
        rate = G(("applicable", "tax rate"),
                 ("rate(%)",),
                 ("applicable", "%"),
                 ("rate", "%"))
    irn       = G(("irn",))
    period    = G(("gstr-1",), ("gstr-6", "period"), ("period",))
    filing_dt = G(("filing date",))

    # If we marked the row as Credit Note (already flipped), the numbers are
    # already negative — no further work.
    return {
        "sheet_type":   stype,
        "category":     cat,
        "is_cn":        "Yes" if is_credit_note else "",
        "gstin_cp":     gstin_cp,
        "trade_name":   trade_name,
        "doc_num":      doc_num,
        "doc_type":     doc_type,
        "doc_supply":   doc_supply,
        "doc_date":     doc_date,
        "doc_value":    doc_value,
        "place":        place,
        "rcm":          rcm,
        "taxable":      taxable,
        "igst":         igst,
        "cgst":         cgst,
        "sgst":         sgst,
        "cess":         cess,
        "itc_avail":    itc_avail,
        "reason":       reason,
        "source":       source,
        "rate":         rate,
        "irn":          irn,
        "period":       period,
        "filing_dt":    filing_dt,
    }


def write_consolidated_gstr2b(all_data, out_path):
    """Build a consolidated workbook from all parsed GSTR-2B files."""
    import openpyxl
    from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HDR_FILL = PatternFill("solid", start_color="1F4E78")
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    META_FILL = PatternFill("solid", start_color="DDEBF7")
    TOTAL_FILL = PatternFill("solid", start_color="FFE699")
    CDN_FILL = PatternFill("solid", start_color="FCE4D6")  # light orange — flipped rows
    SECTION_FILL = PatternFill("solid", start_color="FFF2CC")

    WHITE_B = XLFont(name="Calibri", bold=True, color="FFFFFF", size=10)
    BOLD = XLFont(name="Calibri", bold=True, size=10)
    REG = XLFont(name="Calibri", size=10)
    NEG_FONT = XLFont(name="Calibri", size=10, color="C00000")
    TITLE = XLFont(name="Calibri", bold=True, color="FFFFFF", size=14)

    thin = Side(border_style="thin", color="B4B4B4")
    BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT = Alignment(horizontal="right", vertical="center")

    NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ─── Cover sheet ──────────────────────────────────────────
    ws = wb.create_sheet("Cover")
    ws.merge_cells("A1:H1")
    ws["A1"] = "GSTR-2B — Multi-State Consolidated"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30
    ws["A3"] = "Total state-month returns consolidated:"
    ws["B3"] = len(all_data)
    ws["A3"].font = BOLD; ws["B3"].font = BOLD
    ws["A4"] = "Note:"
    ws["A4"].font = BOLD
    ws["B4"] = ("Credit Note rows (in CDNR/CDNRA sheets) are flipped to NEGATIVE for "
                "correct netting. Original GSTN format is preserved; State/GSTIN/Month "
                "columns added on the left for filtering.")
    ws["B4"].alignment = LEFT
    ws.row_dimensions[4].height = 50
    ws.merge_cells("B4:H4")

    hdrs = ["S.No.", "Month", "FY", "State Code", "State Name", "GSTIN",
            "Legal Name", "Source File"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=6, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[6].height = 22

    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        vals = [i, m["month_abbr"], m["fy"], m["state_code"], m["state_name"],
                m["gstin"], m["legal_name"], m["source_file"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=6 + i, column=c, value=v)
            cell.font = REG
            cell.alignment = LEFT if c >= 5 else CENTER
            cell.border = BORDER_ALL
            if c == 1: cell.number_format = "0"

    for col, w in zip("ABCDEFGH", [7, 9, 8, 7, 20, 20, 28, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A7"

    # ─── Coverage Matrix sheet ────────────────────────────────
    # State × Month grid showing which (GSTIN, Month) combinations were
    # consolidated. Useful for verifying complete coverage (e.g. "did
    # all states get processed for Apr-26?").
    ws = wb.create_sheet("Coverage")

    def _month_sort_key(month_abbr):
        """Convert 'Apr-25' → '202504' for chronological sorting."""
        if not month_abbr or "-" not in month_abbr: return "9999-99"
        abbr, yy = month_abbr.split("-", 1)
        month_nums = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04",
                      "May":"05","Jun":"06","Jul":"07","Aug":"08",
                      "Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
        full_year = ("20" + yy.strip()) if len(yy.strip()) == 2 else yy.strip()
        return full_year + month_nums.get(abbr, "00")

    # Build GSTIN → {month → txn_count} map
    gstin_to_months = {}
    gstin_info = {}  # gstin → (state_code, state_name, legal_name)
    months_set = set()
    for data in all_data:
        m = data["meta"]
        gstin = m.get("gstin", "")
        month = m.get("month_abbr", "")
        if not gstin or not month: continue
        if gstin not in gstin_info:
            gstin_info[gstin] = (m.get("state_code", ""),
                                  m.get("state_name", ""),
                                  m.get("legal_name", ""))
        n_txn = sum(len(t.get("rows", []))
                    for t in data.get("transactions", {}).values())
        gstin_to_months.setdefault(gstin, {})[month] = n_txn
        months_set.add(month)

    # Sort months chronologically, GSTINs by state code
    sorted_months = sorted(months_set, key=_month_sort_key)
    sorted_gstins = sorted(gstin_to_months.keys(),
                            key=lambda g: (gstin_info[g][0], g))

    # Headers (row 1: title; row 3: column headers)
    ncols_cov = 4 + len(sorted_months) + 1  # 4 meta + months + Total
    ws.merge_cells(start_row=1, start_column=1, end_row=1,
                    end_column=max(ncols_cov, 4))
    ws["A1"] = ("Coverage Matrix — State × Month  "
                "(cell value = txn rows consolidated)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 26

    hdrs = ["State Code", "State Name", "GSTIN", "Legal Name"] \
           + sorted_months + ["Total"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 30

    EMPTY_FILL = PatternFill("solid", start_color="FFD7D7")  # light red
    PRESENT_FILL = PatternFill("solid", start_color="E2EFDA") # light green

    r = 4
    month_totals = {m: 0 for m in sorted_months}
    for gstin in sorted_gstins:
        state_code, state_name, legal_name = gstin_info[gstin]
        ws.cell(row=r, column=1, value=state_code).alignment = CENTER
        ws.cell(row=r, column=2, value=state_name).alignment = LEFT
        ws.cell(row=r, column=3, value=gstin).alignment = CENTER
        ws.cell(row=r, column=4, value=legal_name).alignment = LEFT
        for c in (1, 2, 3, 4):
            ws.cell(row=r, column=c).font = REG
            ws.cell(row=r, column=c).border = BORDER_ALL

        row_total = 0
        for j, mo in enumerate(sorted_months, 5):
            n_txn = gstin_to_months[gstin].get(mo)
            cell = ws.cell(row=r, column=j, value=n_txn if n_txn else None)
            cell.border = BORDER_ALL
            cell.alignment = CENTER
            cell.font = REG
            if n_txn is None:
                cell.value = "—"
                cell.fill = EMPTY_FILL
                cell.font = XLFont(name="Calibri", color="C00000", size=10)
            else:
                cell.fill = PRESENT_FILL
                cell.number_format = "0"
                row_total += n_txn
                month_totals[mo] += n_txn
        # Row total
        cell = ws.cell(row=r, column=4 + len(sorted_months) + 1, value=row_total)
        cell.font = BOLD; cell.alignment = CENTER
        cell.border = BORDER_ALL; cell.number_format = "#,##0"
        cell.fill = TOTAL_FILL
        r += 1

    # Totals row
    ws.cell(row=r, column=1, value="TOTAL").font = BOLD
    ws.cell(row=r, column=1).fill = TOTAL_FILL
    ws.cell(row=r, column=1).alignment = CENTER
    ws.cell(row=r, column=1).border = BORDER_ALL
    for c in (2, 3, 4):
        ws.cell(row=r, column=c).fill = TOTAL_FILL
        ws.cell(row=r, column=c).border = BORDER_ALL
    grand = 0
    for j, mo in enumerate(sorted_months, 5):
        cell = ws.cell(row=r, column=j, value=month_totals[mo])
        cell.font = BOLD; cell.alignment = CENTER
        cell.fill = TOTAL_FILL; cell.border = BORDER_ALL
        cell.number_format = "#,##0"
        grand += month_totals[mo]
    cell = ws.cell(row=r, column=4 + len(sorted_months) + 1, value=grand)
    cell.font = BOLD; cell.alignment = CENTER
    cell.fill = TOTAL_FILL; cell.border = BORDER_ALL
    cell.number_format = "#,##0"

    # Column widths
    ws.column_dimensions["A"].width = 7    # State Code
    ws.column_dimensions["B"].width = 18   # State Name
    ws.column_dimensions["C"].width = 18   # GSTIN
    ws.column_dimensions["D"].width = 26   # Legal Name
    for j in range(5, 5 + len(sorted_months)):
        ws.column_dimensions[get_column_letter(j)].width = 9
    ws.column_dimensions[get_column_letter(4 + len(sorted_months) + 1)].width = 11
    ws.freeze_panes = "E4"

    # ─── ITC Summary consolidated sheet ────────────────────────
    ws = wb.create_sheet("ITC Summary (All)")
    ws.merge_cells("A1:N1")
    ws["A1"] = "ITC Summary — All Four Cards (Available / Not Available / Reversal / Rejected) — All States"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    hdrs = ["S.No.", "Month", "State Code", "State Name", "GSTIN",
            "Summary Type", "Level", "Part", "Section", "Heading",
            "GSTR-3B Table", "IGST (₹)", "CGST (₹)", "SGST (₹)", "Cess (₹)"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4
    sn = 0
    for data in all_data:
        m = data["meta"]
        for summary_name in _G2B_SUMMARY_SHEETS:
            rows = data["summary"].get(summary_name, [])
            for row in rows:
                sn += 1
                level = row.get("level", "")
                vals = [
                    sn, m["month_abbr"], m["state_code"], m["state_name"], m["gstin"],
                    summary_name, level,
                    row.get("part", ""), row.get("section", ""),
                    row.get("heading", ""), row.get("gstr3b_table", ""),
                    row.get("igst"), row.get("cgst"), row.get("sgst"), row.get("cess"),
                ]
                is_section = (level == "Section")
                fill = SECTION_FILL if is_section else None
                font = BOLD if is_section else REG
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.font = font; cell.border = BORDER_ALL
                    if fill: cell.fill = fill
                    if c == 1:
                        cell.alignment = CENTER; cell.number_format = "0"
                    elif c in (2, 3, 7, 8, 9, 11):
                        cell.alignment = CENTER
                    elif c in (4, 5, 6, 10):
                        cell.alignment = LEFT
                    elif c >= 12 and isinstance(v, (int, float)):
                        cell.alignment = RIGHT
                        cell.number_format = NUM_FMT
                r += 1

    for col, w in zip("ABCDEFGHIJKLMNO",
                      [6, 8, 6, 18, 18, 18, 8, 8, 8, 50, 12, 14, 14, 14, 12]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "F4"
    if r > 4:
        ws.auto_filter.ref = f"A3:O{r-1}"

    # ─── Console sheet — unified invoice-level across ALL txn sheets ──
    # Every row from every transaction sheet mapped to a common B2B-CDNR-like
    # schema (28 columns total: 5 meta + 23 unified data cols).
    ws = wb.create_sheet("Console")
    console_hdrs = [
        # Meta columns
        "S.No.", "Month", "State Code", "State Name", "GSTIN of Filer",
        # Sheet origin / classification
        "Sheet Type", "Category", "Is Credit Note?",
        # Counterparty
        "GSTIN of Counterparty", "Trade/Legal Name",
        # Document
        "Doc Number", "Doc Type", "Doc Supply Type", "Doc Date", "Doc Value (₹)",
        # Tax dimensions
        "Place of Supply", "Reverse Charge",
        # Values
        "Taxable Value (₹)", "IGST (₹)", "CGST (₹)", "SGST/UTGST (₹)", "Cess (₹)",
        # ITC / source / other
        "ITC Availability", "Reason", "Source", "Tax Rate", "IRN",
        "Filing Period", "Filing Date",
    ]
    ncols_console = len(console_hdrs)

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols_console)
    ws["A1"] = ("Console — Unified Invoice-level View  (every transaction row, "
                "every state, every sheet)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    # Sub-title with note
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols_console)
    ws["A2"] = ("Credit Note rows are negative & highlighted orange · "
                "Filter on Sheet Type / Category / State / Counterparty")
    ws["A2"].font = REG; ws["A2"].alignment = CENTER
    ws["A2"].fill = META_FILL

    # Header row at row 3
    for c, h in enumerate(console_hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    # Discover sheets (same logic as below)
    all_txn_sheets = []
    seen = set()
    for data in all_data:
        for sn_name in data["transactions"]:
            if sn_name not in seen:
                seen.add(sn_name)
                all_txn_sheets.append(sn_name)
    ordered_for_console = [s for s in _G2B_TXN_SHEETS if s in seen]
    ordered_for_console += [s for s in all_txn_sheets if s not in ordered_for_console]

    # Walk every transaction row from every file and every sheet
    r = 4
    sn = 0
    console_flip_count = 0
    for data in all_data:
        m = data["meta"]
        for sheet_name in ordered_for_console:
            tx = data["transactions"].get(sheet_name)
            if not tx: continue
            for row in tx["rows"]:
                if all(v is None or (isinstance(v, str) and not v.strip())
                       for v in row):
                    continue
                flipped_row, was_flipped = _g2b_flip_cdn_row(
                    sheet_name, tx["headers"], row)
                if was_flipped:
                    console_flip_count += 1

                mapped = _g2b_map_to_console(
                    sheet_name, tx["headers"], flipped_row, was_flipped)

                sn += 1
                vals = [
                    sn,
                    m["month_abbr"], m["state_code"], m["state_name"], m["gstin"],
                    mapped["sheet_type"], mapped["category"], mapped["is_cn"],
                    mapped["gstin_cp"], mapped["trade_name"],
                    mapped["doc_num"], mapped["doc_type"], mapped["doc_supply"],
                    mapped["doc_date"], mapped["doc_value"],
                    mapped["place"], mapped["rcm"],
                    mapped["taxable"], mapped["igst"], mapped["cgst"],
                    mapped["sgst"], mapped["cess"],
                    mapped["itc_avail"], mapped["reason"],
                    mapped["source"], mapped["rate"], mapped["irn"],
                    mapped["period"], mapped["filing_dt"],
                ]
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.border = BORDER_ALL
                    if was_flipped:
                        cell.fill = CDN_FILL
                    if c == 1:
                        cell.alignment = CENTER
                        cell.number_format = "0"
                        cell.font = BOLD if was_flipped else REG
                    elif c in (2, 3, 6, 7, 8):
                        cell.alignment = CENTER
                        cell.font = BOLD if (was_flipped and c == 8) else REG
                    elif isinstance(v, (int, float)) and c >= 15:
                        cell.alignment = RIGHT
                        cell.number_format = NUM_FMT
                        cell.font = NEG_FONT if (isinstance(v, (int, float)) and v < 0) else REG
                    else:
                        cell.alignment = LEFT
                        cell.font = REG
                r += 1

    # Column widths
    widths = [6, 8, 6, 18, 16,         # meta
              8, 12, 8,                # sheet/category/iscn
              18, 24,                  # gstin_cp, trade_name
              20, 12, 12, 11, 14,      # doc fields
              18, 8,                   # place, rcm
              14, 14, 14, 14, 12,      # taxable + tax
              12, 18, 11, 9, 28,       # itc/reason/source/rate/irn
              10, 12]                  # period/filing_dt
    for i, w in enumerate(widths[:ncols_console], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "I4"  # freeze meta + classification cols
    if r > 4:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols_console)}{r-1}"

    # ─── Per-transaction-sheet consolidated outputs ────────────
    # Discover all unique txn sheet names across all input files
    all_txn_sheets = []
    seen = set()
    for data in all_data:
        for sn_name in data["transactions"]:
            if sn_name not in seen:
                seen.add(sn_name)
                all_txn_sheets.append(sn_name)
    # Maintain canonical order
    ordered = [s for s in _G2B_TXN_SHEETS if s in seen]
    ordered += [s for s in all_txn_sheets if s not in ordered]

    cdn_flip_count = 0
    for sheet_name in ordered:
        # Excel sheet name length limit = 31
        out_name = sheet_name[:31]
        ws = wb.create_sheet(out_name)
        # Get canonical headers from the FIRST file that has this sheet
        canon_headers = None
        for data in all_data:
            tx = data["transactions"].get(sheet_name)
            if tx and tx["headers"]:
                canon_headers = tx["headers"]
                break
        if canon_headers is None: canon_headers = []
        n_orig_cols = len(canon_headers)

        # Title
        ncols_total = 6 + n_orig_cols  # 6 meta columns + data columns
        ws.merge_cells(start_row=1, start_column=1, end_row=1,
                       end_column=max(ncols_total, 6))
        ws["A1"] = f"{sheet_name} — All States Consolidated"
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 26

        # Header row at row 3
        meta_hdrs = ["S.No.", "Month", "State Code", "State Name", "GSTIN", "Is Credit Note?"]
        all_hdrs = meta_hdrs + canon_headers
        for c, h in enumerate(all_hdrs, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[3].height = 42

        # Write data rows
        r = 4
        sn = 0
        for data in all_data:
            m = data["meta"]
            tx = data["transactions"].get(sheet_name)
            if not tx: continue
            # If this file's headers differ in column count, align by index
            for row in tx["rows"]:
                # Skip empty rows (already filtered, but double-check)
                if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
                    continue

                # Apply credit note flipping if applicable
                flipped_row, was_flipped = _g2b_flip_cdn_row(
                    sheet_name, tx["headers"], row)
                if was_flipped:
                    cdn_flip_count += 1

                sn += 1
                meta_vals = [sn, m["month_abbr"], m["state_code"],
                             m["state_name"], m["gstin"],
                             "Yes" if was_flipped else ""]
                for c, v in enumerate(meta_vals, 1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.font = REG; cell.border = BORDER_ALL
                    cell.fill = META_FILL
                    if c == 1:
                        cell.alignment = CENTER; cell.number_format = "0"
                    elif c in (2, 3, 6):
                        cell.alignment = CENTER
                    else:
                        cell.alignment = LEFT
                    if was_flipped and c == 6:
                        cell.font = BOLD; cell.fill = CDN_FILL

                # Write data values
                for c, v in enumerate(flipped_row, 7):
                    if c - 7 >= n_orig_cols: break
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.border = BORDER_ALL
                    if was_flipped:
                        cell.fill = CDN_FILL
                    if isinstance(v, (int, float)):
                        cell.alignment = RIGHT
                        cell.number_format = NUM_FMT
                        if v < 0:
                            cell.font = NEG_FONT
                        else:
                            cell.font = REG
                    else:
                        cell.alignment = LEFT
                        cell.font = REG
                r += 1

        # Column widths — meta cols + data cols
        widths_meta = [6, 8, 6, 18, 18, 12]
        for i, w in enumerate(widths_meta, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        # Data column widths — default 18, except common known columns
        for i in range(7, 7 + n_orig_cols):
            ws.column_dimensions[get_column_letter(i)].width = 18

        ws.freeze_panes = "G4"
        if r > 4 and n_orig_cols > 0:
            ws.auto_filter.ref = f"A3:{get_column_letter(6 + n_orig_cols)}{r-1}"

    wb.save(out_path)
    return cdn_flip_count


# ════════════════════════════════════════════════════════════════
#  GSTR-9 / 9C CONSOLIDATOR ENGINE
# ════════════════════════════════════════════════════════════════
# Extracts data from GSTR-9 (Annual Return) and GSTR-9C (Reconciliation)
# PDFs. Builds a unified "Console" sheet across all input files + per-form
# detail sheets. Designed for multi-state / multi-FY consolidation.

# Per-table value-column profile.
# For each (form_type, table_no): if the standard row has fewer values than
# 5, map them to the 5 standard tax columns [Taxable, CGST, SGST, IGST, Cess].
# Default for tables with 5 values: positions 0..4 = Taxable, CGST, SGST, IGST, Cess.

_G9_NUM_RE = re.compile(r'-?[\d,]+\.\d{1,2}|-?[\d,]+')


def _g9_parse_number(s):
    """Convert '25,21,705.81' or '0.00' or '-' to float, or None."""
    if not s or s in ('-', '—'): return None
    s = s.strip().replace(',', '').replace('₹', '').strip()
    if not s: return None
    try: return float(s)
    except: return None


def _g9_extract_pdf_text(filepath):
    """Extract full text from PDF as single string."""
    import pdfplumber
    pages = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text() or "")
    return "\n".join(pages)


def _g9_clean_lines(text):
    """Strip 'FINAL' watermark single-letter lines and blanks."""
    out = []
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        # Watermark "FINAL" letters
        if len(line) == 1 and line in 'FINAL': continue
        if line == 'FINAL': continue
        out.append(line)
    return out


def _g9_merge_wrapped(lines):
    """Merge continuation lines into their parent rows.
    A new row starts when the line begins with one of:
      - 'Pt.' / 'PART'
      - Sr.No / column header marker
      - <digit>+space+<capital letter> (table title)
      - <letter><opt-digit>+space (row label)
    Otherwise the line is appended to the previous one.
    """
    merged = []
    for line in lines:
        if re.match(r'^(Pt\.|PART|Sr\.)', line):
            merged.append(line); continue
        if re.match(r'^\d+\s+[A-Z][a-zA-Z]', line):
            merged.append(line); continue
        if line.startswith(('Description', 'Central Tax', 'Tax(₹)',
                            'Form GSTR', 'See rule', 'Annual Return',
                            'Reconciliation Statement', '1 2 3', '(Amount')):
            merged.append(line); continue
        if re.match(r'^[A-Z][0-9-]?\s+', line):
            merged.append(line); continue
        if re.match(r'^(Inputs|Capital Goods|Input Services)(\s|$)', line):
            if re.search(r'\d', line):
                merged.append(line); continue
        if merged:
            merged[-1] = merged[-1] + ' ' + line
        else:
            merged.append(line)
    return merged


def _g9_split_desc_values(text):
    """Find the trailing stretch of numeric tokens. Description = rest."""
    tokens = text.split()
    if not tokens: return text, []
    is_num = [bool(_G9_NUM_RE.fullmatch(t)) for t in tokens]
    n = len(tokens)
    # Find last number token
    last_num_idx = -1
    for i in range(n - 1, -1, -1):
        if is_num[i]:
            last_num_idx = i; break
    if last_num_idx == -1:
        return text.strip(), []
    # Find first number of trailing contiguous block
    first_num_idx = last_num_idx
    for i in range(last_num_idx - 1, -1, -1):
        if is_num[i]:
            first_num_idx = i
        else:
            break
    values = [_g9_parse_number(tokens[i])
              for i in range(first_num_idx, last_num_idx + 1)]
    desc_tokens = tokens[:first_num_idx] + tokens[last_num_idx + 1:]
    return ' '.join(desc_tokens).strip(), values


# Per-(form, table) profile: how N values map to (Taxable, CGST, SGST, IGST, Cess)
# None means leave that column blank. Position is 0-indexed into the value list.
def _g9_map_values(form_type, table_no, sr, values):
    """Return dict with taxable/cgst/sgst/igst/cess based on table & sr_no."""
    out = {"taxable": None, "cgst": None, "sgst": None, "igst": None,
           "cess": None, "extra": None}
    n = len(values)

    if form_type == "GSTR-9":
        if table_no == "4":
            # 5 vals: Taxable, CGST, SGST, IGST, Cess
            # Special rows C/D (Export/SEZ on payment) have 3 vals: Taxable, IGST, Cess
            if n == 5:
                out["taxable"], out["cgst"], out["sgst"], out["igst"], out["cess"] = values
            elif n == 3:
                out["taxable"], out["igst"], out["cess"] = values
            elif n == 4:
                # Sometimes Cess is dropped → 4 vals = Taxable, CGST, SGST, IGST
                out["taxable"], out["cgst"], out["sgst"], out["igst"] = values
        elif table_no == "5":
            # Most rows: 1 val (Taxable)
            # Row N has 5 vals (Total Turnover with tax)
            if n == 1:
                out["taxable"] = values[0]
            elif n == 5:
                out["taxable"], out["cgst"], out["sgst"], out["igst"], out["cess"] = values
        elif table_no == "6":
            # 4 vals: CGST, SGST, IGST, Cess
            # Row E (Import goods): 2 vals (IGST, Cess) — no CGST/SGST for imports
            # Row F (Import services): 2 vals (IGST, Cess)
            # Row G (ISD): 4 vals (CGST, SGST, IGST, Cess)
            if n == 4:
                out["cgst"], out["sgst"], out["igst"], out["cess"] = values
            elif n == 2:
                out["igst"], out["cess"] = values
        elif table_no in ("7", "8", "12", "13", "15", "16"):
            # 4 vals: CGST, SGST, IGST, Cess
            if n == 4:
                out["cgst"], out["sgst"], out["igst"], out["cess"] = values
            elif n == 5:
                out["taxable"], out["cgst"], out["sgst"], out["igst"], out["cess"] = values
        elif table_no == "9":
            # Tax Paid: variable column count per row.
            # Typical: [Tax Payable, Cash, CGST_ITC, SGST_ITC, IGST_ITC, Cess_ITC,
            #          Total Paid, Diff]
            # But rows often omit columns (e.g. own-tax-ITC is blank).
            # Strategy: put Tax Payable in 'taxable', stash full breakdown in 'extra'.
            if n >= 1:
                out["taxable"] = values[0]
            extra_parts = []
            if n >= 2:
                extra_parts.append(f"Cash: {_g9_fmt(values[1])}")
            if n >= 3:
                # Trailing values: depends on row, but last 2 are usually
                # Total Paid + Diff. The middle are ITC components.
                # Safest: dump all as positional
                pos_labels = ["Cash", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8"]
                for idx, v in enumerate(values[1:], 1):
                    label = pos_labels[idx - 1] if idx - 1 < len(pos_labels) else f"V{idx+1}"
                # Better: identify the Total Paid as second-to-last,
                # Diff as last
                if n >= 4:
                    out["extra"] = (f"Cash: {_g9_fmt(values[1])}; "
                                    f"Total Paid: {_g9_fmt(values[-2])}; "
                                    f"Diff: {_g9_fmt(values[-1])}")
                    # Map middle values to CGST/SGST/IGST/Cess based on row
                    middle = values[2:-2]
                    if len(middle) >= 4:
                        out["cgst"], out["sgst"], out["igst"], out["cess"] = middle[:4]
                    elif len(middle) == 3:
                        out["cgst"], out["sgst"], out["igst"] = middle
                    elif len(middle) == 2:
                        out["cgst"], out["sgst"] = middle
        elif table_no in ("10", "11"):
            # 5 vals: Taxable, CGST, SGST, IGST, Cess
            if n == 5:
                out["taxable"], out["cgst"], out["sgst"], out["igst"], out["cess"] = values
        elif table_no in ("14",):
            # Differential tax
            if n >= 1: out["taxable"] = values[0]
            if n >= 2: out["extra"] = "Paid: " + _g9_fmt(values[1])
        elif table_no == "19":
            # Late fee — Payable, Paid
            if n >= 1: out["taxable"] = values[0]
            if n >= 2: out["extra"] = "Paid: " + _g9_fmt(values[1])
        else:
            # Default: 5-val mapping
            if n == 5:
                out["taxable"], out["cgst"], out["sgst"], out["igst"], out["cess"] = values
            elif n == 4:
                out["cgst"], out["sgst"], out["igst"], out["cess"] = values
            elif n == 1:
                out["taxable"] = values[0]

    elif form_type == "GSTR-9C":
        if table_no in ("5", "7", "12", "14", "16"):
            # Reconciliation tables: 1 value (Amount)
            if n == 1:
                out["taxable"] = values[0]
            elif n == 3:
                # Table 14: Value, Total ITC, Eligible ITC
                out["taxable"] = values[0]
                out["cgst"] = values[1]  # Total ITC
                out["sgst"] = values[2]  # Eligible ITC
                out["extra"] = "Col2=Total ITC; Col3=Eligible ITC"
        elif table_no in ("9", "11", "17"):
            # Rate-wise: Taxable, CGST, SGST, IGST, Cess
            if n == 5:
                out["taxable"], out["cgst"], out["sgst"], out["igst"], out["cess"] = values
            elif n == 4:
                # Some rows only have 4 (no Cess)
                out["taxable"], out["cgst"], out["sgst"], out["igst"] = values
            elif n == 3:
                # G 28%: Taxable, IGST, Cess only (no CGST/SGST)
                out["taxable"], out["igst"], out["cess"] = values
            elif n == 2:
                out["taxable"], out["igst"] = values

    return out


def _g9_fmt(v):
    """Format number for display."""
    if v is None: return ""
    if isinstance(v, float) and v == int(v):
        return f"{int(v):,}"
    return f"{v:,.2f}"


def _g9_extract_meta(text, filepath):
    """Extract GSTIN, FY, ARN, etc. from header text."""
    head = ' '.join(text.split('\n')[:30])
    meta = {"source_file": os.path.basename(filepath)}
    m = re.search(r'Financial Year\s+(\d{4}-\d{2})', head)
    meta["fy"] = m.group(1) if m else ""
    m = re.search(r'\b(\d{2}[A-Z0-9]{13})\b', head)
    meta["gstin"] = m.group(1) if m else ""
    meta["state_code"] = meta["gstin"][:2] if meta["gstin"] else ""
    meta["state_name"] = _STATE_CODES.get(meta["state_code"], "Unknown")
    m = re.search(r'Legal [Nn]ame[^:]*?(?:registered person)?\s+([A-Z][^\n0-9]+?)(?:\s+3\(b\)|\s+Trade|$)',
                  head)
    meta["legal_name"] = m.group(1).strip() if m else ""
    m = re.search(r'Trade [Nn]ame[^:]*?\s+([A-Z][^\n]+?)(?:\s+3\(c\)|\s+ARN|$)', head)
    meta["trade_name"] = m.group(1).strip() if m else ""
    m = re.search(r'ARN\s+([A-Z0-9]+)', head)
    meta["arn"] = m.group(1) if m else ""
    # Filing date: GSTR-9 uses "Date of Filing", GSTR-9C uses "ARN Date"
    m = re.search(r'(?:Date of Filing|ARN Date)\s+(\d{2}-\d{2}-\d{4})', head)
    meta["filing_date"] = m.group(1) if m else ""

    # Detect form type
    if 'GSTR-9C' in head or 'Reconciliation Statement' in head:
        meta["form_type"] = "GSTR-9C"
    else:
        meta["form_type"] = "GSTR-9"
    return meta


def _g9_parse_table6_with_subtypes(label, rest):
    """Handle GSTR-9 Table 6 rows that have Inputs/Capital Goods/Input Services
    sub-rows. Returns a list of row dicts (one per sub-type, or one if none).
    """
    parts = re.split(r'\s+(Inputs|Capital Goods|Input Services)\s+', rest)
    if len(parts) == 1:
        # No sub-types
        desc, values = _g9_split_desc_values(parts[0])
        return [{"sr": label, "desc": desc, "subtype": "", "values": values}]

    main_desc = parts[0].strip()
    out = []
    for i in range(1, len(parts), 2):
        subtype = parts[i]
        rest_str = parts[i + 1] if i + 1 < len(parts) else ""
        # Extract leading numbers; trailing tokens become description continuation
        toks = rest_str.split()
        values = []
        trail_desc = []
        in_values = True
        for tok in toks:
            if in_values and _G9_NUM_RE.fullmatch(tok):
                values.append(_g9_parse_number(tok))
            else:
                in_values = False
                trail_desc.append(tok)
        if trail_desc:
            main_desc = (main_desc + ' ' + ' '.join(trail_desc)).strip()
        out.append({"sr": label, "desc": main_desc,
                    "subtype": subtype, "values": values})
    # Final desc applies to all
    for r in out: r["desc"] = main_desc
    return out


def parse_gstr9_pdf(filepath):
    """Parse a GSTR-9 OR GSTR-9C PDF. Returns dict with meta + rows."""
    text = _g9_extract_pdf_text(filepath)
    lines = _g9_clean_lines(text)
    merged = _g9_merge_wrapped(lines)

    meta = _g9_extract_meta(text, filepath)
    form_type = meta["form_type"]

    rows = []
    current_pt = ""
    current_table = ""
    current_table_title = ""

    for line in merged:
        # Skip column headers / form headers
        if re.match(r'^(Sr\.|Description|Central Tax|Tax\(|1 2 3'
                    r'|Form GSTR|See rule|Annual Return|Reconciliation Statement'
                    r'|PART|\(Amount|Verification)', line):
            continue
        # Skip empty/short
        if len(line) < 3: continue

        # Pt header — capture Roman + descriptive text
        m = re.match(r'^Pt\.?\s*(I+|IV|V|VI|VII)\b\s*(.*)?', line)
        if m:
            current_pt = m.group(1)
            continue

        # Table header: digit + space + capital letter + alpha
        m = re.match(r'^(\d+)\s+([A-Z][a-zA-Z].+)', line)
        if m:
            current_table = m.group(1)
            current_table_title = m.group(2)[:120]
            continue

        # Row label: A, B, C, A1, A2, K-2 etc.
        m = re.match(r'^([A-Z][0-9]?(?:-[0-9])?)\s+(.+)', line)
        if m:
            label = m.group(1)
            rest = m.group(2)

            # Skip if rest looks like another section header word
            if rest in ('Nature of Supplies', 'Description', 'Details', 'Type'):
                continue

            # Table 6 sub-type splitting
            if form_type == "GSTR-9" and current_table == "6":
                sub_rows = _g9_parse_table6_with_subtypes(label, rest)
                for sr in sub_rows:
                    mapped = _g9_map_values(form_type, current_table,
                                            sr["sr"], sr["values"])
                    rows.append({
                        "pt": current_pt,
                        "table": current_table,
                        "table_title": current_table_title,
                        "sr": sr["sr"], "desc": sr["desc"],
                        "subtype": sr["subtype"],
                        **mapped,
                        "raw_values": sr["values"],
                    })
                continue

            # Regular row parse
            desc, values = _g9_split_desc_values(rest)
            mapped = _g9_map_values(form_type, current_table, label, values)
            rows.append({
                "pt": current_pt,
                "table": current_table,
                "table_title": current_table_title,
                "sr": label, "desc": desc,
                "subtype": "",
                **mapped,
                "raw_values": values,
            })

    return {"meta": meta, "rows": rows}


def parse_gstr9_or_9c_pdf(filepath):
    """Auto-detect and parse a GSTR-9 or GSTR-9C PDF."""
    return parse_gstr9_pdf(filepath)


def write_consolidated_gstr9_9c(all_data, out_path):
    """Build consolidated workbook for GSTR-9 + GSTR-9C data.
    all_data is a list of dicts from parse_gstr9_pdf(), each with
    meta.form_type indicating which form."""
    import openpyxl
    from openpyxl.styles import Font as XF, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Style constants
    HDR_FILL = PatternFill("solid", start_color="1F4E78")
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    META_FILL = PatternFill("solid", start_color="DDEBF7")
    GROUP_FILL_9 = PatternFill("solid", start_color="E2EFDA")   # light green
    GROUP_FILL_9C = PatternFill("solid", start_color="FFF2CC")  # light yellow

    WHITE_B = XF(name="Calibri", bold=True, color="FFFFFF", size=10)
    BOLD = XF(name="Calibri", bold=True, size=10)
    REG = XF(name="Calibri", size=10)
    NEG_FONT = XF(name="Calibri", size=10, color="C00000")
    TITLE = XF(name="Calibri", bold=True, color="FFFFFF", size=14)

    thin = Side(border_style="thin", color="B4B4B4")
    BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT = Alignment(horizontal="right", vertical="center")
    NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Split into GSTR-9 and GSTR-9C
    data_9 = [d for d in all_data if d["meta"].get("form_type") == "GSTR-9"]
    data_9c = [d for d in all_data if d["meta"].get("form_type") == "GSTR-9C"]

    # ─── Cover sheet ──────────────────────────────────────────
    ws = wb.create_sheet("Cover")
    ws.merge_cells("A1:I1")
    ws["A1"] = "GSTR-9 / GSTR-9C — Multi-State Consolidated"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30
    ws["A3"] = "GSTR-9 returns:"
    ws["B3"] = len(data_9)
    ws["A4"] = "GSTR-9C statements:"
    ws["B4"] = len(data_9c)
    ws["A3"].font = BOLD; ws["A4"].font = BOLD
    ws["B3"].font = BOLD; ws["B4"].font = BOLD

    hdrs = ["S.No.", "Form Type", "FY", "State Code", "State Name", "GSTIN",
            "Legal Name", "Filing Date", "Source File"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=6, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[6].height = 22

    sn = 0
    for data in (data_9 + data_9c):
        sn += 1; m = data["meta"]
        vals = [sn, m["form_type"], m["fy"], m["state_code"], m["state_name"],
                m["gstin"], m["legal_name"], m["filing_date"], m["source_file"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=6 + sn, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            cell.alignment = LEFT if c >= 5 else CENTER
            if c == 1: cell.number_format = "0"
    for col, w in zip("ABCDEFGHI", [6, 10, 8, 7, 20, 20, 28, 12, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A7"

    # ─── Helper to write a Console sheet ──────────────────────
    def _write_console(sheet_name, data_list, form_type, group_fill):
        if not data_list: return
        ws = wb.create_sheet(sheet_name)
        cols = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name",
                "FY", "Filing Date", "Form Type",
                "Pt", "Table", "Sr.No",
                "Description", "Sub-Type",
                "Taxable Value (₹)", "CGST (₹)", "SGST/UTGST (₹)",
                "IGST (₹)", "Cess (₹)", "Extra Info"]
        ncols = len(cols)

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        ws["A1"] = (f"{form_type} Console — every row from every filing  "
                    f"({len(data_list)} files)")
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 28

        for c, h in enumerate(cols, 1):
            cell = ws.cell(row=3, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[3].height = 38

        r = 4; sn = 0
        for data in data_list:
            m = data["meta"]
            for row in data["rows"]:
                sn += 1
                vals = [
                    sn, m["state_code"], m["state_name"], m["gstin"],
                    m["legal_name"], m["fy"], m["filing_date"], m["form_type"],
                    row["pt"], row["table"], row["sr"],
                    row["desc"], row["subtype"],
                    row.get("taxable"), row.get("cgst"), row.get("sgst"),
                    row.get("igst"), row.get("cess"), row.get("extra"),
                ]
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.border = BORDER_ALL
                    if c == 1:
                        cell.alignment = CENTER; cell.number_format = "0"
                        cell.font = REG
                    elif c in (2, 3, 8, 9, 10, 11):
                        cell.alignment = CENTER; cell.font = REG
                    elif isinstance(v, (int, float)) and c >= 14:
                        cell.alignment = RIGHT
                        cell.number_format = NUM_FMT
                        cell.font = NEG_FONT if v < 0 else REG
                    else:
                        cell.alignment = LEFT; cell.font = REG
                r += 1

        widths = [6, 6, 18, 18, 26,
                  10, 12, 10,
                  6, 6, 6,
                  50, 14,
                  18, 16, 16, 16, 12, 36]
        for i, w in enumerate(widths[:ncols], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "L4"
        if r > 4:
            ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{r-1}"

    # ─── GSTR-9 Console ───────────────────────────────────────
    if data_9:
        _write_console("GSTR-9 Console", data_9, "GSTR-9", GROUP_FILL_9)

    # ─── GSTR-9C Console ──────────────────────────────────────
    if data_9c:
        _write_console("GSTR-9C Console", data_9c, "GSTR-9C", GROUP_FILL_9C)

    wb.save(out_path)
    return len(data_9), len(data_9c)


# ════════════════════════════════════════════════════════════════
#  ECRRS — Electronic Credit Reversal and Reclaimed Statement
# ════════════════════════════════════════════════════════════════
# Parses GSTN portal CSV "Electronic Credit Reversal and Re-claimed
# Statement" exports. Multi-GSTIN/state consolidation with:
#   • Cover sheet
#   • Console sheet — every monthly row across all GSTINs, long format
#   • Summary sheet — period totals per GSTIN with health indicators
#   • Per-GSTIN detail sheets — original portal layout preserved


def _ecrrs_parse_num(s):
    """Convert '54,144,976' / '0' / '-' / '' → float or None."""
    if s is None: return None
    s = str(s).strip().replace(',', '').replace('₹', '').strip()
    if not s or s in ('-', '—', 'NA', 'N/A'): return None
    try: return float(s)
    except: return None


def parse_ecrrs_csv(filepath):
    """Parse one ECRRS CSV file. Returns dict with meta + rows."""
    import csv
    raw_rows = []
    with open(filepath, encoding='utf-8-sig', errors='replace') as f:
        reader = csv.reader(f)
        for row in reader:
            raw_rows.append(row)

    meta = {
        "source_file": os.path.basename(filepath),
        "gstin": "", "legal_name": "", "from_date": "", "to_date": "",
        "state_code": "", "state_name": "",
    }
    # First few rows have meta in column-pair format
    for r in raw_rows[:10]:
        if len(r) < 5: continue
        label = (r[3] or "").strip() if len(r) > 3 else ""
        value = (r[4] or "").strip() if len(r) > 4 else ""
        if not label or not value: continue
        if label == "GSTIN": meta["gstin"] = value
        elif label == "Legal Name": meta["legal_name"] = value
        elif label == "From": meta["from_date"] = value
        elif label == "To": meta["to_date"] = value
    if meta["gstin"]:
        meta["state_code"] = meta["gstin"][:2]
        meta["state_name"] = _STATE_CODES.get(meta["state_code"], "Unknown")

    # Find header rows
    # Row with 'S.No.' is the main header; the next row has tax-type sub-headers
    hdr_idx = None
    for i, r in enumerate(raw_rows):
        if r and r[0] and r[0].strip().startswith("S.No"):
            hdr_idx = i; break
    if hdr_idx is None:
        return {"meta": meta, "rows": []}

    data_start = hdr_idx + 2
    rows = []
    for r in raw_rows[data_start:]:
        if not r or not r[0] or not str(r[0]).strip():
            continue
        # Skip rows with fewer than 21 columns (data has 21 columns including S.No)
        # but allow Opening/Closing Balance rows with dashes
        def col(i):
            return r[i].strip() if i < len(r) and r[i] else ""

        sr_no = col(0)
        date = col(1)
        ref_no = col(2)
        period = col(3)
        desc = col(4)

        # ITC Claimed (Table 4A(5)) — IGST, CGST, SGST, Cess
        claimed_igst = _ecrrs_parse_num(col(5))
        claimed_cgst = _ecrrs_parse_num(col(6))
        claimed_sgst = _ecrrs_parse_num(col(7))
        claimed_cess = _ecrrs_parse_num(col(8))

        # ITC Reversal (Table 4B(2))
        rev_igst = _ecrrs_parse_num(col(9))
        rev_cgst = _ecrrs_parse_num(col(10))
        rev_sgst = _ecrrs_parse_num(col(11))
        rev_cess = _ecrrs_parse_num(col(12))

        # ITC Reclaimed (Table 4D(1))
        rec_igst = _ecrrs_parse_num(col(13))
        rec_cgst = _ecrrs_parse_num(col(14))
        rec_sgst = _ecrrs_parse_num(col(15))
        rec_cess = _ecrrs_parse_num(col(16))

        # Closing Balance
        clo_igst = _ecrrs_parse_num(col(17))
        clo_cgst = _ecrrs_parse_num(col(18))
        clo_sgst = _ecrrs_parse_num(col(19))
        clo_cess = _ecrrs_parse_num(col(20))

        # Classify row type
        if "Opening Balance" in desc:
            row_type = "Opening"
        elif "Closing Balance" in desc:
            row_type = "Closing"
        else:
            row_type = "Monthly"

        rows.append({
            "sr_no": sr_no, "date": date, "ref_no": ref_no,
            "period": period, "desc": desc, "row_type": row_type,
            # Claimed
            "claimed_igst": claimed_igst, "claimed_cgst": claimed_cgst,
            "claimed_sgst": claimed_sgst, "claimed_cess": claimed_cess,
            # Reversal
            "rev_igst": rev_igst, "rev_cgst": rev_cgst,
            "rev_sgst": rev_sgst, "rev_cess": rev_cess,
            # Reclaimed
            "rec_igst": rec_igst, "rec_cgst": rec_cgst,
            "rec_sgst": rec_sgst, "rec_cess": rec_cess,
            # Closing
            "clo_igst": clo_igst, "clo_cgst": clo_cgst,
            "clo_sgst": clo_sgst, "clo_cess": clo_cess,
        })

    return {"meta": meta, "rows": rows}


def _ecrrs_dedup_by_gstin(all_data):
    """Combine multiple ECRRS files for the same GSTIN and remove duplicate
    rows by Reference No.

    Real-world need: users often download overlapping periods (e.g. Apr-25 to
    Mar-26 AND 20-Mar-26 to 20-Apr-26). The Feb-26 GSTR-3B filing (20-Mar-26)
    appears in BOTH. Reference No is GSTN's unique transaction ID, so we use
    that as the dedup key.

    Returns (combined_list, dedup_stats).
    `combined_list` has one entry per GSTIN with all unique monthly rows
    sorted by date. `dedup_stats` is a list of dicts logging what was removed.
    """
    from datetime import datetime

    # Group by GSTIN
    by_gstin = {}
    for data in all_data:
        gstin = data["meta"].get("gstin", "")
        if not gstin: continue
        by_gstin.setdefault(gstin, []).append(data)

    combined = []
    stats = []

    def parse_date(s):
        """Parse '20/03/2026' or '19/04/2025' to a sortable datetime, or None."""
        if not s or s in ('-', '—'): return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try: return datetime.strptime(s.strip(), fmt)
            except: continue
        return None

    for gstin, files in by_gstin.items():
        # Sort files by from_date so earliest comes first
        files_sorted = sorted(files,
                              key=lambda d: parse_date(d["meta"].get("from_date", ""))
                                            or datetime.min)

        # Collect all monthly rows, keep unique by reference number
        # If duplicates found, prefer the row from the earliest file
        # (which is usually the more complete year statement)
        seen_refs = {}              # ref_no → row
        all_monthly = []
        duplicates_skipped = 0
        dup_examples = []
        opening_row = None          # from earliest file
        closing_row = None          # from latest file

        for data in files_sorted:
            for row in data["rows"]:
                rt = row["row_type"]
                if rt == "Opening":
                    if opening_row is None:
                        opening_row = dict(row)  # take FIRST opening
                elif rt == "Closing":
                    closing_row = dict(row)      # take LAST closing
                else:
                    ref = (row.get("ref_no") or "").strip()
                    # Fallback dedup key for rows missing ref_no
                    if not ref or ref == "-":
                        ref = (f"{row.get('date','')}_{row.get('period','')}"
                               f"_{row.get('desc','')[:20]}")
                    if ref in seen_refs:
                        duplicates_skipped += 1
                        if len(dup_examples) < 3:
                            dup_examples.append(
                                f"{row.get('period','?')} "
                                f"(Ref: {row.get('ref_no','?')[:20]})")
                        continue
                    seen_refs[ref] = row
                    all_monthly.append(dict(row))

        # Sort monthly rows by date so they appear chronologically
        all_monthly.sort(key=lambda r: parse_date(r.get("date", ""))
                                       or datetime.min)

        # Re-number Sr.No
        rows_out = []
        sn = 0
        if opening_row:
            sn += 1
            opening_row["sr_no"] = str(sn)
            rows_out.append(opening_row)
        for row in all_monthly:
            sn += 1
            row["sr_no"] = str(sn)
            rows_out.append(row)
        if closing_row:
            sn += 1
            closing_row["sr_no"] = str(sn)
            rows_out.append(closing_row)

        # Build the merged meta — combine periods of all source files
        first_meta = files_sorted[0]["meta"]
        last_meta = files_sorted[-1]["meta"]
        merged_meta = dict(first_meta)
        merged_meta["from_date"] = first_meta.get("from_date", "")
        merged_meta["to_date"] = last_meta.get("to_date", "")
        if len(files_sorted) > 1:
            merged_meta["source_file"] = (
                f"{len(files_sorted)} files merged: "
                + ", ".join(d["meta"].get("source_file", "?")
                            for d in files_sorted))

        combined.append({"meta": merged_meta, "rows": rows_out})

        if len(files_sorted) > 1 or duplicates_skipped > 0:
            stats.append({
                "gstin": gstin,
                "files_merged": len(files_sorted),
                "duplicates_skipped": duplicates_skipped,
                "duplicate_examples": dup_examples,
                "unique_monthly": len(all_monthly),
            })

    return combined, stats


def write_consolidated_ecrrs(all_data, out_path):
    """Build consolidated workbook for ECRRS data from multiple GSTINs."""
    import openpyxl
    from openpyxl.styles import (Font as XF, PatternFill, Alignment,
                                  Border, Side)
    from openpyxl.utils import get_column_letter

    # Step 1: Dedup by GSTIN + Reference No (handles overlapping period downloads)
    all_data, dedup_stats = _ecrrs_dedup_by_gstin(all_data)

    HDR_FILL = PatternFill("solid", start_color="1F4E78")
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    GRP_FILL_CLAIM = PatternFill("solid", start_color="E2EFDA")   # green
    GRP_FILL_REV = PatternFill("solid", start_color="FCE4D6")     # orange
    GRP_FILL_REC = PatternFill("solid", start_color="DDEBF7")     # blue
    GRP_FILL_BAL = PatternFill("solid", start_color="FFF2CC")     # yellow
    META_FILL = PatternFill("solid", start_color="F2F2F2")
    OPEN_FILL = PatternFill("solid", start_color="E7E6E6")
    CLOSE_FILL = PatternFill("solid", start_color="FFE699")
    GOOD_FILL = PatternFill("solid", start_color="C6EFCE")
    WARN_FILL = PatternFill("solid", start_color="FFEB9C")
    BAD_FILL = PatternFill("solid", start_color="FFC7CE")

    WHITE_B = XF(name="Calibri", bold=True, color="FFFFFF", size=10)
    BOLD = XF(name="Calibri", bold=True, size=10)
    REG = XF(name="Calibri", size=10)
    NEG_FONT = XF(name="Calibri", size=10, color="C00000")
    TITLE = XF(name="Calibri", bold=True, color="FFFFFF", size=14)
    SUBTITLE = XF(name="Calibri", italic=True, size=10, color="404040")

    thin = Side(border_style="thin", color="B4B4B4")
    BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT = Alignment(horizontal="right", vertical="center")
    NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ─── 1. Cover ─────────────────────────────────────────────
    ws = wb.create_sheet("Cover")
    ws.merge_cells("A1:H1")
    ws["A1"] = "Electronic Credit Reversal & Re-claimed Statement — Consolidated"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:H2")
    ws["A2"] = ("Multi-GSTIN consolidation · ITC Claimed / Reversed / Reclaimed "
                "/ Closing Balance · long-format Console + per-GSTIN sheets")
    ws["A2"].font = SUBTITLE; ws["A2"].alignment = CENTER

    hdrs = ["S.No.", "GSTIN", "State Code", "State Name", "Legal Name",
            "Period From", "Period To", "Source File"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[4].height = 22

    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        vals = [i, m["gstin"], m["state_code"], m["state_name"],
                m["legal_name"], m["from_date"], m["to_date"],
                m["source_file"]]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=4 + i, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            cell.alignment = LEFT if c >= 4 else CENTER
            if c == 1: cell.number_format = "0"

    for col, w in zip("ABCDEFGH", [6, 18, 7, 18, 28, 12, 12, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ─── 2. Console (long format) ─────────────────────────────
    ws = wb.create_sheet("Console")
    cols = ["S.No.", "GSTIN", "State Code", "State Name", "Legal Name",
            "Row Type", "Sr.No.", "Date", "Reference No.", "Return Period",
            "Description",
            # 4 groups × 4 tax types = 16 numeric cols
            "Claimed IGST", "Claimed CGST", "Claimed SGST/UTGST", "Claimed Cess",
            "Reversal IGST", "Reversal CGST", "Reversal SGST/UTGST", "Reversal Cess",
            "Reclaimed IGST", "Reclaimed CGST", "Reclaimed SGST/UTGST", "Reclaimed Cess",
            "Closing IGST", "Closing CGST", "Closing SGST/UTGST", "Closing Cess"]
    ncols = len(cols)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws["A1"] = ("Console — every row, every GSTIN, every month  "
                "(filter by State / GSTIN / Period / Row Type)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    # Group header row 2 — coloured banners
    def grp(c_start, c_end, label, fill):
        ws.merge_cells(start_row=2, start_column=c_start, end_row=2,
                       end_column=c_end)
        cell = ws.cell(row=2, column=c_start, value=label)
        cell.font = BOLD; cell.fill = fill
        cell.alignment = CENTER; cell.border = BORDER_ALL

    grp(1, 11, "Metadata", META_FILL)
    grp(12, 15, "ITC Claimed  [Table 4A(5)]", GRP_FILL_CLAIM)
    grp(16, 19, "ITC Reversal  [Table 4B(2)]", GRP_FILL_REV)
    grp(20, 23, "ITC Reclaimed  [Table 4D(1)]", GRP_FILL_REC)
    grp(24, 27, "Closing Balance", GRP_FILL_BAL)
    ws.row_dimensions[2].height = 22

    # Column header row 3
    for c, h in enumerate(cols, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4; sn = 0
    for data in all_data:
        m = data["meta"]
        for row in data["rows"]:
            sn += 1
            vals = [
                sn, m["gstin"], m["state_code"], m["state_name"],
                m["legal_name"],
                row["row_type"], row["sr_no"], row["date"],
                row["ref_no"], row["period"], row["desc"],
                row["claimed_igst"], row["claimed_cgst"],
                row["claimed_sgst"], row["claimed_cess"],
                row["rev_igst"], row["rev_cgst"],
                row["rev_sgst"], row["rev_cess"],
                row["rec_igst"], row["rec_cgst"],
                row["rec_sgst"], row["rec_cess"],
                row["clo_igst"], row["clo_cgst"],
                row["clo_sgst"], row["clo_cess"],
            ]
            row_bold = row["row_type"] in ("Opening", "Closing")
            row_fill = (OPEN_FILL if row["row_type"] == "Opening"
                        else CLOSE_FILL if row["row_type"] == "Closing"
                        else None)
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER_ALL
                if row_fill: cell.fill = row_fill
                if c == 1:
                    cell.alignment = CENTER
                    cell.number_format = "0"
                    cell.font = BOLD if row_bold else REG
                elif c in (2, 3, 6, 7, 8, 10):
                    cell.alignment = CENTER
                    cell.font = BOLD if row_bold else REG
                elif c in (4, 5, 9, 11):
                    cell.alignment = LEFT
                    cell.font = BOLD if row_bold else REG
                elif isinstance(v, (int, float)) and c >= 12:
                    cell.alignment = RIGHT
                    cell.number_format = NUM_FMT
                    if v < 0:
                        cell.font = NEG_FONT
                    else:
                        cell.font = BOLD if row_bold else REG
                else:
                    cell.alignment = LEFT
                    cell.font = BOLD if row_bold else REG
            r += 1

    widths = [6, 18, 6, 16, 22,
              9, 6, 11, 18, 10, 26,
              14, 14, 16, 11,
              14, 14, 16, 11,
              14, 14, 16, 11,
              14, 14, 16, 11]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "L4"
    if r > 4:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{r-1}"

    # ─── 3. Summary (per GSTIN totals + health) ───────────────
    ws = wb.create_sheet("Summary")
    ws.merge_cells("A1:R1")
    ws["A1"] = ("Summary — Total Claimed / Reversed / Reclaimed / Closing  "
                "per GSTIN  (with reversal-reclaim balance check)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    # Group banner row 2
    def grp2(c_start, c_end, label, fill):
        ws.merge_cells(start_row=2, start_column=c_start, end_row=2,
                       end_column=c_end)
        cell = ws.cell(row=2, column=c_start, value=label)
        cell.font = BOLD; cell.fill = fill
        cell.alignment = CENTER; cell.border = BORDER_ALL

    grp2(1, 5, "GSTIN Details", META_FILL)
    grp2(6, 8, "Period & Counts", META_FILL)
    grp2(9, 11, "Total Reversal (sum of Monthly Reversal)", GRP_FILL_REV)
    grp2(12, 14, "Total Reclaimed (sum of Monthly Reclaimed)", GRP_FILL_REC)
    grp2(15, 17, "Closing Balance (latest)", GRP_FILL_BAL)
    ws.merge_cells(start_row=2, start_column=18, end_row=2, end_column=18)
    cell = ws.cell(row=2, column=18, value="Status")
    cell.font = BOLD; cell.fill = META_FILL
    cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[2].height = 22

    sum_hdrs = ["S.No.", "GSTIN", "State", "Legal Name", "FY/Period",
                "From", "To", "Months",
                "Reversal IGST", "Reversal CGST+SGST", "Reversal Cess",
                "Reclaimed IGST", "Reclaimed CGST+SGST", "Reclaimed Cess",
                "Closing IGST", "Closing CGST+SGST", "Closing Cess",
                "Health"]
    for c, h in enumerate(sum_hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4
    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        monthly = [row for row in data["rows"] if row["row_type"] == "Monthly"]
        closing = next((row for row in data["rows"]
                        if row["row_type"] == "Closing"), None)

        def sum_field(field):
            return sum((row[field] or 0) for row in monthly)

        rev_i = sum_field("rev_igst")
        rev_c = sum_field("rev_cgst") + sum_field("rev_sgst")
        rev_x = sum_field("rev_cess")
        rec_i = sum_field("rec_igst")
        rec_c = sum_field("rec_cgst") + sum_field("rec_sgst")
        rec_x = sum_field("rec_cess")
        clo_i = (closing["clo_igst"] if closing else 0) or 0
        clo_c = ((closing["clo_cgst"] or 0) + (closing["clo_sgst"] or 0)
                 if closing else 0)
        clo_x = (closing["clo_cess"] if closing else 0) or 0

        # Health: total closing should equal opening + reversal - reclaimed
        # Simple status:
        #   Green if closing balance > 0 and reversal ≈ reclaimed
        #   Yellow if reversal > reclaimed (still parked)
        #   Red if closing < 0 (impossible — likely data issue)
        total_rev = rev_i + rev_c + rev_x
        total_rec = rec_i + rec_c + rec_x
        total_clo = clo_i + clo_c + clo_x

        if total_clo < 0:
            health = "⚠  Negative balance"
            health_fill = BAD_FILL
        elif total_rev == 0:
            health = "—  No activity"
            health_fill = META_FILL
        elif total_rec >= total_rev * 0.95:
            health = "✅  Mostly reclaimed"
            health_fill = GOOD_FILL
        elif total_rec >= total_rev * 0.5:
            health = "○  Partial reclaim"
            health_fill = WARN_FILL
        else:
            health = "⚠  Heavy parked balance"
            health_fill = WARN_FILL

        vals = [i, m["gstin"], m["state_name"], m["legal_name"], "",
                m["from_date"], m["to_date"], len(monthly),
                rev_i, rev_c, rev_x,
                rec_i, rec_c, rec_x,
                clo_i, clo_c, clo_x,
                health]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            if c == 1:
                cell.alignment = CENTER; cell.number_format = "0"
            elif c in (2, 3, 6, 7, 8):
                cell.alignment = CENTER
            elif c == 4:
                cell.alignment = LEFT
            elif c >= 9 and c <= 17 and isinstance(v, (int, float)):
                cell.alignment = RIGHT
                cell.number_format = NUM_FMT
                if v < 0: cell.font = NEG_FONT
            elif c == 18:
                cell.alignment = CENTER
                cell.fill = health_fill
                cell.font = BOLD
            else:
                cell.alignment = LEFT
        r += 1

    widths_sum = [6, 18, 16, 26, 10, 12, 12, 7,
                  14, 16, 11, 14, 16, 11, 14, 16, 11, 22]
    for i, w in enumerate(widths_sum, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    if r > 4:
        ws.auto_filter.ref = f"A3:R{r-1}"

    # ─── 4. Per-GSTIN detail sheets (portal layout) ───────────
    for data in all_data:
        m = data["meta"]
        # Sheet name: "<StateCode>_<last4 of GSTIN>"  (≤31 chars)
        sname = f"{m['state_code']}_{m['gstin'][-4:]}" if m['gstin'] else f"GSTIN_{len(wb.sheetnames)}"
        sname = sname[:31]
        # Handle duplicates
        base_sname = sname
        idx = 1
        while sname in wb.sheetnames:
            idx += 1
            sname = f"{base_sname[:28]}_{idx}"
        ws = wb.create_sheet(sname)

        # Title and meta
        ws.merge_cells("A1:U1")
        ws["A1"] = f"ECRRS — {m['gstin']}  ({m['state_name']})"
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 26

        ws.merge_cells("A2:U2")
        ws["A2"] = (f"{m['legal_name']}  |  Period: {m['from_date']} to {m['to_date']}")
        ws["A2"].font = SUBTITLE
        ws["A2"].alignment = CENTER
        ws.row_dimensions[2].height = 18

        # Group banner row 4
        gh = ["S.No.", "Date", "Reference No.", "Return Period", "Description"]
        for c, h in enumerate(gh, 1):
            ws.merge_cells(start_row=4, start_column=c, end_row=5, end_column=c)
            cell = ws.cell(row=4, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL

        # 4 grouped columns × 4 tax types
        group_specs = [
            (6, 9,  "ITC Claimed  [Table 4A(5)]",  GRP_FILL_CLAIM),
            (10, 13, "ITC Reversal  [Table 4B(2)]", GRP_FILL_REV),
            (14, 17, "ITC Reclaimed  [Table 4D(1)]", GRP_FILL_REC),
            (18, 21, "Closing Balance", GRP_FILL_BAL),
        ]
        for cs, ce, lab, fl in group_specs:
            ws.merge_cells(start_row=4, start_column=cs, end_row=4, end_column=ce)
            cell = ws.cell(row=4, column=cs, value=lab)
            cell.font = BOLD; cell.fill = fl
            cell.alignment = CENTER; cell.border = BORDER_ALL

        # Sub-headers row 5: IGST / CGST / SGST / Cess × 4
        tax_subs = ["IGST", "CGST", "SGST/UTGST", "Cess"]
        for grp_idx in range(4):
            for sub_idx, sub_label in enumerate(tax_subs):
                col = 6 + grp_idx * 4 + sub_idx
                cell = ws.cell(row=5, column=col, value=sub_label)
                cell.font = WHITE_B; cell.fill = SUB_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[4].height = 22
        ws.row_dimensions[5].height = 22

        # Data rows
        r = 6
        for row in data["rows"]:
            is_balance = row["row_type"] in ("Opening", "Closing")
            row_fill = (OPEN_FILL if row["row_type"] == "Opening"
                        else CLOSE_FILL if row["row_type"] == "Closing"
                        else None)
            vals = [
                row["sr_no"], row["date"], row["ref_no"],
                row["period"], row["desc"],
                row["claimed_igst"], row["claimed_cgst"],
                row["claimed_sgst"], row["claimed_cess"],
                row["rev_igst"], row["rev_cgst"],
                row["rev_sgst"], row["rev_cess"],
                row["rec_igst"], row["rec_cgst"],
                row["rec_sgst"], row["rec_cess"],
                row["clo_igst"], row["clo_cgst"],
                row["clo_sgst"], row["clo_cess"],
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER_ALL
                if row_fill: cell.fill = row_fill
                if c <= 5:
                    cell.alignment = CENTER if c != 5 else LEFT
                    cell.font = BOLD if is_balance else REG
                elif isinstance(v, (int, float)):
                    cell.alignment = RIGHT
                    cell.number_format = NUM_FMT
                    if v < 0:
                        cell.font = NEG_FONT
                    else:
                        cell.font = BOLD if is_balance else REG
                else:
                    cell.alignment = CENTER
                    cell.font = BOLD if is_balance else REG
            r += 1

        # Column widths
        det_widths = [6, 11, 18, 10, 26,
                      14, 14, 16, 11,
                      14, 14, 16, 11,
                      14, 14, 16, 11,
                      14, 14, 16, 11]
        for i, w in enumerate(det_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "F6"
        if r > 6:
            ws.auto_filter.ref = f"A5:U{r-1}"

    wb.save(out_path)
    return len(all_data), dedup_stats


# ════════════════════════════════════════════════════════════════
#  PDF MERGE / SPLIT ENGINES
# ════════════════════════════════════════════════════════════════
# Used by Tab 2 (PDF Tools) — merge multiple PDFs to one, split one
# into many with auto-fit-under-N-MB intelligence for GST portal
# (typical limit: 5 MB per file).


def get_pdf_info(filepath):
    """Return (page_count, size_bytes) for a PDF.
    page_count is 0 if PyMuPDF/pikepdf is unavailable."""
    try:
        size_bytes = os.path.getsize(filepath)
    except Exception:
        return (0, 0)
    pages = 0
    if HAS_PYMUPDF:
        try:
            import fitz
            with fitz.open(filepath) as doc:
                pages = doc.page_count
        except Exception:
            pages = 0
    elif HAS_PIKEPDF:
        try:
            import pikepdf
            with pikepdf.open(filepath) as pdf:
                pages = len(pdf.pages)
        except Exception:
            pages = 0
    return (pages, size_bytes)


def merge_pdfs(filepaths, output_path, progress_cb=None):
    """Merge multiple PDFs in given order. Returns output size in bytes.
    progress_cb(idx, total, name) called per file."""
    if not filepaths:
        raise ValueError("No PDFs to merge.")
    # Try pikepdf first (slightly better metadata handling); fall back to PyMuPDF
    if HAS_PIKEPDF:
        import pikepdf
        merged = pikepdf.Pdf.new()
        try:
            for i, fp in enumerate(filepaths, 1):
                if progress_cb:
                    progress_cb(i, len(filepaths), os.path.basename(fp))
                src = pikepdf.open(fp)
                try:
                    merged.pages.extend(src.pages)
                finally:
                    src.close()
            merged.save(output_path,
                        compress_streams=True,
                        object_stream_mode=pikepdf.ObjectStreamMode.generate)
        finally:
            merged.close()
    elif HAS_PYMUPDF:
        import fitz
        merged = fitz.open()
        try:
            for i, fp in enumerate(filepaths, 1):
                if progress_cb:
                    progress_cb(i, len(filepaths), os.path.basename(fp))
                src = fitz.open(fp)
                merged.insert_pdf(src)
                src.close()
            merged.save(output_path, garbage=4, deflate=True, clean=True)
        finally:
            merged.close()
    else:
        raise RuntimeError("Need PyMuPDF or pikepdf installed for merging.")
    return os.path.getsize(output_path)


def compute_merge_buckets(filepaths, target_mb, preserve_order=False):
    """Bin-pack PDFs into buckets such that each bucket's total size ≤ target_mb.

    Uses First-Fit-Decreasing (sort by size desc, then greedy fit) for best
    packing. With preserve_order=True, packs files in given order, starting a
    new bucket whenever the current one would overflow.

    Returns (buckets, oversize_files) where:
      buckets        — list of buckets, each a list of (filepath, size_bytes)
      oversize_files — list of (filepath, size_bytes) for files that alone
                       exceed target. These each become their own bucket
                       (and need pre-compression to actually fit target).
    """
    target_bytes = target_mb * 1024 * 1024
    items = [(fp, os.path.getsize(fp)) for fp in filepaths]
    # Separate oversize (each alone > target) from normal
    oversize = [(fp, sz) for fp, sz in items if sz > target_bytes]
    normal = [(fp, sz) for fp, sz in items if sz <= target_bytes]

    if not preserve_order:
        # First-Fit Decreasing: pack big first for best fill ratio
        normal.sort(key=lambda x: -x[1])

    buckets = []
    if preserve_order:
        # Sequential pack: add to current bucket if fits, else start new bucket
        current = []
        current_total = 0
        for fp, sz in normal:
            if current and current_total + sz > target_bytes:
                buckets.append(current)
                current = [(fp, sz)]
                current_total = sz
            else:
                current.append((fp, sz))
                current_total += sz
        if current:
            buckets.append(current)
    else:
        # First-Fit: try each existing bucket; create new if none fits
        for fp, sz in normal:
            placed = False
            for bucket in buckets:
                bucket_total = sum(b[1] for b in bucket)
                if bucket_total + sz <= target_bytes:
                    bucket.append((fp, sz))
                    placed = True
                    break
            if not placed:
                buckets.append([(fp, sz)])

    # Each oversize file gets its own bucket
    for fp, sz in oversize:
        buckets.append([(fp, sz)])
    return buckets, oversize


def split_pdf_by_chunk(input_path, output_dir, pages_per_chunk, prefix=None):
    """Split a PDF into chunks of N pages each. Returns list of output paths."""
    if prefix is None:
        prefix = os.path.splitext(os.path.basename(input_path))[0]
    os.makedirs(output_dir, exist_ok=True)
    outputs = []
    if HAS_PIKEPDF:
        import pikepdf
        with pikepdf.open(input_path) as src:
            total = len(src.pages)
            for chunk_idx, start in enumerate(range(0, total, pages_per_chunk), 1):
                end = min(start + pages_per_chunk, total)
                out_path = os.path.join(
                    output_dir,
                    f"{prefix}_part{chunk_idx}_p{start+1}-{end}.pdf")
                out_pdf = pikepdf.Pdf.new()
                try:
                    for i in range(start, end):
                        out_pdf.pages.append(src.pages[i])
                    out_pdf.save(out_path, compress_streams=True,
                                 object_stream_mode=pikepdf.ObjectStreamMode.generate)
                finally:
                    out_pdf.close()
                outputs.append(out_path)
    elif HAS_PYMUPDF:
        import fitz
        with fitz.open(input_path) as src:
            total = src.page_count
            for chunk_idx, start in enumerate(range(0, total, pages_per_chunk), 1):
                end = min(start + pages_per_chunk, total)
                out_path = os.path.join(
                    output_dir,
                    f"{prefix}_part{chunk_idx}_p{start+1}-{end}.pdf")
                out = fitz.open()
                try:
                    out.insert_pdf(src, from_page=start, to_page=end - 1)
                    out.save(out_path, garbage=4, deflate=True, clean=True)
                finally:
                    out.close()
                outputs.append(out_path)
    else:
        raise RuntimeError("Need PyMuPDF or pikepdf installed for splitting.")
    return outputs


def split_pdf_by_ranges(input_path, output_dir, ranges, prefix=None):
    """Split by explicit (start, end) 1-indexed inclusive page ranges.
    e.g. [(1, 5), (6, 10), (11, 20)]. Returns list of output paths."""
    if prefix is None:
        prefix = os.path.splitext(os.path.basename(input_path))[0]
    os.makedirs(output_dir, exist_ok=True)
    outputs = []
    if HAS_PIKEPDF:
        import pikepdf
        with pikepdf.open(input_path) as src:
            total = len(src.pages)
            for chunk_idx, (s1, e1) in enumerate(ranges, 1):
                if s1 < 1 or e1 > total or s1 > e1:
                    raise ValueError(
                        f"Range {s1}-{e1} invalid (PDF has {total} pages).")
                out_path = os.path.join(
                    output_dir, f"{prefix}_part{chunk_idx}_p{s1}-{e1}.pdf")
                out_pdf = pikepdf.Pdf.new()
                try:
                    for i in range(s1 - 1, e1):
                        out_pdf.pages.append(src.pages[i])
                    out_pdf.save(out_path, compress_streams=True,
                                 object_stream_mode=pikepdf.ObjectStreamMode.generate)
                finally:
                    out_pdf.close()
                outputs.append(out_path)
    elif HAS_PYMUPDF:
        import fitz
        with fitz.open(input_path) as src:
            total = src.page_count
            for chunk_idx, (s1, e1) in enumerate(ranges, 1):
                if s1 < 1 or e1 > total or s1 > e1:
                    raise ValueError(
                        f"Range {s1}-{e1} invalid (PDF has {total} pages).")
                out_path = os.path.join(
                    output_dir, f"{prefix}_part{chunk_idx}_p{s1}-{e1}.pdf")
                out = fitz.open()
                try:
                    out.insert_pdf(src, from_page=s1 - 1, to_page=e1 - 1)
                    out.save(out_path, garbage=4, deflate=True, clean=True)
                finally:
                    out.close()
                outputs.append(out_path)
    else:
        raise RuntimeError("Need PyMuPDF or pikepdf installed for splitting.")
    return outputs


def parse_page_ranges(ranges_str, total_pages):
    """Parse '1-5, 6-10, 11-20' or '1,3,5-7' → [(1,5), (6,10), (11,20)]."""
    ranges = []
    parts = [p.strip() for p in ranges_str.replace(';', ',').split(',') if p.strip()]
    for p in parts:
        if '-' in p:
            try:
                s, e = p.split('-', 1)
                s, e = int(s.strip()), int(e.strip())
            except Exception:
                raise ValueError(f"Bad range: {p!r}")
        else:
            try:
                s = e = int(p.strip())
            except Exception:
                raise ValueError(f"Bad page number: {p!r}")
        ranges.append((s, e))
    return ranges


def split_pdf_to_fit(input_path, output_dir, target_mb, prefix=None,
                     progress_cb=None):
    """Smart split — divide PDF into chunks so each output is ≤ target_mb.
    Returns list of output paths. Uses iterative refinement: estimate
    per-page size, chunk accordingly, then re-split any chunk that
    overshoots the target."""
    if prefix is None:
        prefix = os.path.splitext(os.path.basename(input_path))[0]
    os.makedirs(output_dir, exist_ok=True)
    pages, size_bytes = get_pdf_info(input_path)
    target_bytes = int(target_mb * 1024 * 1024)
    if pages <= 0:
        raise RuntimeError("Could not read page count from PDF.")
    # Already fits
    if size_bytes <= target_bytes:
        import shutil
        out_path = os.path.join(
            output_dir, f"{prefix}_fits_{size_bytes/1024/1024:.2f}MB.pdf")
        shutil.copy2(input_path, out_path)
        return [out_path]

    # Estimate per-page size + 15% safety buffer
    per_page = size_bytes / pages
    pages_per_chunk = max(1, int((target_bytes * 0.85) / per_page))
    if progress_cb:
        progress_cb(
            f"Estimate: per-page ≈ {per_page/1024:.1f} KB, "
            f"chunk = {pages_per_chunk} pages")

    initial = split_pdf_by_chunk(input_path, output_dir, pages_per_chunk, prefix)

    # Verify each chunk, sub-split if needed
    refined = []
    for out_path in initial:
        out_size = os.path.getsize(out_path)
        if out_size > target_bytes:
            # Re-split this chunk
            sub_prefix = os.path.splitext(os.path.basename(out_path))[0] + "_sub"
            sub_outputs = split_pdf_to_fit(out_path, output_dir, target_mb,
                                            prefix=sub_prefix,
                                            progress_cb=progress_cb)
            try: os.remove(out_path)
            except Exception: pass
            refined.extend(sub_outputs)
        else:
            refined.append(out_path)

    # Renumber outputs so they look clean (part1, part2, ...)
    final = []
    for i, fp in enumerate(refined, 1):
        # Get page range from original filename if present (best-effort rename)
        base = os.path.basename(fp)
        # Insert a clean ordinal prefix
        new_name = f"{prefix}_part{i:02d}_{base}"
        new_path = os.path.join(output_dir, new_name)
        try:
            os.rename(fp, new_path)
            final.append(new_path)
        except Exception:
            final.append(fp)
    return final


# ════════════════════════════════════════════════════════════════
#  ECL — Electronic Credit Ledger consolidator
# ════════════════════════════════════════════════════════════════
# Parses GSTN portal CSV export of "Electronic Credit Ledger" and
# consolidates multi-GSTIN/multi-period downloads into one workbook
# with Cover, Console (long-format), Summary, and per-GSTIN detail.


def _ledger_parse_num(s):
    """Convert '54,144,976' / '9,81,534.00' / '0' / '-' / '' → float | None."""
    if s is None: return None
    s = str(s).strip().replace(',', '').replace('₹', '').strip()
    if not s or s in ('-', '—', 'NA', 'N/A'): return None
    try: return float(s)
    except: return None


def parse_ecl_csv(filepath):
    """Parse one Electronic Credit Ledger CSV. Returns dict with meta + rows.

    Structure of the CSV (typical GSTN portal export):
       R0:  Title 'Electronic Credit Ledger'
       R2-5: meta (label col 3, value col 4) — GSTIN, Legal Name, From, To
       R6:  main header row (Sr.No, Date, Reference No., Tax Period, ...)
       R7:  sub-header row — Integrated/Central/State/CESS/Total × 2 groups
       R8+: data rows  (Opening Balance, txns, optional Closing Balance)

    Data row columns:
       0: Sr.No
       1: Date
       2: Reference No.
       3: Tax Period
       4: Description
       5: Transaction Type ('Credit'/'Debit'/'-')
       6-10:  Credit/Debit amount  (IGST / CGST / SGST / CESS / Total)
       11-15: Balance after txn   (IGST / CGST / SGST / CESS / Total)
    """
    import csv
    raw_rows = []
    with open(filepath, encoding='utf-8-sig', errors='replace') as f:
        for row in csv.reader(f):
            raw_rows.append(row)

    meta = {
        "source_file": os.path.basename(filepath),
        "gstin": "", "legal_name": "",
        "from_date": "", "to_date": "",
        "state_code": "", "state_name": "",
    }
    # Meta rows: try label cols (3 then 6) and value in next non-empty cell
    for r in raw_rows[:10]:
        for li in (3, 6):
            if li >= len(r): continue
            label = (r[li] or "").strip()
            if not label: continue
            # value in li+1 (or li+2 if li+1 blank)
            value = ""
            for vi in (li + 1, li + 2):
                if vi < len(r) and r[vi] and r[vi].strip():
                    value = r[vi].strip(); break
            if not value: continue
            ll = label.lower()
            if label == "GSTIN" and not meta["gstin"]:
                meta["gstin"] = value
            elif ("legal" in ll or "name" in ll) and not meta["legal_name"]:
                meta["legal_name"] = value
            elif label == "From" and not meta["from_date"]:
                meta["from_date"] = value
            elif label == "To" and not meta["to_date"]:
                meta["to_date"] = value
    if meta["gstin"]:
        meta["state_code"] = meta["gstin"][:2]
        meta["state_name"] = _STATE_CODES.get(meta["state_code"], "Unknown")

    # Find Sr.No header row to locate data start
    hdr_idx = None
    for i, r in enumerate(raw_rows):
        if r and r[0] and str(r[0]).strip().startswith("Sr.No"):
            hdr_idx = i; break
    if hdr_idx is None:
        return {"meta": meta, "rows": []}
    data_start = hdr_idx + 2   # skip sub-header row

    rows = []
    for r in raw_rows[data_start:]:
        if not r or not r[0] or not str(r[0]).strip():
            continue

        def col(i):
            return r[i].strip() if i < len(r) and r[i] else ""

        sr_no = col(0)
        date = col(1)
        ref_no = col(2)
        tax_period = col(3)
        desc = col(4)
        ttype = col(5)
        amt_igst = _ledger_parse_num(col(6))
        amt_cgst = _ledger_parse_num(col(7))
        amt_sgst = _ledger_parse_num(col(8))
        amt_cess = _ledger_parse_num(col(9))
        amt_total = _ledger_parse_num(col(10))
        bal_igst = _ledger_parse_num(col(11))
        bal_cgst = _ledger_parse_num(col(12))
        bal_sgst = _ledger_parse_num(col(13))
        bal_cess = _ledger_parse_num(col(14))
        bal_total = _ledger_parse_num(col(15))

        # Row classification
        dlow = desc.lower()
        if "opening balance" in dlow: row_type = "Opening"
        elif "closing balance" in dlow: row_type = "Closing"
        elif ttype.lower() == "credit": row_type = "Credit"
        elif ttype.lower() == "debit":  row_type = "Debit"
        else: row_type = "Other"

        rows.append({
            "sr_no": sr_no, "date": date, "ref_no": ref_no,
            "tax_period": tax_period, "desc": desc, "ttype": ttype,
            "row_type": row_type,
            "amt_igst": amt_igst, "amt_cgst": amt_cgst,
            "amt_sgst": amt_sgst, "amt_cess": amt_cess,
            "amt_total": amt_total,
            "bal_igst": bal_igst, "bal_cgst": bal_cgst,
            "bal_sgst": bal_sgst, "bal_cess": bal_cess,
            "bal_total": bal_total,
        })

    return {"meta": meta, "rows": rows}


def _ledger_parse_date(s):
    """Parse DD/MM/YYYY → date object, or None on failure."""
    from datetime import datetime
    if not s: return None
    s = str(s).strip()
    if not s or s == '-': return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    return None


def _ledger_txn_category(row):
    """Classify a ledger transaction. Returns 'Regular' or 'Extra'.

    Regular: standard GSTR-3B activity — ITC accrual, output discharge, RCM.
    Extra:   non-3B events — voluntary payments (DRC-03), refund claims,
             demand orders (DRC-07/01), recovery, anything not part of the
             regular monthly filing cycle.

    Used to keep Monthly Balances clean (only regular 3B impact per period)
    while surfacing extras in a dedicated 'Extra Payments' sheet.
    """
    desc = (row.get('desc') or '').lower().strip()
    if not desc or desc == '-':
        return 'Regular'
    extra_patterns = [
        'voluntary payment',
        'refund claimed',
        'refund from',
        'drc-03', 'drc 03', 'drc03',
        'drc-07', 'drc 07', 'drc07',
        'drc-01', 'drc 01', 'drc01',
        'on account of order',
        'demand', 'recovery',
    ]
    for pattern in extra_patterns:
        if pattern in desc:
            return 'Extra'
    return 'Regular'


def _ledger_txn_subcategory(row):
    """For 'Extra' category, identify subtype for Extra Payments sheet."""
    desc = (row.get('desc') or '').lower().strip()
    if 'voluntary' in desc: return 'Voluntary Payment'
    if 'refund' in desc:    return 'Refund Claim'
    if 'drc' in desc:       return 'DRC Payment'
    if 'demand' in desc:    return 'Demand Payment'
    if 'recovery' in desc:  return 'Recovery'
    if 'order' in desc:     return 'Order-based Payment'
    return 'Other Extra'


def dedup_ledger_files(all_data):
    """Merge multiple CSVs of the same GSTIN (overlapping date ranges) into
    one logical entry. Dedup transactions by Reference No.

    Use-case: user downloads Credit Ledger as "Apr-Feb" + "Feb-Mar" (because
    GSTN portal date range limit). Feb transactions appear in both files.
    Without dedup, they get double-counted.

    Returns (deduped_list, stats_list).
    stats: per-GSTIN dict with files_merged, duplicates_skipped, unique_txns.
    """
    by_gstin = {}
    for d in all_data:
        gstin = d['meta'].get('gstin', '')
        if not gstin:
            # Files without GSTIN — keep as standalone, can't dedup
            by_gstin.setdefault(f"_NO_GSTIN_{len(by_gstin)}", []).append(d)
        else:
            by_gstin.setdefault(gstin, []).append(d)

    deduped = []
    stats = []

    for gstin, files in by_gstin.items():
        if len(files) == 1:
            deduped.append(files[0])
            continue

        # Multiple files for same GSTIN — merge them
        def _sort_by_from(f):
            d = _ledger_parse_date(f['meta'].get('from_date', ''))
            return d if d else _ledger_parse_date('01/01/2000')
        files_sorted = sorted(files, key=_sort_by_from)

        merged_meta = dict(files_sorted[0]['meta'])
        # Widen date range to cover all files
        from_dates = [_ledger_parse_date(f['meta'].get('from_date', ''))
                      for f in files_sorted]
        to_dates = [_ledger_parse_date(f['meta'].get('to_date', ''))
                    for f in files_sorted]
        from_dates = [d for d in from_dates if d]
        to_dates = [d for d in to_dates if d]
        if from_dates:
            merged_meta['from_date'] = min(from_dates).strftime('%d/%m/%Y')
        if to_dates:
            merged_meta['to_date'] = max(to_dates).strftime('%d/%m/%Y')
        merged_meta['source_file'] = ' + '.join(
            f['meta'].get('source_file', '?') for f in files_sorted)

        # Dedup rows by Reference No within Credit/Debit txns.
        # Opening Balance: keep from EARLIEST file (true period start).
        # Closing Balance: keep from LATEST file (true period end).
        opening_row = None
        latest_closing = None
        seen_refs = set()
        merged_txns = []
        dup_count = 0

        for f in files_sorted:
            for row in f['rows']:
                if row['row_type'] == 'Opening':
                    if opening_row is None:
                        opening_row = row
                elif row['row_type'] == 'Closing':
                    latest_closing = row
                else:
                    ref = (row.get('ref_no') or '').strip()
                    if ref and ref != '-':
                        if ref in seen_refs:
                            dup_count += 1
                            continue
                        seen_refs.add(ref)
                    merged_txns.append(row)

        merged_rows = []
        if opening_row: merged_rows.append(opening_row)
        merged_rows.extend(merged_txns)
        if latest_closing: merged_rows.append(latest_closing)

        deduped.append({'meta': merged_meta, 'rows': merged_rows})
        stats.append({
            'gstin': gstin,
            'state_name': merged_meta.get('state_name', '?'),
            'files_merged': len(files_sorted),
            'duplicates_skipped': dup_count,
            'unique_txns': len(merged_txns),
        })

    return deduped, stats


def _tp_sort_key(tp):
    """Sort 'Apr-25' chronologically. Returns (year, month) tuple.
    Unknown/blank values sort last."""
    if not tp: return (9999, 99)
    s = str(tp).strip()
    if not s or s in ('-', '—', '(no tax period)'): return (9999, 99)
    months_map = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                  "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
    parts = s.replace('/', '-').split('-')
    if len(parts) != 2: return (9999, 99)
    m = parts[0].strip()[:3].capitalize()
    yy = parts[1].strip()
    if m not in months_map: return (9999, 99)
    try:
        full_year = 2000 + int(yy) if len(yy) <= 2 else int(yy)
    except ValueError:
        return (9999, 99)
    return (full_year, months_map[m])


def compute_tax_period_balances(data, ledger_type='ecl', regular_only=True):
    """For each Tax Period, compute opening + closing balance.

    regular_only=True (DEFAULT): excludes Voluntary/Refund/DRC/etc Extra
                                 payments — gives clean monthly view of pure
                                 GSTR-3B activity.
    regular_only=False: includes all txns (legacy behaviour).

    Opening of Tax Period X = balance just BEFORE the FIRST regular txn
                              tagged with Tax Period = X (chronological).
    Closing of Tax Period X = balance AFTER the LAST regular such txn.
    """
    from datetime import datetime
    rows = data["rows"]
    meta = data["meta"]

    def _extract_bal(row):
        if not row:
            return {"igst": 0, "cgst": 0, "sgst": 0, "cess": 0, "total": 0}
        if ledger_type == 'ecl':
            return {
                "igst":  row["bal_igst"]  or 0,
                "cgst":  row["bal_cgst"]  or 0,
                "sgst":  row["bal_sgst"]  or 0,
                "cess":  row["bal_cess"]  or 0,
                "total": row["bal_total"] or 0,
            }
        else:  # cash
            bal = {
                "igst": (row["bal_igst"]["total"] or 0)
                        if isinstance(row["bal_igst"], dict) else 0,
                "cgst": (row["bal_cgst"]["total"] or 0)
                        if isinstance(row["bal_cgst"], dict) else 0,
                "sgst": (row["bal_sgst"]["total"] or 0)
                        if isinstance(row["bal_sgst"], dict) else 0,
                "cess": (row["bal_cess"]["total"] or 0)
                        if isinstance(row["bal_cess"], dict) else 0,
            }
            bal["total"] = bal["igst"] + bal["cgst"] + bal["sgst"] + bal["cess"]
            return bal

    opening_row = next((r for r in rows if r["row_type"] == "Opening"), None)
    initial_bal = _extract_bal(opening_row)

    # Sort ALL non-meta txns chronologically (need full sequence for
    # "balance just before" calculations even when filtering to Regular)
    full_seq = []
    for i, r in enumerate(rows):
        if r["row_type"] not in ("Credit", "Debit"): continue
        d = _ledger_parse_date(r["date"])
        sort_d = d if d else _ledger_parse_date(meta.get("from_date"))
        if not sort_d:
            sort_d = datetime.min.date()
        full_seq.append((sort_d, i, r))
    full_seq.sort(key=lambda x: (x[0], x[1]))

    # bal_after_full_seq[k] = balance after the k-th chronological txn
    bal_after_full = [_extract_bal(r) for (_, _, r) in full_seq]

    # Filter to category if needed; track original full-sequence index
    filtered_indices = []
    for k, (_, _, r) in enumerate(full_seq):
        if regular_only and _ledger_txn_category(r) != 'Regular':
            continue
        filtered_indices.append(k)

    # Group filtered indices by Tax Period
    by_tp = {}
    for k in filtered_indices:
        row = full_seq[k][2]
        tp = (row.get("tax_period", "") or "").strip()
        if not tp or tp == '-':
            tp = "(no tax period)"
        by_tp.setdefault(tp, []).append(k)

    results = []
    for tp, k_list in by_tp.items():
        first_k = min(k_list)
        last_k = max(k_list)
        # Opening = balance just BEFORE first_k. That's bal_after of (first_k-1)
        # in the FULL chronological sequence (even if it was an Extra event).
        # This matches the actual ledger state at that moment.
        opening = (initial_bal if first_k == 0 else bal_after_full[first_k - 1])
        closing = bal_after_full[last_k]
        results.append({
            "tax_period": tp,
            "sort_key": _tp_sort_key(tp),
            "opening": dict(opening),
            "closing": dict(closing),
            "txn_count": len(k_list),
        })
    results.sort(key=lambda x: x["sort_key"])
    return results


def compute_extra_payments(data, ledger_type='ecl'):
    """Return list of extra-payment rows (voluntary, refund, DRC, etc.)
    with full context for the 'Extra Payments' summary sheet.

    Each extra row also includes 'effect_tax_period': the Tax Period of the
    NEXT regular 3B transaction (chronologically) after this extra event.
    This is the period where a gap will appear when comparing the user's
    internal monthly 3B working against the actual ECL — because the next
    regular 3B filing's opening balance will reflect this extra payment.
    """
    from datetime import datetime
    rows = data["rows"]

    # Build chronological index of ALL real transactions (Regular + Extra)
    # so we can look forward to find the next Regular after each Extra.
    indexed = []
    for i, r in enumerate(rows):
        if r["row_type"] not in ("Credit", "Debit"): continue
        d = _ledger_parse_date(r.get("date", ""))
        sort_d = d if d else datetime.min.date()
        indexed.append((sort_d, i, r))
    indexed.sort(key=lambda x: (x[0], x[1]))

    extras = []
    for k, (_, _, r) in enumerate(indexed):
        if _ledger_txn_category(r) != 'Extra': continue

        # Find next REGULAR txn AFTER this extra event → its tax period is
        # the one where the gap will manifest in user-vs-ECL comparison.
        effect_tp = ""
        for j in range(k + 1, len(indexed)):
            next_r = indexed[j][2]
            if _ledger_txn_category(next_r) == 'Regular':
                effect_tp = (next_r.get("tax_period", "") or "").strip()
                break
        # If no future regular txn exists (extra is the last activity), mark
        # as 'next period' — affects the period after the latest one filed
        if not effect_tp:
            effect_tp = "(no future 3B filed)"

        if ledger_type == 'ecl':
            amt_breakdown = {
                "igst": r["amt_igst"] or 0,
                "cgst": r["amt_cgst"] or 0,
                "sgst": r["amt_sgst"] or 0,
                "cess": r["amt_cess"] or 0,
                "total": r["amt_total"] or 0,
            }
            bal_total = r["bal_total"] or 0
        else:
            amt_breakdown = {
                "igst": (r["amt_igst"]["total"] or 0)
                        if isinstance(r["amt_igst"], dict) else 0,
                "cgst": (r["amt_cgst"]["total"] or 0)
                        if isinstance(r["amt_cgst"], dict) else 0,
                "sgst": (r["amt_sgst"]["total"] or 0)
                        if isinstance(r["amt_sgst"], dict) else 0,
                "cess": (r["amt_cess"]["total"] or 0)
                        if isinstance(r["amt_cess"], dict) else 0,
            }
            amt_breakdown["total"] = (amt_breakdown["igst"] + amt_breakdown["cgst"]
                                       + amt_breakdown["sgst"] + amt_breakdown["cess"])
            bal_total = ((r["bal_igst"]["total"] or 0) if isinstance(r["bal_igst"], dict) else 0) \
                        + ((r["bal_cgst"]["total"] or 0) if isinstance(r["bal_cgst"], dict) else 0) \
                        + ((r["bal_sgst"]["total"] or 0) if isinstance(r["bal_sgst"], dict) else 0) \
                        + ((r["bal_cess"]["total"] or 0) if isinstance(r["bal_cess"], dict) else 0)
        extras.append({
            "date": r.get("date", ""),
            "txn_month": _ledger_txn_month_label(r.get("date", "")),
            "ref_no": r.get("ref_no", ""),
            "tax_period": r.get("tax_period", ""),
            "effect_tax_period": effect_tp,
            "desc": r.get("desc", ""),
            "ttype": r.get("ttype", ""),
            "subcategory": _ledger_txn_subcategory(r),
            "amt": amt_breakdown,
            "bal_total_after": bal_total,
        })
    return extras


def _ledger_txn_month_label(date_str):
    """Convert a DD/MM/YYYY date string into 'Apr-25' style month label.
    Used to add a 'Txn Month' column in Extra Payments sheet, mirroring the
    Tax Period format but derived from the actual transaction Date column."""
    d = _ledger_parse_date(date_str)
    if not d: return ""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{months[d.month - 1]}-{d.year % 100:02d}"


def _ledger_month_label(d):
    """[deprecated for monthly balances — kept for back-compat] Format 'Apr-25'."""
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    return f"{months[d.month - 1]}-{d.year % 100:02d}"


def _ledger_add_months(d, n=1):
    """[deprecated for monthly balances — kept for back-compat]."""
    from datetime import date
    total = d.year * 12 + (d.month - 1) + n
    new_year, new_month = divmod(total, 12)
    return date(new_year, new_month + 1, 1)


def compute_monthly_balances(data, ledger_type='ecl'):
    """[Back-compat wrapper] Now delegates to compute_tax_period_balances
    with regular_only=True (clean monthly view, excludes voluntary/refund)."""
    results = compute_tax_period_balances(data, ledger_type, regular_only=True)
    return [{"month": r["tax_period"],
             "year": r["sort_key"][0],
             "month_num": r["sort_key"][1],
             "sort_key": r["sort_key"][0] * 12 + r["sort_key"][1],
             "opening": r["opening"],
             "closing": r["closing"],
             "txn_count": r["txn_count"]}
            for r in results]


def write_consolidated_ecl(all_data, out_path):
    """Build consolidated workbook for Electronic Credit Ledger data.
    Returns count of GSTINs."""
    import openpyxl
    from openpyxl.styles import (Font as XF, PatternFill, Alignment,
                                  Border, Side)
    from openpyxl.utils import get_column_letter

    HDR_FILL = PatternFill("solid", start_color="1F4E78")
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    META_FILL = PatternFill("solid", start_color="F2F2F2")
    AMT_FILL = PatternFill("solid", start_color="E2EFDA")
    BAL_FILL = PatternFill("solid", start_color="FFF2CC")
    OPEN_FILL = PatternFill("solid", start_color="E7E6E6")
    CLOSE_FILL = PatternFill("solid", start_color="FFE699")
    TOT_FILL = PatternFill("solid", start_color="D9E1F2")

    WHITE_B = XF(name="Calibri", bold=True, color="FFFFFF", size=10)
    BOLD = XF(name="Calibri", bold=True, size=10)
    REG = XF(name="Calibri", size=10)
    NEG_FONT = XF(name="Calibri", size=10, color="C00000")
    TITLE = XF(name="Calibri", bold=True, color="FFFFFF", size=14)
    SUBTITLE = XF(name="Calibri", italic=True, size=10, color="404040")

    thin = Side(border_style="thin", color="B4B4B4")
    BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT = Alignment(horizontal="right", vertical="center")
    NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ─── Cover ────────────────────────────────────────────
    ws = wb.create_sheet("Cover")
    ws.merge_cells("A1:H1")
    ws["A1"] = "Electronic Credit Ledger — Multi-GSTIN Consolidated"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:H2")
    ws["A2"] = ("ITC Credit / Debit / Closing Balance per GSTIN  ·  "
                "Console + Summary + per-GSTIN detail")
    ws["A2"].font = SUBTITLE; ws["A2"].alignment = CENTER

    hdrs = ["S.No.", "GSTIN", "State Code", "State Name", "Legal Name",
            "From", "To", "Source File"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[4].height = 22

    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        for c, v in enumerate([i, m["gstin"], m["state_code"], m["state_name"],
                                m["legal_name"], m["from_date"], m["to_date"],
                                m["source_file"]], 1):
            cell = ws.cell(row=4 + i, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            cell.alignment = LEFT if c >= 4 else CENTER
            if c == 1: cell.number_format = "0"

    for col, w in zip("ABCDEFGH", [6, 18, 7, 18, 28, 12, 12, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ─── Monthly Balances — per GSTIN × per Tax Period ────
    # Walks every GSTIN's transactions, groups by Tax Period column (NOT
    # calendar month — important for GST workflow). Opening of Tax Period X
    # = balance just before the first transaction tagged with Tax Period X
    # in chronological (Date) order. Example: Apr-25 opening = balance just
    # before the first Apr-25 entry (typically dated 20-May when GSTR-3B for
    # Apr-25 was filed) — i.e. balance after any preceding May entries.
    ws = wb.create_sheet("Monthly Balances")
    cols = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name",
            "Tax Period", "FY Year", "Txn Count",
            "Opening IGST", "Opening CGST", "Opening SGST/UTGST",
            "Opening CESS", "Opening Total",
            "Closing IGST", "Closing CGST", "Closing SGST/UTGST",
            "Closing CESS", "Closing Total",
            "Net Change"]
    ncols_m = len(cols)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols_m)
    ws["A1"] = ("Monthly Balances — Opening + Closing per GSTIN, per Tax Period   "
                "(opening = balance just before first txn of that tax period)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    def banner(c1, c2, label, fill):
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        cell = ws.cell(row=2, column=c1, value=label)
        cell.font = BOLD; cell.fill = fill
        cell.alignment = CENTER; cell.border = BORDER_ALL
    banner(1, 8, "Metadata", META_FILL)
    banner(9, 13, "Opening Balance (before first txn of tax period)", AMT_FILL)
    banner(14, 18, "Closing Balance (after last txn of tax period)", BAL_FILL)
    banner(19, 19, "Δ", META_FILL)
    ws.row_dimensions[2].height = 22

    for c, h in enumerate(cols, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    GAIN_FILL = PatternFill("solid", start_color="C6EFCE")  # green
    LOSS_FILL = PatternFill("solid", start_color="FFC7CE")  # red

    r = 4; sn = 0
    for data in all_data:
        m = data["meta"]
        monthly_list = compute_monthly_balances(data, ledger_type='ecl')
        for mb in monthly_list:
            sn += 1
            opg, clo = mb["opening"], mb["closing"]
            net_change = (clo["total"] or 0) - (opg["total"] or 0)
            vals = [
                sn, m["state_code"], m["state_name"], m["gstin"], m["legal_name"],
                mb["month"], mb["year"], mb["txn_count"],
                opg["igst"], opg["cgst"], opg["sgst"], opg["cess"], opg["total"],
                clo["igst"], clo["cgst"], clo["sgst"], clo["cess"], clo["total"],
                net_change,
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = REG; cell.border = BORDER_ALL
                if c == 1:
                    cell.alignment = CENTER; cell.number_format = "0"
                elif c in (2, 4, 6, 7, 8):
                    cell.alignment = CENTER
                elif c in (3, 5):
                    cell.alignment = LEFT
                elif isinstance(v, (int, float)) and c >= 9:
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                # Colour the Net Change cell green/red
                if c == 19 and isinstance(v, (int, float)) and v != 0:
                    cell.fill = GAIN_FILL if v > 0 else LOSS_FILL
                    cell.font = BOLD if v > 0 else NEG_FONT
            r += 1

    widths_m = [6, 7, 16, 18, 22,
                10, 7, 8,
                13, 13, 15, 11, 13,
                13, 13, 15, 11, 13,
                14]
    for i, w in enumerate(widths_m, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "I4"
    if r > 4:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols_m)}{r-1}"

    # ─── Extra Payments sheet (Voluntary / Refund / DRC / etc.) ──
    # Tracks non-regular-3B events that affected the credit ledger.
    # These are EXCLUDED from Monthly Balances (which shows only regular
    # GSTR-3B activity per tax period). Listing them here gives full audit
    # trail of voluntary payments, refund claims, DRC payments, etc.
    #
    # Three "period" dimensions per extra payment:
    #   • Txn Month  = calendar month of the payment (Date column)
    #   • Tax Period = period the payment is FOR (relates to)
    #   • Effect TP  = period whose 3B comparison will show a gap (the next
    #                  regular 3B filing's tax period — useful when comparing
    #                  user's internal 3B working vs the actual ECL balance)
    ws = wb.create_sheet("Extra Payments")
    cols_e = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name",
              "Date", "Txn Month", "Reference No.", "Tax Period", "Effect TP",
              "Subcategory", "Description", "Txn Type",
              "Amt IGST", "Amt CGST", "Amt SGST/UTGST", "Amt CESS", "Amt Total",
              "Bal Total After"]
    ncols_e = len(cols_e)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols_e)
    ws["A1"] = ("Extra Payments — Voluntary / Refund / DRC / Demand / Recovery   "
                "(Effect TP = the tax period whose 3B comparison will show a "
                "gap from this extra payment)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    SUBCAT_FILL = PatternFill("solid", start_color="FFE6CC")     # orange
    TXN_MONTH_FILL = PatternFill("solid", start_color="D9E1F2")  # blue
    EFFECT_TP_FILL = PatternFill("solid", start_color="C6EFCE")  # green
    for c, h in enumerate(cols_e, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 30

    r = 4; sn = 0
    subcategory_totals = {}
    txn_month_totals = {}
    effect_tp_totals = {}
    for data in all_data:
        m = data["meta"]
        extras = compute_extra_payments(data, ledger_type='ecl')
        for ex in extras:
            sn += 1
            sub = ex["subcategory"]
            txm = ex["txn_month"] or "(unknown)"
            etp = ex["effect_tax_period"] or "(unknown)"
            subcategory_totals.setdefault(sub, {"count": 0, "total": 0})
            subcategory_totals[sub]["count"] += 1
            subcategory_totals[sub]["total"] += ex["amt"]["total"] or 0
            txn_month_totals.setdefault(txm, {"count": 0, "total": 0})
            txn_month_totals[txm]["count"] += 1
            txn_month_totals[txm]["total"] += ex["amt"]["total"] or 0
            effect_tp_totals.setdefault(etp, {"count": 0, "total": 0})
            effect_tp_totals[etp]["count"] += 1
            effect_tp_totals[etp]["total"] += ex["amt"]["total"] or 0
            vals = [
                sn, m["state_code"], m["state_name"], m["gstin"], m["legal_name"],
                ex["date"], ex["txn_month"], ex["ref_no"], ex["tax_period"],
                ex["effect_tax_period"], sub,
                ex["desc"], ex["ttype"],
                ex["amt"]["igst"], ex["amt"]["cgst"],
                ex["amt"]["sgst"], ex["amt"]["cess"], ex["amt"]["total"],
                ex["bal_total_after"],
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = REG; cell.border = BORDER_ALL
                if c == 1:
                    cell.alignment = CENTER; cell.number_format = "0"
                elif c in (2, 4, 6, 8, 9, 13):
                    cell.alignment = CENTER
                elif c == 7:   # Txn Month
                    cell.alignment = CENTER
                    cell.fill = TXN_MONTH_FILL
                    cell.font = BOLD
                elif c == 10:  # Effect TP — gap indicator
                    cell.alignment = CENTER
                    cell.fill = EFFECT_TP_FILL
                    cell.font = BOLD
                elif c == 11:  # Subcategory
                    cell.alignment = CENTER
                    cell.fill = SUBCAT_FILL
                    cell.font = BOLD
                elif c in (3, 5, 12):
                    cell.alignment = LEFT
                elif isinstance(v, (int, float)) and c >= 14:
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                else:
                    cell.alignment = LEFT
            r += 1

    # Footer — THREE summary tables side by side
    if subcategory_totals or txn_month_totals or effect_tp_totals:
        r += 1
        # Row of section titles
        ws.cell(row=r, column=1, value="SUMMARY BY SUBCATEGORY").font = BOLD
        ws.cell(row=r, column=1).alignment = LEFT
        ws.cell(row=r, column=6, value="SUMMARY BY EFFECT TP (3B gap)").font = BOLD
        ws.cell(row=r, column=6).alignment = LEFT
        ws.cell(row=r, column=11, value="SUMMARY BY TXN MONTH").font = BOLD
        ws.cell(row=r, column=11).alignment = LEFT
        r += 1
        # Column header row
        for c, h in enumerate(["Subcategory", "# Events", "Total Amount"], 1):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        for c, h in enumerate(["Effect TP", "# Events", "Total Amount"], 6):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        for c, h in enumerate(["Txn Month", "# Events", "Total Amount"], 11):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        r += 1
        subcat_sorted = sorted(subcategory_totals.items())
        etp_sorted = sorted(effect_tp_totals.items(),
                            key=lambda x: _tp_sort_key(x[0]))
        txm_sorted = sorted(txn_month_totals.items(),
                            key=lambda x: _tp_sort_key(x[0]))
        max_rows = max(len(subcat_sorted), len(etp_sorted), len(txm_sorted))
        for i in range(max_rows):
            if i < len(subcat_sorted):
                k, vs = subcat_sorted[i]
                ws.cell(row=r, column=1, value=k).font = BOLD
                ws.cell(row=r, column=1).fill = SUBCAT_FILL
                ws.cell(row=r, column=1).border = BORDER_ALL
                ws.cell(row=r, column=1).alignment = CENTER
                ws.cell(row=r, column=2, value=vs["count"]).alignment = CENTER
                ws.cell(row=r, column=2).border = BORDER_ALL
                ws.cell(row=r, column=3, value=vs["total"]).number_format = NUM_FMT
                ws.cell(row=r, column=3).alignment = RIGHT
                ws.cell(row=r, column=3).border = BORDER_ALL
            if i < len(etp_sorted):
                k, vs = etp_sorted[i]
                ws.cell(row=r, column=6, value=k).font = BOLD
                ws.cell(row=r, column=6).fill = EFFECT_TP_FILL
                ws.cell(row=r, column=6).border = BORDER_ALL
                ws.cell(row=r, column=6).alignment = CENTER
                ws.cell(row=r, column=7, value=vs["count"]).alignment = CENTER
                ws.cell(row=r, column=7).border = BORDER_ALL
                ws.cell(row=r, column=8, value=vs["total"]).number_format = NUM_FMT
                ws.cell(row=r, column=8).alignment = RIGHT
                ws.cell(row=r, column=8).border = BORDER_ALL
            if i < len(txm_sorted):
                k, vs = txm_sorted[i]
                ws.cell(row=r, column=11, value=k).font = BOLD
                ws.cell(row=r, column=11).fill = TXN_MONTH_FILL
                ws.cell(row=r, column=11).border = BORDER_ALL
                ws.cell(row=r, column=11).alignment = CENTER
                ws.cell(row=r, column=12, value=vs["count"]).alignment = CENTER
                ws.cell(row=r, column=12).border = BORDER_ALL
                ws.cell(row=r, column=13, value=vs["total"]).number_format = NUM_FMT
                ws.cell(row=r, column=13).alignment = RIGHT
                ws.cell(row=r, column=13).border = BORDER_ALL
            r += 1

    widths_e = [6, 7, 16, 18, 22,
                11, 10, 18, 10, 10, 18, 30, 8,
                13, 13, 15, 11, 13, 16]
    for i, w in enumerate(widths_e, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "I4"
    if sn > 0:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols_e)}{3 + sn}"

    # ─── Console (long format, every row from every GSTIN) ─
    ws = wb.create_sheet("Console")
    cols = ["S.No.", "GSTIN", "State Code", "State Name", "Legal Name",
            "Row Type", "Sr.No.", "Date", "Reference No.", "Tax Period",
            "Description", "Transaction Type",
            "Amt IGST", "Amt CGST", "Amt SGST/UTGST", "Amt CESS", "Amt Total",
            "Bal IGST", "Bal CGST", "Bal SGST/UTGST", "Bal CESS", "Bal Total"]
    ncols = len(cols)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws["A1"] = ("Console — every row, every GSTIN  "
                "(filter by State / GSTIN / Row Type / Tax Period)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    def banner(c1, c2, label, fill):
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        cell = ws.cell(row=2, column=c1, value=label)
        cell.font = BOLD; cell.fill = fill
        cell.alignment = CENTER; cell.border = BORDER_ALL

    banner(1, 12, "Metadata", META_FILL)
    banner(13, 17, "Amount (Credit/Debit)", AMT_FILL)
    banner(18, 22, "Balance after transaction", BAL_FILL)
    ws.row_dimensions[2].height = 22

    for c, h in enumerate(cols, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4; sn = 0
    for data in all_data:
        m = data["meta"]
        for row in data["rows"]:
            sn += 1
            vals = [
                sn, m["gstin"], m["state_code"], m["state_name"],
                m["legal_name"],
                row["row_type"], row["sr_no"], row["date"], row["ref_no"],
                row["tax_period"], row["desc"], row["ttype"],
                row["amt_igst"], row["amt_cgst"], row["amt_sgst"],
                row["amt_cess"], row["amt_total"],
                row["bal_igst"], row["bal_cgst"], row["bal_sgst"],
                row["bal_cess"], row["bal_total"],
            ]
            row_bold = row["row_type"] in ("Opening", "Closing")
            row_fill = (OPEN_FILL if row["row_type"] == "Opening"
                        else CLOSE_FILL if row["row_type"] == "Closing"
                        else None)
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER_ALL
                if row_fill: cell.fill = row_fill
                if c == 1:
                    cell.alignment = CENTER; cell.number_format = "0"
                    cell.font = BOLD if row_bold else REG
                elif c in (2, 3, 6, 7, 8, 10, 12):
                    cell.alignment = CENTER
                    cell.font = BOLD if row_bold else REG
                elif c in (4, 5, 9, 11):
                    cell.alignment = LEFT
                    cell.font = BOLD if row_bold else REG
                elif isinstance(v, (int, float)) and c >= 13:
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                    else: cell.font = BOLD if row_bold else REG
                else:
                    cell.alignment = LEFT
                    cell.font = BOLD if row_bold else REG
            r += 1

    widths = [6, 18, 6, 16, 22,
              9, 6, 11, 18, 10, 28, 11,
              13, 13, 15, 11, 13,
              13, 13, 15, 11, 13]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "M4"
    if r > 4:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{r-1}"

    # ─── Summary — per-GSTIN totals ───────────────────────
    ws = wb.create_sheet("Summary")
    ws.merge_cells("A1:N1")
    ws["A1"] = "Summary — Total Credit / Total Debit / Closing Balance per GSTIN"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    sum_hdrs = ["S.No.", "GSTIN", "State", "Legal Name", "From", "To",
                "Txn Count",
                "Total Credit IGST", "Total Credit CGST+SGST", "Total Credit CESS",
                "Total Debit IGST", "Total Debit CGST+SGST", "Total Debit CESS",
                "Closing Total"]
    for c, h in enumerate(sum_hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4
    grand_totals = [0] * 7  # cr_i, cr_c, cr_x, dr_i, dr_c, dr_x, closing
    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        txns = [row for row in data["rows"]
                if row["row_type"] in ("Credit", "Debit")]
        closing = next((row for row in data["rows"]
                        if row["row_type"] == "Closing"), None)

        def sum_field(rows, key):
            return sum((row[key] or 0) for row in rows)

        credits = [r2 for r2 in txns if r2["row_type"] == "Credit"]
        debits  = [r2 for r2 in txns if r2["row_type"] == "Debit"]
        cr_i = sum_field(credits, "amt_igst")
        cr_c = sum_field(credits, "amt_cgst") + sum_field(credits, "amt_sgst")
        cr_x = sum_field(credits, "amt_cess")
        dr_i = sum_field(debits, "amt_igst")
        dr_c = sum_field(debits, "amt_cgst") + sum_field(debits, "amt_sgst")
        dr_x = sum_field(debits, "amt_cess")
        clo_total = (closing["bal_total"] if closing else 0) or 0

        for k, val in enumerate([cr_i, cr_c, cr_x, dr_i, dr_c, dr_x, clo_total]):
            grand_totals[k] += val or 0

        vals = [i, m["gstin"], m["state_name"], m["legal_name"],
                m["from_date"], m["to_date"], len(txns),
                cr_i, cr_c, cr_x, dr_i, dr_c, dr_x, clo_total]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            if c == 1:
                cell.alignment = CENTER; cell.number_format = "0"
            elif c in (2, 3, 5, 6, 7):
                cell.alignment = CENTER
            elif c == 4:
                cell.alignment = LEFT
            elif c >= 8 and isinstance(v, (int, float)):
                cell.alignment = RIGHT; cell.number_format = NUM_FMT
                if v < 0: cell.font = NEG_FONT
            else:
                cell.alignment = LEFT
        r += 1

    # Grand total row
    if all_data:
        cell = ws.cell(row=r, column=1, value="TOTAL")
        cell.font = BOLD; cell.fill = TOT_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
        for c in (2, 3, 4, 5, 6, 7):
            ws.cell(row=r, column=c).fill = TOT_FILL
            ws.cell(row=r, column=c).border = BORDER_ALL
        for k, val in enumerate(grand_totals):
            cell = ws.cell(row=r, column=8 + k, value=val)
            cell.font = BOLD; cell.fill = TOT_FILL
            cell.alignment = RIGHT
            cell.border = BORDER_ALL; cell.number_format = NUM_FMT
            if val < 0: cell.font = XF(name="Calibri", bold=True, color="C00000", size=10)
        r += 1

    widths_sum = [6, 18, 16, 26, 12, 12, 8, 14, 16, 11, 14, 16, 11, 14]
    for i, w in enumerate(widths_sum, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    # ─── Per-GSTIN detail sheets ──────────────────────────
    for data in all_data:
        m = data["meta"]
        sname = (f"{m['state_code']}_{m['gstin'][-4:]}"
                 if m['gstin'] else f"GSTIN_{len(wb.sheetnames)}")
        sname = sname[:31]
        base = sname; idx = 1
        while sname in wb.sheetnames:
            idx += 1; sname = f"{base[:28]}_{idx}"
        ws = wb.create_sheet(sname)

        ws.merge_cells("A1:P1")
        ws["A1"] = f"Credit Ledger — {m['gstin']}  ({m['state_name']})"
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 26
        ws.merge_cells("A2:P2")
        ws["A2"] = (f"{m['legal_name']}  |  Period: {m['from_date']} to "
                    f"{m['to_date']}")
        ws["A2"].font = SUBTITLE; ws["A2"].alignment = CENTER
        ws.row_dimensions[2].height = 18

        gh = ["S.No.", "Date", "Reference No.", "Tax Period", "Description",
              "Transaction Type"]
        for c, h in enumerate(gh, 1):
            ws.merge_cells(start_row=4, start_column=c, end_row=5, end_column=c)
            cell = ws.cell(row=4, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL

        ws.merge_cells(start_row=4, start_column=7, end_row=4, end_column=11)
        cell = ws.cell(row=4, column=7, value="Amount (Credit/Debit)")
        cell.font = BOLD; cell.fill = AMT_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.merge_cells(start_row=4, start_column=12, end_row=4, end_column=16)
        cell = ws.cell(row=4, column=12, value="Balance after transaction")
        cell.font = BOLD; cell.fill = BAL_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL

        subs = ["IGST", "CGST", "SGST/UTGST", "CESS", "Total"]
        for gi in range(2):
            for j, sub in enumerate(subs):
                col = 7 + gi * 5 + j
                cell = ws.cell(row=5, column=col, value=sub)
                cell.font = WHITE_B; cell.fill = SUB_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[4].height = 22
        ws.row_dimensions[5].height = 22

        r = 6
        for row in data["rows"]:
            is_balance = row["row_type"] in ("Opening", "Closing")
            row_fill = (OPEN_FILL if row["row_type"] == "Opening"
                        else CLOSE_FILL if row["row_type"] == "Closing"
                        else None)
            vals = [
                row["sr_no"], row["date"], row["ref_no"], row["tax_period"],
                row["desc"], row["ttype"],
                row["amt_igst"], row["amt_cgst"], row["amt_sgst"],
                row["amt_cess"], row["amt_total"],
                row["bal_igst"], row["bal_cgst"], row["bal_sgst"],
                row["bal_cess"], row["bal_total"],
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER_ALL
                if row_fill: cell.fill = row_fill
                if c <= 6:
                    cell.alignment = CENTER if c != 5 else LEFT
                    cell.font = BOLD if is_balance else REG
                elif isinstance(v, (int, float)):
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                    else: cell.font = BOLD if is_balance else REG
                else:
                    cell.alignment = CENTER
                    cell.font = BOLD if is_balance else REG
            r += 1

        det_widths = [6, 11, 18, 10, 30, 10,
                      13, 13, 15, 11, 13,
                      13, 13, 15, 11, 13]
        for i, w in enumerate(det_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "G6"
        if r > 6:
            ws.auto_filter.ref = f"A5:P{r-1}"

    wb.save(out_path)
    return len(all_data)


# ════════════════════════════════════════════════════════════════
#  ECashL — Electronic Cash Ledger consolidator
# ════════════════════════════════════════════════════════════════
# Parses GSTN portal CSV export of "Electronic Cash Ledger". Cash Ledger
# has FINER granularity than Credit — every tax type (IGST / CGST / SGST /
# CESS) has 6 sub-cols (Tax / Interest / Penalty / Fee / Others / Total),
# giving 4×6 = 24 amount cols + 4×6 = 24 balance cols (48 numeric cols).
# Console shows Total per tax type for compact view; per-GSTIN sheet has
# the full breakdown.


def parse_ecashl_csv(filepath):
    """Parse one Electronic Cash Ledger CSV. Returns dict with meta + rows.

    Structure:
       R0:  Title 'Electronic Cash Ledger'
       R2-5: meta (label col 6, value col 7) — GSTIN, Name(Legal), From, To
       R6:  main header row
       R7:  sub-header row — Tax/Interest/Penalty/Fee/Others/Total × 8 groups
       R8+: data rows

    Data row columns:
       0: Sr.No
       1: Date of deposit/Debit
       2: Time of deposit
       3: Reporting date (by bank)
       4: Reference No.
       5: Tax Period
       6: Description
       7: Transaction Type ('Debit'/'Credit')
       8-13:  IGST amount  (Tax/Interest/Penalty/Fee/Others/Total)
       14-19: CGST amount
       20-25: SGST amount
       26-31: CESS amount
       32-37: IGST balance
       38-43: CGST balance
       44-49: SGST balance
       50-55: CESS balance
    """
    import csv
    raw_rows = []
    with open(filepath, encoding='utf-8-sig', errors='replace') as f:
        for row in csv.reader(f):
            raw_rows.append(row)

    meta = {
        "source_file": os.path.basename(filepath),
        "gstin": "", "legal_name": "",
        "from_date": "", "to_date": "",
        "state_code": "", "state_name": "",
    }
    # Cash Ledger meta is in column 6 (label) + col 7 (value); fall back to 3
    for r in raw_rows[:10]:
        for li in (6, 3):
            if li >= len(r): continue
            label = (r[li] or "").strip()
            if not label: continue
            value = ""
            for vi in (li + 1, li + 2):
                if vi < len(r) and r[vi] and r[vi].strip():
                    value = r[vi].strip(); break
            if not value: continue
            ll = label.lower()
            if label == "GSTIN" and not meta["gstin"]:
                meta["gstin"] = value
            elif ("legal" in ll or "name" in ll) and not meta["legal_name"]:
                meta["legal_name"] = value
            elif label == "From" and not meta["from_date"]:
                meta["from_date"] = value
            elif label == "To" and not meta["to_date"]:
                meta["to_date"] = value
    if meta["gstin"]:
        meta["state_code"] = meta["gstin"][:2]
        meta["state_name"] = _STATE_CODES.get(meta["state_code"], "Unknown")

    # Locate Sr.No header row
    hdr_idx = None
    for i, r in enumerate(raw_rows):
        if r and r[0] and str(r[0]).strip().startswith("Sr.No"):
            hdr_idx = i; break
    if hdr_idx is None:
        return {"meta": meta, "rows": []}
    data_start = hdr_idx + 2

    rows = []
    for r in raw_rows[data_start:]:
        if not r or not r[0] or not str(r[0]).strip():
            continue

        def col(i):
            return r[i].strip() if i < len(r) and r[i] else ""

        def block(start):
            return {
                "tax":      _ledger_parse_num(col(start + 0)),
                "interest": _ledger_parse_num(col(start + 1)),
                "penalty":  _ledger_parse_num(col(start + 2)),
                "fee":      _ledger_parse_num(col(start + 3)),
                "others":   _ledger_parse_num(col(start + 4)),
                "total":    _ledger_parse_num(col(start + 5)),
            }

        sr_no = col(0); date = col(1); time = col(2)
        rep_date = col(3); ref_no = col(4); tax_period = col(5)
        desc = col(6); ttype = col(7)

        amt_igst = block(8);  amt_cgst = block(14)
        amt_sgst = block(20); amt_cess = block(26)
        bal_igst = block(32); bal_cgst = block(38)
        bal_sgst = block(44); bal_cess = block(50)

        dlow = desc.lower()
        if "opening balance" in dlow: row_type = "Opening"
        elif "closing balance" in dlow: row_type = "Closing"
        elif ttype.lower() == "credit": row_type = "Credit"
        elif ttype.lower() == "debit":  row_type = "Debit"
        else: row_type = "Other"

        rows.append({
            "sr_no": sr_no, "date": date, "time": time,
            "rep_date": rep_date, "ref_no": ref_no,
            "tax_period": tax_period, "desc": desc, "ttype": ttype,
            "row_type": row_type,
            "amt_igst": amt_igst, "amt_cgst": amt_cgst,
            "amt_sgst": amt_sgst, "amt_cess": amt_cess,
            "bal_igst": bal_igst, "bal_cgst": bal_cgst,
            "bal_sgst": bal_sgst, "bal_cess": bal_cess,
        })

    return {"meta": meta, "rows": rows}


def write_consolidated_ecashl(all_data, out_path):
    """Build consolidated workbook for Electronic Cash Ledger.
    Returns count of GSTINs."""
    import openpyxl
    from openpyxl.styles import (Font as XF, PatternFill, Alignment,
                                  Border, Side)
    from openpyxl.utils import get_column_letter

    HDR_FILL = PatternFill("solid", start_color="1F4E78")
    SUB_FILL = PatternFill("solid", start_color="2E75B6")
    META_FILL = PatternFill("solid", start_color="F2F2F2")
    IGST_FILL = PatternFill("solid", start_color="DDEBF7")
    CGST_FILL = PatternFill("solid", start_color="E2EFDA")
    SGST_FILL = PatternFill("solid", start_color="FCE4D6")
    CESS_FILL = PatternFill("solid", start_color="FFF2CC")
    BAL_FILL = PatternFill("solid", start_color="F4B084")
    OPEN_FILL = PatternFill("solid", start_color="E7E6E6")
    CLOSE_FILL = PatternFill("solid", start_color="FFE699")
    TOT_FILL = PatternFill("solid", start_color="D9E1F2")

    WHITE_B = XF(name="Calibri", bold=True, color="FFFFFF", size=10)
    BOLD = XF(name="Calibri", bold=True, size=10)
    REG = XF(name="Calibri", size=10)
    NEG_FONT = XF(name="Calibri", size=10, color="C00000")
    TITLE = XF(name="Calibri", bold=True, color="FFFFFF", size=14)
    SUBTITLE = XF(name="Calibri", italic=True, size=10, color="404040")

    thin = Side(border_style="thin", color="B4B4B4")
    BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    RIGHT = Alignment(horizontal="right", vertical="center")
    NUM_FMT = '#,##0.00;[Red](#,##0.00);"-"'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ─── Cover ────────────────────────────────────────────
    ws = wb.create_sheet("Cover")
    ws.merge_cells("A1:H1")
    ws["A1"] = "Electronic Cash Ledger — Multi-GSTIN Consolidated"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:H2")
    ws["A2"] = ("Cash deposit / debit / closing balance per GSTIN  ·  "
                "Tax / Interest / Penalty / Fee / Others breakdown")
    ws["A2"].font = SUBTITLE; ws["A2"].alignment = CENTER

    hdrs = ["S.No.", "GSTIN", "State Code", "State Name", "Legal Name",
            "From", "To", "Source File"]
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[4].height = 22

    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        for c, v in enumerate([i, m["gstin"], m["state_code"], m["state_name"],
                                m["legal_name"], m["from_date"], m["to_date"],
                                m["source_file"]], 1):
            cell = ws.cell(row=4 + i, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            cell.alignment = LEFT if c >= 4 else CENTER
            if c == 1: cell.number_format = "0"
    for col, w in zip("ABCDEFGH", [6, 18, 7, 18, 28, 12, 12, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    # ─── Monthly Balances — per GSTIN × per Tax Period (Cash) ──
    # Cash Ledger: shows Totals per tax type. Full Tax/Interest/Penalty/Fee/
    # Others breakdown remains in per-GSTIN detail sheets. Same tax-period
    # logic as Credit Ledger — opening = balance just before first txn of
    # that tax period.
    ws = wb.create_sheet("Monthly Balances")
    cols = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name",
            "Tax Period", "FY Year", "Txn Count",
            "Opening IGST", "Opening CGST", "Opening SGST/UTGST",
            "Opening CESS", "Opening Total",
            "Closing IGST", "Closing CGST", "Closing SGST/UTGST",
            "Closing CESS", "Closing Total",
            "Net Change"]
    ncols_m = len(cols)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols_m)
    ws["A1"] = ("Monthly Balances — Cash Ledger Opening + Closing per GSTIN, "
                "per Tax Period   (opening = balance just before first txn of "
                "that tax period)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    def banner(c1, c2, label, fill):
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        cell = ws.cell(row=2, column=c1, value=label)
        cell.font = BOLD; cell.fill = fill
        cell.alignment = CENTER; cell.border = BORDER_ALL
    banner(1, 8, "Metadata", META_FILL)
    banner(9, 13, "Opening Balance (before first txn of tax period)", CGST_FILL)
    banner(14, 18, "Closing Balance (after last txn of tax period)", BAL_FILL)
    banner(19, 19, "Δ", META_FILL)
    ws.row_dimensions[2].height = 22

    for c, h in enumerate(cols, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    GAIN_FILL = PatternFill("solid", start_color="C6EFCE")
    LOSS_FILL = PatternFill("solid", start_color="FFC7CE")

    r = 4; sn = 0
    for data in all_data:
        m = data["meta"]
        monthly_list = compute_monthly_balances(data, ledger_type='cash')
        for mb in monthly_list:
            sn += 1
            opg, clo = mb["opening"], mb["closing"]
            net_change = (clo["total"] or 0) - (opg["total"] or 0)
            vals = [
                sn, m["state_code"], m["state_name"], m["gstin"], m["legal_name"],
                mb["month"], mb["year"], mb["txn_count"],
                opg["igst"], opg["cgst"], opg["sgst"], opg["cess"], opg["total"],
                clo["igst"], clo["cgst"], clo["sgst"], clo["cess"], clo["total"],
                net_change,
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = REG; cell.border = BORDER_ALL
                if c == 1:
                    cell.alignment = CENTER; cell.number_format = "0"
                elif c in (2, 4, 6, 7, 8):
                    cell.alignment = CENTER
                elif c in (3, 5):
                    cell.alignment = LEFT
                elif isinstance(v, (int, float)) and c >= 9:
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                if c == 19 and isinstance(v, (int, float)) and v != 0:
                    cell.fill = GAIN_FILL if v > 0 else LOSS_FILL
                    cell.font = BOLD if v > 0 else NEG_FONT
            r += 1

    widths_m = [6, 7, 16, 18, 22,
                10, 7, 8,
                13, 13, 15, 11, 13,
                13, 13, 15, 11, 13,
                14]
    for i, w in enumerate(widths_m, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "I4"
    if r > 4:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols_m)}{r-1}"

    # ─── Extra Payments sheet (Cash — Voluntary / DRC / Refund / etc.) ──
    # Same logic as Credit Ledger Extra Payments — Txn Month is calendar
    # month of payment, Tax Period is period it relates to, Effect TP is
    # the tax period whose 3B comparison will show a gap.
    ws = wb.create_sheet("Extra Payments")
    cols_e = ["S.No.", "State Code", "State Name", "GSTIN", "Legal Name",
              "Date", "Txn Month", "Reference No.", "Tax Period", "Effect TP",
              "Subcategory", "Description", "Txn Type",
              "Amt IGST", "Amt CGST", "Amt SGST/UTGST", "Amt CESS", "Amt Total",
              "Bal Total After"]
    ncols_e = len(cols_e)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols_e)
    ws["A1"] = ("Extra Payments — Voluntary / Refund / DRC / Demand / Recovery   "
                "(Cash Ledger; Effect TP = tax period whose 3B comparison "
                "will show a gap from this payment)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    SUBCAT_FILL = PatternFill("solid", start_color="FFE6CC")
    TXN_MONTH_FILL = PatternFill("solid", start_color="D9E1F2")
    EFFECT_TP_FILL = PatternFill("solid", start_color="C6EFCE")
    for c, h in enumerate(cols_e, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 30

    r = 4; sn = 0
    subcategory_totals = {}
    txn_month_totals = {}
    effect_tp_totals = {}
    for data in all_data:
        m = data["meta"]
        extras = compute_extra_payments(data, ledger_type='cash')
        for ex in extras:
            sn += 1
            sub = ex["subcategory"]
            txm = ex["txn_month"] or "(unknown)"
            etp = ex["effect_tax_period"] or "(unknown)"
            subcategory_totals.setdefault(sub, {"count": 0, "total": 0})
            subcategory_totals[sub]["count"] += 1
            subcategory_totals[sub]["total"] += ex["amt"]["total"] or 0
            txn_month_totals.setdefault(txm, {"count": 0, "total": 0})
            txn_month_totals[txm]["count"] += 1
            txn_month_totals[txm]["total"] += ex["amt"]["total"] or 0
            effect_tp_totals.setdefault(etp, {"count": 0, "total": 0})
            effect_tp_totals[etp]["count"] += 1
            effect_tp_totals[etp]["total"] += ex["amt"]["total"] or 0
            vals = [
                sn, m["state_code"], m["state_name"], m["gstin"], m["legal_name"],
                ex["date"], ex["txn_month"], ex["ref_no"], ex["tax_period"],
                ex["effect_tax_period"], sub,
                ex["desc"], ex["ttype"],
                ex["amt"]["igst"], ex["amt"]["cgst"],
                ex["amt"]["sgst"], ex["amt"]["cess"], ex["amt"]["total"],
                ex["bal_total_after"],
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font = REG; cell.border = BORDER_ALL
                if c == 1:
                    cell.alignment = CENTER; cell.number_format = "0"
                elif c in (2, 4, 6, 8, 9, 13):
                    cell.alignment = CENTER
                elif c == 7:
                    cell.alignment = CENTER
                    cell.fill = TXN_MONTH_FILL
                    cell.font = BOLD
                elif c == 10:
                    cell.alignment = CENTER
                    cell.fill = EFFECT_TP_FILL
                    cell.font = BOLD
                elif c == 11:
                    cell.alignment = CENTER
                    cell.fill = SUBCAT_FILL
                    cell.font = BOLD
                elif c in (3, 5, 12):
                    cell.alignment = LEFT
                elif isinstance(v, (int, float)) and c >= 14:
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                else:
                    cell.alignment = LEFT
            r += 1

    if subcategory_totals or txn_month_totals or effect_tp_totals:
        r += 1
        ws.cell(row=r, column=1, value="SUMMARY BY SUBCATEGORY").font = BOLD
        ws.cell(row=r, column=1).alignment = LEFT
        ws.cell(row=r, column=6, value="SUMMARY BY EFFECT TP (3B gap)").font = BOLD
        ws.cell(row=r, column=6).alignment = LEFT
        ws.cell(row=r, column=11, value="SUMMARY BY TXN MONTH").font = BOLD
        ws.cell(row=r, column=11).alignment = LEFT
        r += 1
        for c, h in enumerate(["Subcategory", "# Events", "Total Amount"], 1):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        for c, h in enumerate(["Effect TP", "# Events", "Total Amount"], 6):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        for c, h in enumerate(["Txn Month", "# Events", "Total Amount"], 11):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL
        r += 1
        subcat_sorted = sorted(subcategory_totals.items())
        etp_sorted = sorted(effect_tp_totals.items(),
                            key=lambda x: _tp_sort_key(x[0]))
        txm_sorted = sorted(txn_month_totals.items(),
                            key=lambda x: _tp_sort_key(x[0]))
        max_rows = max(len(subcat_sorted), len(etp_sorted), len(txm_sorted))
        for i in range(max_rows):
            if i < len(subcat_sorted):
                k, vs = subcat_sorted[i]
                ws.cell(row=r, column=1, value=k).font = BOLD
                ws.cell(row=r, column=1).fill = SUBCAT_FILL
                ws.cell(row=r, column=1).border = BORDER_ALL
                ws.cell(row=r, column=1).alignment = CENTER
                ws.cell(row=r, column=2, value=vs["count"]).alignment = CENTER
                ws.cell(row=r, column=2).border = BORDER_ALL
                ws.cell(row=r, column=3, value=vs["total"]).number_format = NUM_FMT
                ws.cell(row=r, column=3).alignment = RIGHT
                ws.cell(row=r, column=3).border = BORDER_ALL
            if i < len(etp_sorted):
                k, vs = etp_sorted[i]
                ws.cell(row=r, column=6, value=k).font = BOLD
                ws.cell(row=r, column=6).fill = EFFECT_TP_FILL
                ws.cell(row=r, column=6).border = BORDER_ALL
                ws.cell(row=r, column=6).alignment = CENTER
                ws.cell(row=r, column=7, value=vs["count"]).alignment = CENTER
                ws.cell(row=r, column=7).border = BORDER_ALL
                ws.cell(row=r, column=8, value=vs["total"]).number_format = NUM_FMT
                ws.cell(row=r, column=8).alignment = RIGHT
                ws.cell(row=r, column=8).border = BORDER_ALL
            if i < len(txm_sorted):
                k, vs = txm_sorted[i]
                ws.cell(row=r, column=11, value=k).font = BOLD
                ws.cell(row=r, column=11).fill = TXN_MONTH_FILL
                ws.cell(row=r, column=11).border = BORDER_ALL
                ws.cell(row=r, column=11).alignment = CENTER
                ws.cell(row=r, column=12, value=vs["count"]).alignment = CENTER
                ws.cell(row=r, column=12).border = BORDER_ALL
                ws.cell(row=r, column=13, value=vs["total"]).number_format = NUM_FMT
                ws.cell(row=r, column=13).alignment = RIGHT
                ws.cell(row=r, column=13).border = BORDER_ALL
            r += 1

    widths_e = [6, 7, 16, 18, 22,
                11, 10, 18, 10, 10, 18, 30, 8,
                13, 13, 15, 11, 13, 16]
    for i, w in enumerate(widths_e, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "I4"
    if sn > 0:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols_e)}{3 + sn}"

    # ─── Console (Totals view) ────────────────────────────
    ws = wb.create_sheet("Console")
    cols = ["S.No.", "GSTIN", "State Code", "State Name", "Legal Name",
            "Row Type", "Sr.No.", "Date", "Reference No.", "Tax Period",
            "Description", "Transaction Type",
            "IGST Amt Total", "CGST Amt Total",
            "SGST Amt Total", "CESS Amt Total",
            "IGST Bal Total", "CGST Bal Total",
            "SGST Bal Total", "CESS Bal Total"]
    ncols = len(cols)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws["A1"] = ("Console — every row, every GSTIN  "
                "(Totals shown; full Tax/Interest/Penalty/Fee/Others "
                "breakdown in per-GSTIN detail sheets)")
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    def banner(c1, c2, label, fill):
        ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
        cell = ws.cell(row=2, column=c1, value=label)
        cell.font = BOLD; cell.fill = fill
        cell.alignment = CENTER; cell.border = BORDER_ALL
    banner(1, 12, "Metadata", META_FILL)
    banner(13, 16, "Cash Amount (Credit/Debit) — Totals", CGST_FILL)
    banner(17, 20, "Cash Balance after transaction — Totals", BAL_FILL)
    ws.row_dimensions[2].height = 22

    for c, h in enumerate(cols, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4; sn = 0
    for data in all_data:
        m = data["meta"]
        for row in data["rows"]:
            sn += 1
            vals = [
                sn, m["gstin"], m["state_code"], m["state_name"],
                m["legal_name"],
                row["row_type"], row["sr_no"], row["date"], row["ref_no"],
                row["tax_period"], row["desc"], row["ttype"],
                row["amt_igst"]["total"], row["amt_cgst"]["total"],
                row["amt_sgst"]["total"], row["amt_cess"]["total"],
                row["bal_igst"]["total"], row["bal_cgst"]["total"],
                row["bal_sgst"]["total"], row["bal_cess"]["total"],
            ]
            row_bold = row["row_type"] in ("Opening", "Closing")
            row_fill = (OPEN_FILL if row["row_type"] == "Opening"
                        else CLOSE_FILL if row["row_type"] == "Closing"
                        else None)
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER_ALL
                if row_fill: cell.fill = row_fill
                if c == 1:
                    cell.alignment = CENTER; cell.number_format = "0"
                    cell.font = BOLD if row_bold else REG
                elif c in (2, 3, 6, 7, 8, 10, 12):
                    cell.alignment = CENTER
                    cell.font = BOLD if row_bold else REG
                elif c in (4, 5, 9, 11):
                    cell.alignment = LEFT
                    cell.font = BOLD if row_bold else REG
                elif isinstance(v, (int, float)) and c >= 13:
                    cell.alignment = RIGHT; cell.number_format = NUM_FMT
                    if v < 0: cell.font = NEG_FONT
                    else: cell.font = BOLD if row_bold else REG
                else:
                    cell.alignment = LEFT
                    cell.font = BOLD if row_bold else REG
            r += 1

    widths = [6, 18, 6, 16, 22,
              9, 6, 11, 18, 10, 28, 11,
              14, 14, 14, 13,
              14, 14, 14, 13]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "M4"
    if r > 4:
        ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{r-1}"

    # ─── Summary (per-GSTIN totals) ───────────────────────
    ws = wb.create_sheet("Summary")
    ws.merge_cells("A1:N1")
    ws["A1"] = "Summary — Cash Credit / Debit / Closing Balance per GSTIN"
    ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 28

    sum_hdrs = ["S.No.", "GSTIN", "State", "Legal Name", "From", "To",
                "Txn Count",
                "Cr IGST", "Cr CGST+SGST", "Cr CESS",
                "Dr IGST", "Dr CGST+SGST", "Dr CESS",
                "Closing (4-head total)"]
    for c, h in enumerate(sum_hdrs, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = WHITE_B; cell.fill = SUB_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
    ws.row_dimensions[3].height = 36

    r = 4
    grand_totals = [0] * 7
    for i, data in enumerate(all_data, 1):
        m = data["meta"]
        txns = [row for row in data["rows"]
                if row["row_type"] in ("Credit", "Debit")]
        closing = next((row for row in data["rows"]
                        if row["row_type"] == "Closing"), None)

        def sum_total(rows, key):
            return sum((row[key]["total"] or 0) for row in rows)

        credits = [r2 for r2 in txns if r2["row_type"] == "Credit"]
        debits  = [r2 for r2 in txns if r2["row_type"] == "Debit"]
        cr_i = sum_total(credits, "amt_igst")
        cr_c = sum_total(credits, "amt_cgst") + sum_total(credits, "amt_sgst")
        cr_x = sum_total(credits, "amt_cess")
        dr_i = sum_total(debits, "amt_igst")
        dr_c = sum_total(debits, "amt_cgst") + sum_total(debits, "amt_sgst")
        dr_x = sum_total(debits, "amt_cess")
        clo_total = 0
        if closing:
            clo_total = ((closing["bal_igst"]["total"] or 0)
                         + (closing["bal_cgst"]["total"] or 0)
                         + (closing["bal_sgst"]["total"] or 0)
                         + (closing["bal_cess"]["total"] or 0))

        for k, val in enumerate([cr_i, cr_c, cr_x, dr_i, dr_c, dr_x, clo_total]):
            grand_totals[k] += val or 0

        vals = [i, m["gstin"], m["state_name"], m["legal_name"],
                m["from_date"], m["to_date"], len(txns),
                cr_i, cr_c, cr_x, dr_i, dr_c, dr_x, clo_total]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = REG; cell.border = BORDER_ALL
            if c == 1:
                cell.alignment = CENTER; cell.number_format = "0"
            elif c in (2, 3, 5, 6, 7):
                cell.alignment = CENTER
            elif c == 4:
                cell.alignment = LEFT
            elif c >= 8 and isinstance(v, (int, float)):
                cell.alignment = RIGHT; cell.number_format = NUM_FMT
                if v < 0: cell.font = NEG_FONT
            else:
                cell.alignment = LEFT
        r += 1

    if all_data:
        cell = ws.cell(row=r, column=1, value="TOTAL")
        cell.font = BOLD; cell.fill = TOT_FILL
        cell.alignment = CENTER; cell.border = BORDER_ALL
        for c in (2, 3, 4, 5, 6, 7):
            ws.cell(row=r, column=c).fill = TOT_FILL
            ws.cell(row=r, column=c).border = BORDER_ALL
        for k, val in enumerate(grand_totals):
            cell = ws.cell(row=r, column=8 + k, value=val)
            cell.font = BOLD; cell.fill = TOT_FILL
            cell.alignment = RIGHT
            cell.border = BORDER_ALL; cell.number_format = NUM_FMT
            if val < 0:
                cell.font = XF(name="Calibri", bold=True, color="C00000", size=10)
        r += 1

    widths_sum = [6, 18, 16, 26, 12, 12, 8,
                  14, 16, 11, 14, 16, 11, 18]
    for i, w in enumerate(widths_sum, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    # ─── Per-GSTIN detail sheets (full 48-col layout) ─────
    SUBS = ["Tax", "Interest", "Penalty", "Fee", "Others", "Total"]
    GROUPS = [
        ("IGST",     "amt_igst", IGST_FILL),
        ("CGST",     "amt_cgst", CGST_FILL),
        ("SGST",     "amt_sgst", SGST_FILL),
        ("CESS",     "amt_cess", CESS_FILL),
        ("IGST Bal", "bal_igst", BAL_FILL),
        ("CGST Bal", "bal_cgst", BAL_FILL),
        ("SGST Bal", "bal_sgst", BAL_FILL),
        ("CESS Bal", "bal_cess", BAL_FILL),
    ]
    for data in all_data:
        m = data["meta"]
        sname = (f"{m['state_code']}_{m['gstin'][-4:]}"
                 if m['gstin'] else f"G_{len(wb.sheetnames)}")
        sname = sname[:31]
        base = sname; idx = 1
        while sname in wb.sheetnames:
            idx += 1; sname = f"{base[:28]}_{idx}"
        ws = wb.create_sheet(sname)

        total_cols = 8 + 8 * 6  # 8 meta + 48 numeric = 56
        last_col = get_column_letter(total_cols)
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=total_cols)
        ws["A1"] = f"Cash Ledger — {m['gstin']}  ({m['state_name']})"
        ws["A1"].font = TITLE; ws["A1"].fill = HDR_FILL
        ws["A1"].alignment = CENTER
        ws.row_dimensions[1].height = 26
        ws.merge_cells(start_row=2, start_column=1,
                       end_row=2, end_column=total_cols)
        ws["A2"] = (f"{m['legal_name']}  |  Period: {m['from_date']} to "
                    f"{m['to_date']}")
        ws["A2"].font = SUBTITLE; ws["A2"].alignment = CENTER
        ws.row_dimensions[2].height = 18

        meta_h = ["S.No.", "Date", "Time", "Ref. Date", "Reference No.",
                  "Tax Period", "Description", "Transaction Type"]
        for c, h in enumerate(meta_h, 1):
            ws.merge_cells(start_row=4, start_column=c, end_row=5, end_column=c)
            cell = ws.cell(row=4, column=c, value=h)
            cell.font = WHITE_B; cell.fill = SUB_FILL
            cell.alignment = CENTER; cell.border = BORDER_ALL

        for gi, (label, _, fill) in enumerate(GROUPS):
            cs = 9 + gi * 6
            ce = cs + 5
            ws.merge_cells(start_row=4, start_column=cs,
                           end_row=4, end_column=ce)
            cell = ws.cell(row=4, column=cs, value=label)
            cell.font = BOLD; cell.fill = fill
            cell.alignment = CENTER; cell.border = BORDER_ALL
            for si, sub in enumerate(SUBS):
                col = cs + si
                cell = ws.cell(row=5, column=col, value=sub)
                cell.font = WHITE_B; cell.fill = SUB_FILL
                cell.alignment = CENTER; cell.border = BORDER_ALL
        ws.row_dimensions[4].height = 22
        ws.row_dimensions[5].height = 22

        r = 6
        for row in data["rows"]:
            is_balance = row["row_type"] in ("Opening", "Closing")
            row_fill = (OPEN_FILL if row["row_type"] == "Opening"
                        else CLOSE_FILL if row["row_type"] == "Closing"
                        else None)
            base_vals = [row["sr_no"], row["date"], row["time"],
                         row["rep_date"], row["ref_no"], row["tax_period"],
                         row["desc"], row["ttype"]]
            for c, v in enumerate(base_vals, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = BORDER_ALL
                if row_fill: cell.fill = row_fill
                cell.alignment = CENTER if c != 7 else LEFT
                cell.font = BOLD if is_balance else REG

            for gi, (_, key, _) in enumerate(GROUPS):
                blk = row[key]
                for si, sub in enumerate(SUBS):
                    v = blk[sub.lower()]
                    col = 9 + gi * 6 + si
                    cell = ws.cell(row=r, column=col, value=v)
                    cell.border = BORDER_ALL
                    if row_fill: cell.fill = row_fill
                    if isinstance(v, (int, float)):
                        cell.alignment = RIGHT; cell.number_format = NUM_FMT
                        if v < 0: cell.font = NEG_FONT
                        else: cell.font = BOLD if is_balance else REG
                    else:
                        cell.alignment = CENTER
                        cell.font = BOLD if is_balance else REG
            r += 1

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 11
        ws.column_dimensions["C"].width = 9
        ws.column_dimensions["D"].width = 11
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 10
        ws.column_dimensions["G"].width = 26
        ws.column_dimensions["H"].width = 11
        for c in range(9, 9 + 8 * 6):
            ws.column_dimensions[get_column_letter(c)].width = 11
        ws.freeze_panes = "I6"
        if r > 6:
            ws.auto_filter.ref = f"A5:{last_col}{r-1}"

    wb.save(out_path)
    return len(all_data)


# ════════════════════════════════════════════════════════════════
#  APP  — entire GUI wrapped so any error pops a messagebox
# ════════════════════════════════════════════════════════════════
try:

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("GST Tools Suite")
            # Compact window so RUN button stays visible even with Windows 125% scaling
            self.geometry("780x640")
            self.minsize(720, 540)
            self.configure(bg=BG)
            self.resizable(True, True)
            self.rows = []
            self._build_ui()
            # Bring window to front on startup
            try:
                self.update_idletasks()
                self.attributes("-topmost", True)
                self.lift()
                self.focus_force()
                self.after(500, lambda: self.attributes("-topmost", False))
            except Exception:
                pass

        def _build_ui(self):
            tk.Frame(self, bg=ACCENT, height=4).pack(fill="x")
            tk.Label(self, text="🗂  GST Tools Suite",
                     font=("Segoe UI", 14, "bold"),
                     bg=BG, fg=TEXT).pack(pady=(6, 0))

            style = ttk.Style(self)
            try: style.theme_use("clam")
            except: pass
            style.configure("TNotebook",     background=BG, borderwidth=0)
            style.configure("TNotebook.Tab", background=CARD, foreground=SUBTEXT,
                            font=("Segoe UI", 9, "bold"), padding=[8, 5])
            style.map("TNotebook.Tab",
                      background=[("selected", ACCENT)],
                      foreground=[("selected", "white")])
            style.configure("TCombobox", fieldbackground=ENTRY_BG,
                            background=ENTRY_BG, foreground=TEXT)

            nb = ttk.Notebook(self)
            nb.pack(fill="both", expand=True, padx=14, pady=(8, 10))
            t1 = tk.Frame(nb, bg=CARD)
            t2 = tk.Frame(nb, bg=CARD)
            t3 = tk.Frame(nb, bg=CARD)
            t4 = tk.Frame(nb, bg=CARD)
            t5 = tk.Frame(nb, bg=CARD)
            t6 = tk.Frame(nb, bg=CARD)
            t8 = tk.Frame(nb, bg=CARD)
            t9 = tk.Frame(nb, bg=CARD)
            t10 = tk.Frame(nb, bg=CARD)     # Credit Ledger
            t11 = tk.Frame(nb, bg=CARD)     # Cash Ledger
            t7 = tk.Frame(nb, bg=CARD)
            nb.add(t1, text=" 📁 Coll ")
            nb.add(t2, text=" 📄 Comp ")
            nb.add(t3, text=" 📊 Tax ")
            nb.add(t4, text=" 📋 GSTR-1 ")
            nb.add(t5, text=" 🧾 GSTR-3B ")
            nb.add(t6, text=" 📑 GSTR-2B ")
            nb.add(t8, text=" 📔 GSTR-9 ")
            nb.add(t9, text=" 🔄 Reclaim ")
            nb.add(t10, text=" 💳 Credit ")
            nb.add(t11, text=" 💵 Cash ")
            nb.add(t7, text=" 🗃 File ")
            self._build_collector(t1)
            self._build_compressor(t2)
            self._build_tax_comparison(t3)
            self._build_extractor_tab(t4, kind="gstr1")
            self._build_extractor_tab(t5, kind="gstr3b")
            self._build_gstr2b_tab(t6)
            self._build_gstr9_9c_tab(t8)
            self._build_ecrrs_tab(t9)
            self._build_ecl_tab(t10)
            self._build_ecashl_tab(t11)
            self._build_file_mgmt_tab(t7)

        def _sec(self, p, t):
            tk.Label(p, text=t, font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", padx=20, pady=(6, 2))

        def _entry(self, parent, var):
            return tk.Entry(parent, textvariable=var,
                            font=("Segoe UI", 10), bg=ENTRY_BG, fg=TEXT,
                            insertbackground=TEXT, relief="flat",
                            highlightthickness=1, highlightbackground=BORDER,
                            highlightcolor=ACCENT)

        def _btn(self, parent, text, cmd, bg=ACCENT, fg="white",
                 abg=None, font=("Segoe UI", 10, "bold"), pady=6, padx=14, **kw):
            return tk.Button(parent, text=text, command=cmd,
                             font=font, bg=bg, fg=fg,
                             activebackground=abg or ACCENT2,
                             activeforeground=fg, relief="flat",
                             cursor="hand2", padx=padx, pady=pady, **kw)

        def _logbox(self, parent):
            f = tk.Frame(parent, bg=CARD)
            f.pack(fill="both", padx=20, pady=(0, 16), expand=True)
            box = tk.Text(f, bg=ENTRY_BG, fg=TEXT, font=("Consolas", 9),
                          relief="flat", wrap="word", height=8,
                          insertbackground=TEXT,
                          highlightthickness=1, highlightbackground=BORDER,
                          state="disabled")
            box.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(f, command=box.yview)
            sb.pack(side="right", fill="y")
            box.configure(yscrollcommand=sb.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                box.tag_config(tag, foreground=color,
                               font=("Consolas", 9, weight))
            return box

        def _log(self, box, msg, tag=""):
            box.configure(state="normal")
            box.insert("end", msg + "\n", tag)
            box.see("end")
            box.configure(state="disabled")

        def _clear(self, box):
            box.configure(state="normal")
            box.delete("1.0", "end")
            box.configure(state="disabled")

        # ════════════════════════════════════════════════════
        #  TAB 1 — FILE COLLECTOR
        # ════════════════════════════════════════════════════
        def _build_collector(self, tab):
            # ─── PIN run button + log to the BOTTOM first ──────────
            # Anything packed with side="bottom" gets placed before
            # widgets packed later, so it stays at the bottom edge
            # even when the window is shrunk.

            # Bottom: Log section
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log1 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log1.pack(side="left", fill="both", expand=True)
            sb1 = ttk.Scrollbar(log_frame, command=self.log1.yview)
            sb1.pack(side="right", fill="y")
            self.log1.configure(yscrollcommand=sb1.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log1.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            # Bottom (above log): RUN button — BIG green
            self.run_btn1 = tk.Button(tab,
                                      text="▶   RUN — COLLECT ALL FILES",
                                      command=self._run_collect,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn1.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            # ─── Top: form fields (fill remaining space above) ──────
            top = tk.Frame(tab, bg=CARD)
            top.pack(side="top", fill="x", padx=0, pady=(0, 0))

            self._sec(top, "📁  Source Folder")
            rf = tk.Frame(top, bg=CARD); rf.pack(fill="x", padx=20, pady=(0, 8))
            self.folder_var = tk.StringVar()
            self._entry(rf, self.folder_var).pack(side="left", fill="x",
                                                  expand=True, ipady=6, padx=(0, 8))
            self._btn(rf, "Browse", self._browse_src).pack(side="left")

            self._sec(top, "📂  Output Folder Name")
            of = tk.Frame(top, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self.outname_var = tk.StringVar(value="GST_All_States_Output")
            self._entry(of, self.outname_var).pack(fill="x", ipady=6)

            opt = tk.Frame(top, bg=CARD); opt.pack(fill="x", padx=20, pady=(6, 0))
            self.lenient_var = tk.BooleanVar(value=True)
            tk.Checkbutton(opt,
                           text="  Lenient match  (GSTR1 matches GSTR-1, GSTR_1, GSTR 1)",
                           variable=self.lenient_var, bg=CARD,
                           activebackground=CARD, fg=TEXT, font=("Segoe UI", 9),
                           selectcolor=ENTRY_BG, cursor="hand2"
                           ).pack(anchor="w")
            self.search_path_var = tk.BooleanVar(value=False)
            tk.Checkbutton(opt,
                           text="  Also match keyword in folder path",
                           variable=self.search_path_var, bg=CARD,
                           activebackground=CARD, fg=TEXT, font=("Segoe UI", 9),
                           selectcolor=ENTRY_BG, cursor="hand2"
                           ).pack(anchor="w", pady=(0, 2))

            self._sec(top, "🔍  Rules  —  keyword + file type")

            h = tk.Frame(top, bg=CARD); h.pack(fill="x", padx=20, pady=(0, 4))
            tk.Label(h, text="Keyword", font=("Segoe UI", 8, "bold"),
                     bg=CARD, fg=SUBTEXT, anchor="w", width=22).pack(side="left")
            tk.Label(h, text="File type", font=("Segoe UI", 8, "bold"),
                     bg=CARD, fg=SUBTEXT, anchor="w", width=22
                     ).pack(side="left", padx=(8, 0))
            tk.Label(h, text="Custom ext", font=("Segoe UI", 8, "bold"),
                     bg=CARD, fg=SUBTEXT, anchor="w", width=18
                     ).pack(side="left", padx=(8, 0))
            tk.Label(h, text="State prefix", font=("Segoe UI", 8, "bold"),
                     bg=CARD, fg=SUBTEXT).pack(side="left", padx=(8, 0))

            self.rows_frame = tk.Frame(top, bg=CARD)
            self.rows_frame.pack(fill="x", padx=20, pady=(0, 4))

            for kw, ext_label in [("GSTR1",  "PDF (.pdf)"),
                                  ("GSTR3B", "PDF (.pdf)"),
                                  ("GSTR2A", "Excel (.xlsx, .xls)"),
                                  ("GSTR2B", "Excel (.xlsx, .xls)")]:
                self._add_row(kw, ext_label)

            self._btn(top, "＋  Add Rule",
                      lambda: self._add_row("", "Any (all formats)"),
                      bg=CARD, fg=ACCENT2, abg=CARD
                      ).pack(anchor="w", padx=20, pady=(0, 4))

        def _add_row(self, kw="", ext_label="Any (all formats)"):
            row = tk.Frame(self.rows_frame, bg=CARD)
            row.pack(fill="x", pady=3)

            kw_v = tk.StringVar(value=kw)
            ext_v = tk.StringVar(value=ext_label)
            custom_v = tk.StringVar()
            pfx_v = tk.BooleanVar(value=False)

            tk.Entry(row, textvariable=kw_v, font=("Segoe UI", 10),
                     bg=ENTRY_BG, fg=TEXT, insertbackground=TEXT,
                     relief="flat", highlightthickness=1,
                     highlightbackground=BORDER, highlightcolor=ACCENT,
                     width=22).pack(side="left", ipady=5)

            ext_cb = ttk.Combobox(row, textvariable=ext_v, width=20,
                                  values=list(EXT_PRESETS.keys()),
                                  state="readonly", font=("Segoe UI", 9))
            ext_cb.pack(side="left", padx=(8, 0), ipady=2)

            custom_entry = tk.Entry(row, textvariable=custom_v,
                                    font=("Segoe UI", 9),
                                    bg=ENTRY_BG, fg=TEXT, insertbackground=TEXT,
                                    relief="flat", highlightthickness=1,
                                    highlightbackground=BORDER,
                                    highlightcolor=ACCENT, width=18)
            custom_entry.pack(side="left", padx=(8, 0), ipady=4)

            def _toggle_custom(*_):
                if ext_v.get() == "Custom...":
                    custom_entry.configure(state="normal")
                else:
                    custom_v.set("")
                    custom_entry.configure(state="disabled")
            ext_v.trace_add("write", _toggle_custom)
            _toggle_custom()

            tk.Checkbutton(row, variable=pfx_v, bg=CARD,
                           activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2"
                           ).pack(side="left", padx=(16, 0))

            entry = (kw_v, ext_v, custom_v, pfx_v, row)

            def _del():
                row.destroy()
                if entry in self.rows: self.rows.remove(entry)

            tk.Button(row, text="✕", font=("Segoe UI", 9), bg=CARD, fg=ERROR,
                      activebackground=CARD, relief="flat", cursor="hand2",
                      command=_del, padx=6).pack(side="left", padx=(12, 0))
            self.rows.append(entry)

        def _browse_src(self):
            d = filedialog.askdirectory()
            if d: self.folder_var.set(d)

        def _resolve_ext_filter(self, ext_label, custom_text):
            preset = EXT_PRESETS.get(ext_label)
            if preset is None: return None
            if preset == "CUSTOM":
                parts = [p.strip().lower() for p in custom_text.split(",") if p.strip()]
                parts = [p if p.startswith(".") else "." + p for p in parts]
                return tuple(parts) if parts else None
            return preset

        def _run_collect(self):
            mf = self.folder_var.get().strip()
            if not mf or not os.path.isdir(mf):
                messagebox.showerror("Error", "Select a valid source folder.")
                return
            rules = []
            for kw_v, ext_v, custom_v, pfx_v, _ in self.rows:
                kw = kw_v.get().strip()
                if not kw: continue
                ext_filter = self._resolve_ext_filter(ext_v.get(),
                                                     custom_v.get().strip())
                rules.append((kw, ext_filter, ext_v.get(), pfx_v.get()))
            if not rules:
                messagebox.showerror("Error", "Add at least one keyword rule.")
                return
            out_name = self.outname_var.get().strip() or "GST_All_States_Output"
            self.run_btn1.configure(state="disabled", text="⏳  RUNNING…")
            self._clear(self.log1)
            threading.Thread(target=self._collect_worker,
                             args=(mf, out_name, rules,
                                   self.lenient_var.get(),
                                   self.search_path_var.get()),
                             daemon=True).start()

        def _collect_worker(self, mf, out_name, rules, lenient, search_path):
            try:
                grand = os.path.join(os.path.dirname(mf), out_name)
                os.makedirs(grand, exist_ok=True)
                total = 0
                for kw, ext_filter, ext_label, add_prefix in rules:
                    ext_show = ext_label if ext_filter is not None else "all formats"
                    self._log(self.log1,
                              f"\n▶  Keyword '{kw}'  ·  {ext_show}", "hdr")
                    safe = re.sub(r"[\\/:*?\"<>|]", "_",
                                  f"{kw}_{ext_show.split(' ')[0]}")
                    dest = os.path.join(grand, safe)
                    os.makedirs(dest, exist_ok=True)
                    copied = 0; seen = set()
                    for sf in sorted(os.listdir(mf)):
                        sp = os.path.join(mf, sf)
                        if not os.path.isdir(sp): continue
                        for root, _, files in os.walk(sp):
                            for fn in files:
                                if not _ext_matches(fn, ext_filter): continue
                                target = fn
                                if search_path:
                                    rel = os.path.relpath(root, mf)
                                    target = rel + os.sep + fn
                                if lenient:
                                    if not _keyword_matches(target, kw): continue
                                else:
                                    if kw.lower() not in target.lower(): continue
                                src = os.path.join(root, fn)
                                new = f"{sf}_{fn}" if add_prefix else fn
                                dst = os.path.join(dest, new)
                                if dst in seen or os.path.exists(dst):
                                    base, ext = os.path.splitext(
                                        new if add_prefix else f"{sf}_{fn}")
                                    new = f"{base}{ext}"
                                    dst = os.path.join(dest, new)
                                    i = 1
                                    while dst in seen or os.path.exists(dst):
                                        new = f"{base}_{i}{ext}"
                                        dst = os.path.join(dest, new)
                                        i += 1
                                    self._log(self.log1,
                                              f"   ⚠  {sf} → {new}", "ren")
                                else:
                                    self._log(self.log1,
                                              f"   ✓  {sf} → {new}", "ok")
                                shutil.copy2(src, dst)
                                seen.add(dst); copied += 1
                    self._log(self.log1,
                              "   (no matching files)" if not copied
                              else f"   → {copied} file(s) copied", "dim")
                    total += copied
                self._log(self.log1,
                          f"\n✅  Done!  {total} files → {grand}", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(grand)
                except Exception: pass
            except Exception as e:
                self._log(self.log1,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn1.configure(state="normal",
                                        text="▶   RUN — COLLECT ALL FILES")

        # ════════════════════════════════════════════════════
        #  TAB 2 — PDF COMPRESSOR
        # ════════════════════════════════════════════════════
        def _build_compressor(self, tab):
            """PDF Tools tab — 3 modes: Compress / Merge / Split.
            Mode selector at top, dynamic UI below, single Run button."""

            # ─── Pin log + run button to bottom first ───────
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 10))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log2 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log2.pack(side="left", fill="both", expand=True)
            sb2 = ttk.Scrollbar(log_frame, command=self.log2.yview)
            sb2.pack(side="right", fill="y")
            self.log2.configure(yscrollcommand=sb2.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log2.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            self.run_btn2 = tk.Button(tab,
                                      text="🗜   COMPRESS PDF(S)",
                                      command=self._run_pdf_tool,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn2.pack(side="bottom", fill="x", padx=20, pady=(4, 6))

            # ─── Status banner ───────────────────────────────
            if HAS_PYMUPDF:
                if HAS_PIKEPDF:
                    txt = "  ✅  Ready — full features (compress, merge, split)."
                else:
                    txt = ("  ✅  Ready (basic mode). pikepdf missing → some "
                           "advanced compression options disabled.")
                bgc = "#0f2a1a"; fgc = SUCCESS
            else:
                txt = ("  ⚠  PyMuPDF not installed.  In CMD run:\n"
                       "       pip install PyMuPDF pikepdf pillow\n"
                       "  Then restart this app.")
                bgc = "#2d1f0e"; fgc = WARNING
            info = tk.Frame(tab, bg=bgc)
            info.pack(side="top", fill="x", padx=20, pady=(10, 4))
            tk.Label(info, text=txt,
                     font=("Segoe UI", 9), bg=bgc, fg=fgc, justify="left",
                     wraplength=700).pack(padx=10, pady=6, anchor="w")

            # ─── Mode selector ──────────────────────────────
            self.pdf_mode_var = tk.StringVar(value="compress")
            mode_frame = tk.Frame(tab, bg=CARD)
            mode_frame.pack(side="top", fill="x", padx=20, pady=(4, 6))
            tk.Label(mode_frame, text="Mode:",
                     font=("Segoe UI", 10, "bold"),
                     bg=CARD, fg=ACCENT2).pack(side="left", padx=(0, 10))
            for label, val, accent in [
                ("📦  Compress", "compress", SUCCESS),
                ("🔗  Merge",    "merge",    ACCENT),
                ("✂  Split",     "split",    WARNING),
            ]:
                tk.Radiobutton(mode_frame, text=f"  {label}  ",
                               variable=self.pdf_mode_var, value=val,
                               command=self._pdf_mode_change,
                               bg=CARD, activebackground=CARD, fg=TEXT,
                               selectcolor=ENTRY_BG, cursor="hand2",
                               font=("Segoe UI", 10, "bold")
                               ).pack(side="left", padx=(0, 6))

            # ─── Dynamic input container (scrollable) ────────
            # Use a Canvas + scrollbar so any of the 3 modes can fit
            # comfortably regardless of how many files are in the merge list.
            self.pdf_dyn_outer = tk.Frame(tab, bg=CARD)
            self.pdf_dyn_outer.pack(side="top", fill="both", expand=True,
                                    padx=0, pady=(0, 4))
            self.pdf_dyn_canvas = tk.Canvas(
                self.pdf_dyn_outer, bg=CARD,
                highlightthickness=0, bd=0)
            self.pdf_dyn_canvas.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(self.pdf_dyn_outer, orient="vertical",
                                command=self.pdf_dyn_canvas.yview)
            sb.pack(side="right", fill="y")
            self.pdf_dyn_canvas.configure(yscrollcommand=sb.set)
            # Inner frame that we actually pack content into
            self.pdf_dyn = tk.Frame(self.pdf_dyn_canvas, bg=CARD)
            self.pdf_dyn_win = self.pdf_dyn_canvas.create_window(
                (0, 0), window=self.pdf_dyn, anchor="nw")
            def _on_dyn_configure(event):
                # Update scroll region whenever content size changes
                self.pdf_dyn_canvas.configure(
                    scrollregion=self.pdf_dyn_canvas.bbox("all"))
            self.pdf_dyn.bind("<Configure>", _on_dyn_configure)
            def _on_canvas_configure(event):
                # Make inner frame width match canvas width
                self.pdf_dyn_canvas.itemconfigure(
                    self.pdf_dyn_win, width=event.width)
            self.pdf_dyn_canvas.bind("<Configure>", _on_canvas_configure)
            # Mouse-wheel scroll support
            def _on_mousewheel(event):
                # Windows / Mac: event.delta;  Linux: event.num (4/5)
                if hasattr(event, 'delta') and event.delta:
                    self.pdf_dyn_canvas.yview_scroll(
                        int(-event.delta / 120), "units")
                elif getattr(event, 'num', None) == 4:
                    self.pdf_dyn_canvas.yview_scroll(-1, "units")
                elif getattr(event, 'num', None) == 5:
                    self.pdf_dyn_canvas.yview_scroll(1, "units")
            for w in (self.pdf_dyn_canvas, self.pdf_dyn):
                w.bind("<MouseWheel>", _on_mousewheel)
                w.bind("<Button-4>", _on_mousewheel)
                w.bind("<Button-5>", _on_mousewheel)

            # Initialize shared state vars
            self.pdf_src_var = tk.StringVar()
            self.pdf_out_var = tk.StringVar()
            self.target_mb_var = tk.StringVar(value="1.0")
            self.mode_var = tk.StringVar(value="preserve_text")
            # Merge state
            self.merge_files = []   # list of file paths in order
            self.merge_out_var = tk.StringVar()           # for single mode (save-as)
            self.merge_out_dir_var = tk.StringVar()       # for multi mode (folder)
            self.merge_prefix_var = tk.StringVar(value="Merged_part")
            self.merge_target_mb_var = tk.StringVar(value="5.0")
            self.merge_autocompress_var = tk.BooleanVar(value=False)  # optional
            self.merge_strategy_var = tk.StringVar(value="single")    # single / multi
            self.merge_preserve_order_var = tk.BooleanVar(value=False)
            # Split state
            self.split_in_var = tk.StringVar()
            self.split_out_var = tk.StringVar()
            self.split_method_var = tk.StringVar(value="to_fit")
            self.split_target_mb_var = tk.StringVar(value="5.0")
            self.split_pages_per_var = tk.StringVar(value="5")
            self.split_ranges_var = tk.StringVar()

            # Render initial mode (compress)
            self._pdf_render_compress()

        def _pdf_mode_change(self):
            """User flipped the mode radio — re-render the dynamic area."""
            mode = self.pdf_mode_var.get()
            # Update Run button text
            txts = {
                "compress": "🗜   COMPRESS PDF(S)",
                "merge":    "🔗   MERGE PDFs",
                "split":    "✂   SPLIT PDF",
            }
            self.run_btn2.configure(text=txts.get(mode, "RUN"))
            # Clear dynamic area
            for child in self.pdf_dyn.winfo_children():
                child.destroy()
            # Render new mode
            if mode == "compress": self._pdf_render_compress()
            elif mode == "merge":   self._pdf_render_merge()
            else:                   self._pdf_render_split()

        # ─── Mode renderer: Compress (original UI) ───────────
        def _pdf_render_compress(self):
            p = self.pdf_dyn
            self._sec(p, "📄  Select PDF File  OR  Folder of PDFs")
            sf = tk.Frame(p, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self._entry(sf, self.pdf_src_var).pack(side="left", fill="x",
                                                   expand=True, ipady=6,
                                                   padx=(0, 8))
            self._btn(sf, "File",   self._browse_pdf_file
                      ).pack(side="left", padx=(0, 6))
            self._btn(sf, "Folder", self._browse_pdf_folder
                      ).pack(side="left")

            self._sec(p, "📂  Output Folder")
            of2 = tk.Frame(p, bg=CARD); of2.pack(fill="x", padx=20, pady=(0, 8))
            self._entry(of2, self.pdf_out_var).pack(side="left", fill="x",
                                                    expand=True, ipady=6,
                                                    padx=(0, 8))
            self._btn(of2, "Browse", self._browse_pdf_out).pack(side="left")

            self._sec(p, "🎯  Target File Size")
            sz = tk.Frame(p, bg=CARD); sz.pack(fill="x", padx=20, pady=(0, 4))
            tk.Label(sz, text="Target:", font=("Segoe UI", 10),
                     bg=CARD, fg=TEXT).pack(side="left")
            tk.Entry(sz, textvariable=self.target_mb_var, width=6,
                     font=("Segoe UI", 11, "bold"), bg=ENTRY_BG, fg=SUCCESS,
                     insertbackground=TEXT, relief="flat",
                     highlightthickness=2, highlightbackground=ACCENT,
                     highlightcolor=ACCENT).pack(side="left", padx=6, ipady=4)
            tk.Label(sz, text="MB", font=("Segoe UI", 10),
                     bg=CARD, fg=TEXT).pack(side="left", padx=(0, 12))
            for label, val in [("< 500KB", "0.5"), ("< 1MB", "1.0"),
                               ("< 2MB", "2.0"), ("< 5MB", "5.0")]:
                self._btn(sz, label, lambda v=val: self.target_mb_var.set(v),
                          font=("Segoe UI", 9, "bold"), pady=3, padx=8
                          ).pack(side="left", padx=(0, 4))

            mc = tk.Frame(p, bg=CARD); mc.pack(fill="x", padx=20, pady=(4, 0))
            tk.Radiobutton(mc,
                           text="  Preserve searchable text (recommended)",
                           variable=self.mode_var, value="preserve_text",
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9, "bold")
                           ).pack(anchor="w")
            tk.Radiobutton(mc,
                           text="  Allow text loss to hit target (last resort)",
                           variable=self.mode_var, value="allow_text_loss",
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9, "bold")
                           ).pack(anchor="w", pady=(2, 0))

        # ─── Mode renderer: Merge ─────────────────────────────
        def _pdf_render_merge(self):
            p = self.pdf_dyn
            self._sec(p, "📑  PDFs to Merge")
            list_frame = tk.Frame(p, bg=CARD)
            list_frame.pack(fill="x", padx=20, pady=(0, 4))

            # Listbox with scrollbar — smaller height for compact layout
            lb_frame = tk.Frame(list_frame, bg=CARD)
            lb_frame.pack(side="left", fill="both", expand=True)
            self.merge_lb = tk.Listbox(
                lb_frame, height=4,
                bg=ENTRY_BG, fg=TEXT, font=("Segoe UI", 9),
                selectmode="extended",
                highlightthickness=1, highlightbackground=BORDER,
                relief="flat")
            self.merge_lb.pack(side="left", fill="both", expand=True)
            lb_sb = ttk.Scrollbar(lb_frame, command=self.merge_lb.yview)
            lb_sb.pack(side="right", fill="y")
            self.merge_lb.configure(yscrollcommand=lb_sb.set)
            self._merge_refresh_list()

            # Right-side buttons column (compact)
            bf = tk.Frame(list_frame, bg=CARD)
            bf.pack(side="left", fill="y", padx=(8, 0))
            self._btn(bf, "+  Add", self._merge_add_files,
                      font=("Segoe UI", 9, "bold"), pady=2, padx=6
                      ).pack(fill="x", pady=(0, 2))
            self._btn(bf, "–  Remove", self._merge_remove_selected,
                      font=("Segoe UI", 9), pady=2, padx=6
                      ).pack(fill="x", pady=(0, 2))
            self._btn(bf, "↑  Up", self._merge_move_up,
                      font=("Segoe UI", 9), pady=2, padx=6
                      ).pack(fill="x", pady=(0, 2))
            self._btn(bf, "↓  Down", self._merge_move_down,
                      font=("Segoe UI", 9), pady=2, padx=6
                      ).pack(fill="x", pady=(0, 2))
            self._btn(bf, "✕  Clear", self._merge_clear,
                      font=("Segoe UI", 9), pady=2, padx=6
                      ).pack(fill="x", pady=(0, 2))

            # ─── Live estimate label — right under file list ─
            self.merge_estimate_lbl = tk.Label(
                p, text="📊  Add files to see size estimate",
                font=("Segoe UI", 9, "bold"),
                bg="#0f1830", fg=ACCENT2, anchor="w", justify="left",
                wraplength=720, padx=8, pady=4)
            self.merge_estimate_lbl.pack(fill="x", padx=20, pady=(4, 6))

            # ─── Strategy radio + target MB ─────────────────
            st = tk.Frame(p, bg=CARD); st.pack(fill="x", padx=20, pady=(0, 2))
            tk.Label(st, text="🎯  Strategy:",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w")
            tk.Radiobutton(
                st, text="  Single merged file (all → one output)",
                variable=self.merge_strategy_var, value="single",
                command=self._merge_strategy_change,
                bg=CARD, activebackground=CARD, fg=TEXT,
                selectcolor=ENTRY_BG, cursor="hand2",
                font=("Segoe UI", 9)
                ).pack(anchor="w", padx=(12, 0))

            mr = tk.Frame(p, bg=CARD); mr.pack(fill="x", padx=32, pady=(0, 0))
            tk.Radiobutton(
                mr, text="  Multiple files — auto-bucket under",
                variable=self.merge_strategy_var, value="multi",
                command=self._merge_strategy_change,
                bg=CARD, activebackground=CARD, fg=TEXT,
                selectcolor=ENTRY_BG, cursor="hand2",
                font=("Segoe UI", 9)
                ).pack(side="left")
            tk.Entry(mr, textvariable=self.merge_target_mb_var, width=5,
                     font=("Segoe UI", 10, "bold"), bg=ENTRY_BG, fg=SUCCESS,
                     insertbackground=TEXT, relief="flat",
                     highlightthickness=2, highlightbackground=ACCENT,
                     highlightcolor=ACCENT).pack(side="left", padx=4, ipady=1)
            tk.Label(mr, text="MB each  (GST portal: 5)",
                     font=("Segoe UI", 8, "italic"),
                     bg=CARD, fg=SUBTEXT).pack(side="left")

            # Preserve order toggle
            po = tk.Frame(p, bg=CARD); po.pack(fill="x", padx=32, pady=(0, 4))
            tk.Checkbutton(
                po, text="  Preserve input order  "
                        "(else best-fit by size for tighter packing)",
                variable=self.merge_preserve_order_var,
                command=self._merge_update_estimate,
                bg=CARD, activebackground=CARD, fg=TEXT,
                selectcolor=ENTRY_BG, cursor="hand2",
                font=("Segoe UI", 8)
                ).pack(anchor="w")

            # ─── Dynamic Output area ────────────────────────
            self.merge_output_frame = tk.Frame(p, bg=CARD)
            self.merge_output_frame.pack(fill="x", pady=(2, 0))
            self._merge_render_output()

            # ─── Compression (optional) — compact ─────────────
            cf = tk.Frame(p, bg=CARD); cf.pack(fill="x", padx=20, pady=(4, 2))
            tk.Checkbutton(
                cf, text="  ⚙ Auto-compress (lossless only, fast) any output "
                         "exceeding target  ·  use 📦 Compress for heavy",
                variable=self.merge_autocompress_var,
                bg=CARD, activebackground=CARD, fg=TEXT,
                selectcolor=ENTRY_BG, cursor="hand2",
                font=("Segoe UI", 9)
                ).pack(anchor="w")

            # Wire target MB live updates → re-estimate
            try:
                self.merge_target_mb_var.trace_add(
                    "write", lambda *a: self._merge_update_estimate())
            except AttributeError:
                self.merge_target_mb_var.trace(
                    "w", lambda *a: self._merge_update_estimate())

            self._merge_update_estimate()

        def _merge_render_output(self):
            """Render output controls based on current strategy."""
            for child in self.merge_output_frame.winfo_children():
                child.destroy()
            strat = self.merge_strategy_var.get()
            if strat == "single":
                lbl = tk.Label(self.merge_output_frame, text="💾  Output:",
                               font=("Segoe UI", 9, "bold"),
                               bg=CARD, fg=ACCENT2)
                lbl.pack(anchor="w", padx=20, pady=(2, 0))
                of = tk.Frame(self.merge_output_frame, bg=CARD)
                of.pack(fill="x", padx=20, pady=(0, 2))
                self._entry(of, self.merge_out_var).pack(
                    side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
                self._btn(of, "Save as", self._browse_merge_out
                          ).pack(side="left")
            else:
                lbl = tk.Label(
                    self.merge_output_frame,
                    text="💾  Output folder & prefix  "
                         "(each bucket → one PDF):",
                    font=("Segoe UI", 9, "bold"),
                    bg=CARD, fg=ACCENT2)
                lbl.pack(anchor="w", padx=20, pady=(2, 0))
                of = tk.Frame(self.merge_output_frame, bg=CARD)
                of.pack(fill="x", padx=20, pady=(0, 2))
                self._entry(of, self.merge_out_dir_var).pack(
                    side="left", fill="x", expand=True, ipady=4, padx=(0, 6))
                self._btn(of, "Browse", self._browse_merge_out_dir
                          ).pack(side="left")
                pf = tk.Frame(self.merge_output_frame, bg=CARD)
                pf.pack(fill="x", padx=20, pady=(2, 2))
                tk.Label(pf, text="Prefix:", font=("Segoe UI", 9),
                         bg=CARD, fg=TEXT).pack(side="left")
                self._entry(pf, self.merge_prefix_var).pack(
                    side="left", fill="x", expand=True, ipady=3, padx=(6, 0))

        def _merge_strategy_change(self):
            """Strategy radio changed — re-render output area and estimate."""
            self._merge_render_output()
            self._merge_update_estimate()

        def _merge_refresh_list(self):
            self.merge_lb.delete(0, "end")
            for i, fp in enumerate(self.merge_files, 1):
                pages, size = get_pdf_info(fp)
                size_mb = size / (1024 * 1024)
                self.merge_lb.insert(
                    "end",
                    f" {i:2d}.  {os.path.basename(fp):<50}  "
                    f"({pages} pp, {size_mb:.2f} MB)")

        def _merge_update_estimate(self):
            if not self.merge_files:
                self.merge_estimate_lbl.configure(
                    text="📊  Add files to see size estimate")
                return
            total_bytes = 0
            total_pages = 0
            for fp in self.merge_files:
                pages, size = get_pdf_info(fp)
                total_bytes += size
                total_pages += pages
            total_mb = total_bytes / (1024 * 1024)
            try:
                target_mb = float(self.merge_target_mb_var.get())
                if target_mb <= 0: target_mb = 5.0
            except Exception:
                target_mb = 5.0

            strat = self.merge_strategy_var.get()

            if strat == "single":
                # Single merged file estimates (as before)
                est_merge_mb = total_mb * 0.97       # tiny dedup
                est_lossless_mb = est_merge_mb * 0.65
                est_aggressive_mb = est_merge_mb * 0.25
                fits_now = est_merge_mb <= target_mb
                fits_lossless = est_lossless_mb <= target_mb
                status = ("✅ Will fit under target" if fits_now
                          else "○ Lossless compress should fit"
                          if fits_lossless else
                          "⚠ May need text-loss compression to fit")
                self.merge_estimate_lbl.configure(
                    text=(f"📊  {len(self.merge_files)} files · "
                          f"{total_pages} pages · "
                          f"Sum {total_mb:.2f} MB  →  "
                          f"Expected merged ≈ {est_merge_mb:.2f} MB  "
                          f"·  Lossless ≈ {est_lossless_mb:.2f} MB  "
                          f"·  Max compress ≈ {est_aggressive_mb:.2f} MB\n"
                          f"     Target: {target_mb} MB  ·  {status}"))
            else:
                # Multi-bucket — compute bucketing and show preview
                preserve = self.merge_preserve_order_var.get()
                buckets, oversize = compute_merge_buckets(
                    self.merge_files, target_mb, preserve_order=preserve)
                # Build summary text
                lines = []
                lines.append(
                    f"📊  {len(self.merge_files)} files · "
                    f"{total_pages} pages · "
                    f"Sum {total_mb:.2f} MB  →  Bucketing into ≤ "
                    f"{target_mb} MB each:")
                for i, bucket in enumerate(buckets, 1):
                    btotal = sum(b[1] for b in bucket) / (1024 * 1024)
                    over_flag = ""
                    if len(bucket) == 1 and bucket[0][1] > target_mb * 1024 * 1024:
                        over_flag = "  ⚠ over target (single file)"
                    lines.append(
                        f"   Part {i:02d}: {len(bucket)} file"
                        f"{'s' if len(bucket) != 1 else ''}, "
                        f"~{btotal:.2f} MB{over_flag}")
                # Summary status
                if oversize:
                    if self.merge_autocompress_var.get():
                        status = (f"⚙ {len(oversize)} file(s) alone over target "
                                  "— will be compressed if Auto-compress is on")
                    else:
                        status = (f"⚠ {len(oversize)} file(s) alone over target "
                                  "— enable Auto-compress or pre-compress them")
                else:
                    status = (f"✅ All {len(buckets)} output files will be ≤ "
                              f"{target_mb} MB  (no compression needed)")
                lines.append(f"     {status}")
                self.merge_estimate_lbl.configure(text="\n".join(lines))

        def _browse_merge_out_dir(self):
            d = filedialog.askdirectory()
            if d: self.merge_out_dir_var.set(d)

        def _merge_add_files(self):
            files = filedialog.askopenfilenames(
                filetypes=[("PDF files", "*.pdf")],
                title="Pick PDFs to merge (will be added in selection order)")
            if not files: return
            for fp in files:
                if fp not in self.merge_files:
                    self.merge_files.append(fp)
            self._merge_refresh_list()
            self._merge_update_estimate()

        def _merge_remove_selected(self):
            sel = list(self.merge_lb.curselection())
            if not sel: return
            for idx in sorted(sel, reverse=True):
                del self.merge_files[idx]
            self._merge_refresh_list()
            self._merge_update_estimate()

        def _merge_move_up(self):
            sel = list(self.merge_lb.curselection())
            if not sel: return
            for idx in sel:
                if idx > 0:
                    self.merge_files[idx - 1], self.merge_files[idx] = (
                        self.merge_files[idx], self.merge_files[idx - 1])
            self._merge_refresh_list()
            # Re-select shifted items
            for idx in sel:
                if idx > 0:
                    self.merge_lb.selection_set(idx - 1)

        def _merge_move_down(self):
            sel = list(self.merge_lb.curselection())
            if not sel: return
            for idx in sorted(sel, reverse=True):
                if idx < len(self.merge_files) - 1:
                    self.merge_files[idx + 1], self.merge_files[idx] = (
                        self.merge_files[idx], self.merge_files[idx + 1])
            self._merge_refresh_list()
            for idx in sel:
                if idx < len(self.merge_files) - 1:
                    self.merge_lb.selection_set(idx + 1)

        def _merge_clear(self):
            self.merge_files = []
            self._merge_refresh_list()
            self._merge_update_estimate()

        def _browse_merge_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile="Merged.pdf")
            if p: self.merge_out_var.set(p)

        # ─── Mode renderer: Split ─────────────────────────────
        def _pdf_render_split(self):
            p = self.pdf_dyn
            self._sec(p, "📄  Input PDF File")
            sf = tk.Frame(p, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 6))
            self._entry(sf, self.split_in_var).pack(side="left", fill="x",
                                                     expand=True, ipady=6,
                                                     padx=(0, 8))
            self._btn(sf, "Browse", self._browse_split_in).pack(side="left")

            self.split_info_lbl = tk.Label(
                p, text="📊  Pick a PDF to see info",
                font=("Segoe UI", 10, "bold"),
                bg=CARD, fg=ACCENT2, anchor="w", wraplength=720)
            self.split_info_lbl.pack(fill="x", padx=20, pady=(0, 6))
            self._split_update_info()

            self._sec(p, "🛠  Split method")
            mf = tk.Frame(p, bg=CARD); mf.pack(fill="x", padx=20, pady=(0, 4))

            # Auto-fit method (default)
            af = tk.Frame(mf, bg=CARD); af.pack(fill="x", pady=(0, 4))
            tk.Radiobutton(af,
                           text="  Auto-split to fit under",
                           variable=self.split_method_var, value="to_fit",
                           command=self._split_update_info,
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9, "bold")
                           ).pack(side="left")
            tk.Entry(af, textvariable=self.split_target_mb_var, width=5,
                     font=("Segoe UI", 10, "bold"), bg=ENTRY_BG, fg=SUCCESS,
                     insertbackground=TEXT, relief="flat",
                     highlightthickness=2, highlightbackground=ACCENT,
                     highlightcolor=ACCENT).pack(side="left", padx=6, ipady=2)
            tk.Label(af, text="MB per output file  (GST portal: 5)",
                     font=("Segoe UI", 9, "italic"),
                     bg=CARD, fg=SUBTEXT).pack(side="left")

            # By chunk
            cf = tk.Frame(mf, bg=CARD); cf.pack(fill="x", pady=(0, 4))
            tk.Radiobutton(cf,
                           text="  Split every",
                           variable=self.split_method_var, value="by_chunk",
                           command=self._split_update_info,
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9, "bold")
                           ).pack(side="left")
            tk.Entry(cf, textvariable=self.split_pages_per_var, width=4,
                     font=("Segoe UI", 10, "bold"), bg=ENTRY_BG, fg=SUCCESS,
                     insertbackground=TEXT, relief="flat",
                     highlightthickness=2, highlightbackground=ACCENT,
                     highlightcolor=ACCENT).pack(side="left", padx=6, ipady=2)
            tk.Label(cf, text="pages",
                     font=("Segoe UI", 9, "italic"),
                     bg=CARD, fg=SUBTEXT).pack(side="left")

            # By ranges
            rf = tk.Frame(mf, bg=CARD); rf.pack(fill="x", pady=(0, 4))
            tk.Radiobutton(rf,
                           text="  Custom ranges:",
                           variable=self.split_method_var, value="by_ranges",
                           command=self._split_update_info,
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9, "bold")
                           ).pack(side="left")
            self._entry(rf, self.split_ranges_var).pack(side="left",
                                                         fill="x", expand=True,
                                                         padx=(6, 0), ipady=4)
            tk.Label(p, text="     e.g.  1-5, 6-10, 11-20  (each range → one PDF)",
                     font=("Segoe UI", 8, "italic"),
                     bg=CARD, fg=SUBTEXT).pack(anchor="w", padx=20)

            self._sec(p, "📂  Output Folder")
            of = tk.Frame(p, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self._entry(of, self.split_out_var).pack(side="left", fill="x",
                                                      expand=True, ipady=6,
                                                      padx=(0, 8))
            self._btn(of, "Browse", self._browse_split_out).pack(side="left")

        def _browse_split_in(self):
            f = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
            if f:
                self.split_in_var.set(f)
                self._split_update_info()

        def _browse_split_out(self):
            d = filedialog.askdirectory()
            if d: self.split_out_var.set(d)

        def _split_update_info(self):
            fp = self.split_in_var.get().strip()
            if not fp or not os.path.isfile(fp):
                self.split_info_lbl.configure(
                    text="📊  Pick a PDF to see info")
                return
            pages, size = get_pdf_info(fp)
            size_mb = size / (1024 * 1024)
            method = self.split_method_var.get()
            est = ""
            if method == "to_fit":
                try:
                    target = float(self.split_target_mb_var.get())
                except Exception:
                    target = 5.0
                if pages > 0 and size_mb > 0:
                    per_page = size_mb / pages
                    pages_per_chunk = max(1, int((target * 0.85) / per_page))
                    n_chunks = max(1, (pages + pages_per_chunk - 1) // pages_per_chunk)
                    if size_mb <= target:
                        est = (f"  →  Already under {target} MB — will be "
                               "copied as-is (no split needed)")
                    else:
                        est = (f"  →  Estimate: ~{n_chunks} files of "
                               f"~{pages_per_chunk} pages, ~{target * 0.85:.2f} MB each")
            elif method == "by_chunk":
                try:
                    chunk = int(self.split_pages_per_var.get())
                    n_chunks = max(1, (pages + chunk - 1) // chunk)
                    size_per_chunk = size_mb * (chunk / max(pages, 1))
                    est = (f"  →  Estimate: {n_chunks} files of {chunk} pages, "
                           f"~{size_per_chunk:.2f} MB each")
                except Exception:
                    est = ""
            elif method == "by_ranges":
                rstr = self.split_ranges_var.get().strip()
                if rstr:
                    try:
                        ranges = parse_page_ranges(rstr, pages)
                        est = f"  →  Will produce {len(ranges)} file(s)"
                    except Exception as e:
                        est = f"  →  ⚠ {e}"
            self.split_info_lbl.configure(
                text=f"📊  {os.path.basename(fp)}  ·  "
                     f"{pages} pages  ·  {size_mb:.2f} MB" + est)

        # ─── Dispatcher ──────────────────────────────────────
        def _run_pdf_tool(self):
            mode = self.pdf_mode_var.get()
            if mode == "compress": return self._run_compress()
            if mode == "merge":    return self._run_merge()
            if mode == "split":    return self._run_split()

        def _browse_pdf_file(self):
            f = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
            if f: self.pdf_src_var.set(f)
        def _browse_pdf_folder(self):
            d = filedialog.askdirectory()
            if d: self.pdf_src_var.set(d)
        def _browse_pdf_out(self):
            d = filedialog.askdirectory()
            if d: self.pdf_out_var.set(d)

        def _run_compress(self):
            if not HAS_PYMUPDF:
                messagebox.showerror(
                    "PyMuPDF missing",
                    "PyMuPDF is not installed.  Open CMD and run:\n\n"
                    "    pip install PyMuPDF pikepdf pillow\n\n"
                    "Then restart this app.")
                return
            src = self.pdf_src_var.get().strip()
            if not src or not os.path.exists(src):
                messagebox.showerror("Error", "Select a PDF file or folder.")
                return
            try:
                target_mb = float(self.target_mb_var.get())
                if target_mb <= 0: raise ValueError()
            except Exception:
                messagebox.showerror("Error",
                                     "Enter a valid positive number for target size.")
                return
            out = self.pdf_out_var.get().strip()
            if not out:
                base = src if os.path.isdir(src) else os.path.dirname(src)
                out = os.path.join(base, "Compressed")
            try: os.makedirs(out, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create output folder:\n{e}")
                return
            if os.path.isfile(src):
                pdfs = [src]
            else:
                pdfs = [os.path.join(r, f)
                        for r, _, fs in os.walk(src)
                        for f in fs if f.lower().endswith(".pdf")]
            if not pdfs:
                messagebox.showerror("Error", "No PDF files found.")
                return
            allow_text_loss = (self.mode_var.get() == "allow_text_loss")
            self.run_btn2.configure(state="disabled", text="⏳  COMPRESSING…")
            self._clear(self.log2)
            threading.Thread(target=self._compress_worker,
                             args=(pdfs, out, target_mb, allow_text_loss),
                             daemon=True).start()

        # ─── Merge runner ────────────────────────────────────
        def _run_merge(self):
            if not (HAS_PYMUPDF or HAS_PIKEPDF):
                messagebox.showerror(
                    "Missing dependency",
                    "PyMuPDF or pikepdf required for merging.\n"
                    "Run: pip install PyMuPDF pikepdf")
                return
            if len(self.merge_files) < 2:
                messagebox.showerror("Error",
                                     "Add at least 2 PDFs to merge.")
                return
            for fp in self.merge_files:
                if not os.path.isfile(fp):
                    messagebox.showerror("Error",
                                         f"File no longer exists:\n{fp}")
                    return
            try:
                target_mb = float(self.merge_target_mb_var.get())
                if target_mb <= 0: raise ValueError()
            except Exception:
                target_mb = 5.0
            auto_compress = self.merge_autocompress_var.get()
            strat = self.merge_strategy_var.get()

            if strat == "single":
                # Single output PDF
                out = self.merge_out_var.get().strip()
                if not out:
                    base = os.path.dirname(self.merge_files[0])
                    out = os.path.join(base, "Merged.pdf")
                if not out.lower().endswith(".pdf"):
                    out += ".pdf"
                try:
                    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
                    if os.path.exists(out):
                        with open(out, "ab"): pass
                except PermissionError:
                    messagebox.showerror("Output locked",
                                         f"Output file is open:\n{out}\n"
                                         "Close it and re-run.")
                    return
                self.run_btn2.configure(state="disabled", text="⏳  MERGING…")
                self._clear(self.log2)
                threading.Thread(target=self._merge_worker,
                                 args=(list(self.merge_files), out, target_mb,
                                       auto_compress),
                                 daemon=True).start()
            else:
                # Multi-bucket output
                out_dir = self.merge_out_dir_var.get().strip()
                if not out_dir:
                    out_dir = os.path.join(
                        os.path.dirname(self.merge_files[0]), "MergedBuckets")
                try:
                    os.makedirs(out_dir, exist_ok=True)
                except Exception as e:
                    messagebox.showerror("Error",
                                         f"Cannot create output folder:\n{e}")
                    return
                prefix = self.merge_prefix_var.get().strip() or "Merged_part"
                preserve = self.merge_preserve_order_var.get()
                self.run_btn2.configure(state="disabled",
                                        text="⏳  MERGING (multi)…")
                self._clear(self.log2)
                threading.Thread(target=self._merge_worker_multi,
                                 args=(list(self.merge_files), out_dir, prefix,
                                       target_mb, preserve, auto_compress),
                                 daemon=True).start()

        def _merge_worker(self, files, out_path, target_mb, auto_compress):
            try:
                self._log(self.log2,
                          f"Merging {len(files)} PDF(s)...", "hdr")
                total_in = sum(os.path.getsize(f) for f in files)
                self._log(self.log2,
                          f"   Total input size: {total_in/1024/1024:.2f} MB",
                          "dim")

                def cb(i, n, name):
                    self._log(self.log2,
                              f"   [{i}/{n}]  {name}", "dim")
                merge_pdfs(files, out_path, progress_cb=cb)
                out_size = os.path.getsize(out_path)
                out_mb = out_size / (1024 * 1024)
                self._log(self.log2,
                          f"\n   ✓  Merged → {out_mb:.2f} MB  "
                          f"({os.path.basename(out_path)})", "ok")

                # Auto-compress if over target — LOSSLESS ONLY (fast & predictable)
                if auto_compress and out_mb > target_mb:
                    self._log(self.log2,
                              f"\n⚙  Above target ({target_mb} MB). "
                              "Trying lossless compression…", "hdr")
                    # Move merged to temp, compress to original location
                    temp_in = out_path + ".__merge_temp__.pdf"
                    try:
                        os.rename(out_path, temp_in)
                    except Exception:
                        import shutil
                        shutil.copy2(out_path, temp_in)

                    if HAS_PYMUPDF:
                        try:
                            self._log(self.log2,
                                      "   Running lossless deflate "
                                      "(garbage-collect + stream compress)…",
                                      "dim")
                            compress_lossless(temp_in, out_path)
                            after_size = os.path.getsize(out_path)
                            after_mb = after_size / (1024 * 1024)
                            saved_pct = (1 - after_size /
                                         os.path.getsize(temp_in)) * 100
                            self._log(self.log2,
                                      f"   ✓  Lossless done: "
                                      f"{after_mb:.2f} MB  (-{saved_pct:.0f}%)",
                                      "ok")
                        except Exception as e:
                            self._log(self.log2,
                                      f"   Compress error: {e}", "err")
                            if not os.path.exists(out_path):
                                try: os.rename(temp_in, out_path)
                                except Exception: pass
                    else:
                        self._log(self.log2,
                                  "   ⚠ PyMuPDF not available — keeping "
                                  "merged file as-is.", "ren")
                        try:
                            import shutil; shutil.copy2(temp_in, out_path)
                        except Exception: pass

                    after_size = os.path.getsize(out_path)
                    after_mb = after_size / (1024 * 1024)
                    target_bytes = int(target_mb * 1024 * 1024)
                    if after_size > target_bytes:
                        self._log(self.log2,
                                  f"   ⚠  Still over target "
                                  f"({after_mb:.2f} MB).\n"
                                  "   💡 Tip: switch to '📦 Compress' mode and "
                                  "enable 'Allow text loss' for aggressive "
                                  "image-recompression.",
                                  "ren")
                    else:
                        self._log(self.log2,
                                  f"   ✅  Fits target: {after_mb:.2f} MB "
                                  f"≤ {target_mb} MB", "ok")
                    try: os.remove(temp_in)
                    except Exception: pass

                final_mb = os.path.getsize(out_path) / (1024 * 1024)
                self._log(self.log2,
                          f"\n✅  Done!  Output: {out_path}\n"
                          f"   Total: {len(files)} files merged → "
                          f"{final_mb:.2f} MB", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log2,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn2.configure(state="normal",
                                        text="🔗   MERGE PDFs")

        # ─── Multi-bucket merge worker ───────────────────────
        def _merge_worker_multi(self, files, out_dir, prefix, target_mb,
                                preserve_order, auto_compress):
            """Bucket files under target_mb each, then merge each bucket to
            its own output PDF. Optionally auto-compress over-target outputs."""
            try:
                self._log(self.log2,
                          f"Multi-bucket merge: {len(files)} input PDFs · "
                          f"target ≤ {target_mb} MB per output\n"
                          f"   Strategy: "
                          f"{'preserve input order' if preserve_order else 'best-fit (sort by size)'}",
                          "hdr")

                # Compute buckets
                buckets, oversize = compute_merge_buckets(
                    files, target_mb, preserve_order=preserve_order)

                self._log(self.log2,
                          f"\n   Plan: {len(buckets)} output file(s) "
                          f"({len(oversize)} oversize input(s) needing "
                          "compression)", "dim")

                outputs = []
                target_bytes = int(target_mb * 1024 * 1024)

                for i, bucket in enumerate(buckets, 1):
                    bucket_files = [b[0] for b in bucket]
                    bucket_sum = sum(b[1] for b in bucket)
                    out_name = f"{prefix}_{i:02d}.pdf"
                    out_path = os.path.join(out_dir, out_name)

                    self._log(self.log2,
                              f"\n   📦  Part {i:02d}: merging "
                              f"{len(bucket_files)} file(s) "
                              f"(~{bucket_sum/1024/1024:.2f} MB)…", "hdr")

                    if len(bucket_files) == 1:
                        # Single-file bucket — just copy (or compress)
                        import shutil
                        shutil.copy2(bucket_files[0], out_path)
                    else:
                        merge_pdfs(bucket_files, out_path)

                    after_size = os.path.getsize(out_path)
                    after_mb = after_size / (1024 * 1024)

                    over_target = after_size > target_bytes
                    if over_target:
                        if auto_compress:
                            self._log(self.log2,
                                      f"      {after_mb:.2f} MB > target — "
                                      "running lossless compression…", "ren")
                            temp_in = out_path + ".__tmp__.pdf"
                            try:
                                os.rename(out_path, temp_in)
                            except Exception:
                                import shutil
                                shutil.copy2(out_path, temp_in)
                            if HAS_PYMUPDF:
                                try:
                                    compress_lossless(temp_in, out_path)
                                    saved = (1 - os.path.getsize(out_path)
                                             / os.path.getsize(temp_in)) * 100
                                    self._log(self.log2,
                                              f"      ✓ Lossless saved "
                                              f"{saved:.0f}%", "ok")
                                except Exception as e:
                                    self._log(self.log2,
                                              f"      Compress error: {e}",
                                              "err")
                                    if not os.path.exists(out_path):
                                        try: os.rename(temp_in, out_path)
                                        except Exception: pass
                            else:
                                self._log(self.log2,
                                          "      ⚠ PyMuPDF missing — "
                                          "keeping merged file as-is.", "ren")
                                try:
                                    import shutil
                                    shutil.copy2(temp_in, out_path)
                                except Exception: pass
                            try: os.remove(temp_in)
                            except Exception: pass
                            after_size = os.path.getsize(out_path)
                            after_mb = after_size / (1024 * 1024)

                    flag = "✅" if after_size <= target_bytes else "⚠ OVER"
                    self._log(self.log2,
                              f"      {flag}  {out_name}: {after_mb:.2f} MB"
                              + ("" if after_size <= target_bytes
                                 else "  (still over target)"),
                              "ok" if after_size <= target_bytes else "ren")
                    outputs.append(out_path)

                total_out = sum(os.path.getsize(o) for o in outputs)
                total_in = sum(os.path.getsize(f) for f in files)
                self._log(self.log2,
                          f"\n✅  Done!  {len(outputs)} output file(s) "
                          f"in {out_dir}\n"
                          f"   Input total: {total_in/1024/1024:.2f} MB  "
                          f"→  Output total: {total_out/1024/1024:.2f} MB",
                          "ok")

                # Final guidance based on result
                over = [o for o in outputs
                        if os.path.getsize(o) > target_bytes]
                if over and not auto_compress:
                    self._log(self.log2,
                              f"\n💡  {len(over)} file(s) exceeded target. "
                              "Enable 'Auto-compress' checkbox and re-run, "
                              "OR use '📦 Compress' mode separately for "
                              "aggressive image-recompression.", "ren")
                elif over and auto_compress:
                    self._log(self.log2,
                              f"\n💡  {len(over)} file(s) still over target "
                              "after lossless compression.\n"
                              "   Use '📦 Compress' mode with 'Allow text "
                              "loss' for further size reduction.", "ren")

                try:
                    if sys.platform.startswith("win"):
                        os.startfile(out_dir)
                except Exception: pass
            except Exception as e:
                self._log(self.log2,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn2.configure(state="normal",
                                        text="🔗   MERGE PDFs")

        # ─── Split runner ────────────────────────────────────
        def _run_split(self):
            if not (HAS_PYMUPDF or HAS_PIKEPDF):
                messagebox.showerror(
                    "Missing dependency",
                    "PyMuPDF or pikepdf required for splitting.\n"
                    "Run: pip install PyMuPDF pikepdf")
                return
            inp = self.split_in_var.get().strip()
            if not inp or not os.path.isfile(inp):
                messagebox.showerror("Error", "Select a valid input PDF.")
                return
            out_dir = self.split_out_var.get().strip()
            if not out_dir:
                out_dir = os.path.join(os.path.dirname(inp), "Split")
            try: os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot create output folder:\n{e}")
                return

            method = self.split_method_var.get()
            kwargs = {"input_path": inp, "output_dir": out_dir}
            try:
                if method == "to_fit":
                    target = float(self.split_target_mb_var.get())
                    if target <= 0: raise ValueError("Target must be > 0")
                    kwargs["target_mb"] = target
                elif method == "by_chunk":
                    n = int(self.split_pages_per_var.get())
                    if n < 1: raise ValueError("Must be >= 1 page per chunk")
                    kwargs["pages_per_chunk"] = n
                elif method == "by_ranges":
                    rstr = self.split_ranges_var.get().strip()
                    if not rstr: raise ValueError("Enter ranges like '1-5, 6-10'")
                    pages, _ = get_pdf_info(inp)
                    kwargs["ranges"] = parse_page_ranges(rstr, pages)
            except Exception as e:
                messagebox.showerror("Invalid input", str(e)); return

            self.run_btn2.configure(state="disabled", text="⏳  SPLITTING…")
            self._clear(self.log2)
            threading.Thread(target=self._split_worker,
                             args=(method, kwargs), daemon=True).start()

        def _split_worker(self, method, kwargs):
            try:
                inp = kwargs["input_path"]
                out_dir = kwargs["output_dir"]
                pages, size = get_pdf_info(inp)
                size_mb = size / (1024 * 1024)
                self._log(self.log2,
                          f"Splitting: {os.path.basename(inp)}\n"
                          f"   {pages} pages, {size_mb:.2f} MB",
                          "hdr")

                outputs = []
                if method == "to_fit":
                    target_mb = kwargs["target_mb"]
                    self._log(self.log2,
                              f"   Method: auto-split to fit under "
                              f"{target_mb} MB per file", "dim")
                    def cb(msg): self._log(self.log2, "   " + msg, "dim")
                    outputs = split_pdf_to_fit(inp, out_dir, target_mb,
                                                progress_cb=cb)
                elif method == "by_chunk":
                    chunk = kwargs["pages_per_chunk"]
                    self._log(self.log2,
                              f"   Method: every {chunk} page(s)", "dim")
                    outputs = split_pdf_by_chunk(inp, out_dir, chunk)
                elif method == "by_ranges":
                    ranges = kwargs["ranges"]
                    self._log(self.log2,
                              f"   Method: custom ranges {ranges}", "dim")
                    outputs = split_pdf_by_ranges(inp, out_dir, ranges)

                self._log(self.log2,
                          f"\n✅  Produced {len(outputs)} file(s):", "ok")
                for fp in outputs:
                    fsize_mb = os.path.getsize(fp) / (1024 * 1024)
                    self._log(self.log2,
                              f"   · {os.path.basename(fp):<60}  "
                              f"{fsize_mb:.2f} MB", "ok")

                # Total size check
                total_out = sum(os.path.getsize(f) for f in outputs)
                self._log(self.log2,
                          f"\n   Total output size: "
                          f"{total_out/1024/1024:.2f} MB  "
                          f"(input was {size_mb:.2f} MB)\n"
                          f"   →  {out_dir}", "ok")

                try:
                    if sys.platform.startswith("win"):
                        os.startfile(out_dir)
                except Exception: pass
            except Exception as e:
                self._log(self.log2,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn2.configure(state="normal",
                                        text="✂   SPLIT PDF")

        def _compress_worker(self, pdfs, out_dir, target_mb, allow_text_loss):
            try:
                mode = "ALLOW TEXT LOSS" if allow_text_loss else "PRESERVE TEXT"
                self._log(self.log2,
                          f"Mode: {mode}  ·  Target: ≤ {target_mb} MB\n"
                          f"Found {len(pdfs)} PDF(s).", "hdr")
                done = hit = 0
                for pdf_path in pdfs:
                    fname = os.path.basename(pdf_path)
                    base, ext = os.path.splitext(fname)
                    out_path = os.path.join(out_dir, fname)
                    src_abs = os.path.abspath(pdf_path)

                    # CRITICAL: never let output path be the same as the source —
                    # Windows can't overwrite an open file, and even on Linux it
                    # would corrupt the input mid-read.
                    if os.path.abspath(out_path) == src_abs:
                        out_path = os.path.join(out_dir, f"{base}_compressed{ext}")
                        self._log(self.log2,
                                  f"   ℹ  Output folder = source folder → saving as "
                                  f"'{base}_compressed{ext}' to avoid overwriting source.",
                                  "dim")

                    # If a different file already exists at that path, add counter
                    if os.path.exists(out_path) and os.path.abspath(out_path) != src_abs:
                        cb, ce = os.path.splitext(os.path.basename(out_path))
                        i = 1
                        while os.path.exists(out_path) and \
                              os.path.abspath(out_path) != src_abs:
                            out_path = os.path.join(out_dir, f"{cb}_{i}{ce}")
                            i += 1

                    # Final safety check
                    if os.path.abspath(out_path) == src_abs:
                        self._log(self.log2,
                                  f"   ✗  Cannot determine safe output path. Skipping.",
                                  "err")
                        continue

                    orig_mb = os.path.getsize(pdf_path) / 1024 / 1024
                    self._log(self.log2,
                              f"\n📄  {fname}  (original: {orig_mb:.2f} MB)", "hdr")
                    self._log(self.log2,
                              f"   →  saving to:  {out_path}", "dim")
                    try:
                        sz, strat, ok = smart_compress(
                            pdf_path, out_path, target_mb, allow_text_loss,
                            log=lambda m, t: self._log(self.log2, m, t))
                        done += 1
                        if ok: hit += 1
                        pct = (1 - sz / max(os.path.getsize(pdf_path), 1)) * 100
                        self._log(self.log2,
                                  f"   📦  Final: {sz/1024/1024:.2f} MB  ({pct:+.1f}%)  "
                                  f"·  {strat}",
                                  "ok" if ok else "ren")
                    except PermissionError as pe:
                        done += 1
                        self._log(self.log2,
                                  f"   ✗  Permission denied: {pe}\n"
                                  f"       This usually means the output file is open in another app.\n"
                                  f"       Close it (e.g. Adobe/Chrome PDF viewer) and re-run.",
                                  "err")
                    except Exception as e:
                        done += 1
                        self._log(self.log2,
                                  f"   ✗  {type(e).__name__}: {e}", "err")
                self._log(self.log2,
                          f"\n✅  Done!  {hit}/{done} hit target  →  {out_dir}",
                          "ok" if hit == done else "ren")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(out_dir)
                except Exception: pass
            except Exception as e:
                self._log(self.log2,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn2.configure(state="normal",
                                        text="🗜   COMPRESS PDF(S)")

        # ════════════════════════════════════════════════════
        #  TAB 3 — TAX COMPARISON CONSOLIDATOR
        # ════════════════════════════════════════════════════
        def _build_tax_comparison(self, tab):
            # Pin log + run button to bottom
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log3 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log3.pack(side="left", fill="both", expand=True)
            sb3 = ttk.Scrollbar(log_frame, command=self.log3.yview)
            sb3.pack(side="right", fill="y")
            self.log3.configure(yscrollcommand=sb3.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log3.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            self.run_btn3 = tk.Button(tab,
                                      text="📊   CONSOLIDATE COMPARISON FILES",
                                      command=self._run_consolidate,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn3.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            # ─── Top: form ─────────────────────────────────────
            info = tk.Frame(tab, bg="#0f2a1a")
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info,
                     text=("  ℹ  Consolidates GSTN portal 'Tax liability & ITC comparison' "
                           "Excel files from multiple states/GSTINs into one workbook.\n"
                           "      Input files: as downloaded from portal, filename pattern "
                           "'<FY>_<GSTIN>_Tax_liability_and_ITC_comparison.xlsx'."),
                     font=("Segoe UI", 9), bg="#0f2a1a", fg=SUCCESS,
                     justify="left", wraplength=700
                     ).pack(padx=10, pady=6, anchor="w")

            self._sec(tab, "📁  Folder containing comparison Excel files")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self.tc_src_var = tk.StringVar()
            self._entry(sf, self.tc_src_var).pack(side="left", fill="x",
                                                  expand=True, ipady=6, padx=(0, 8))
            self._btn(sf, "Browse", self._browse_tc_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file  (blank = auto-named next to source)")
            of3 = tk.Frame(tab, bg=CARD); of3.pack(fill="x", padx=20, pady=(0, 8))
            self.tc_out_var = tk.StringVar()
            self._entry(of3, self.tc_out_var).pack(side="left", fill="x",
                                                   expand=True, ipady=6, padx=(0, 8))
            self._btn(of3, "Save as", self._browse_tc_out).pack(side="left")

            self._sec(tab, "📐  Output Structure")
            mc = tk.Frame(tab, bg=CARD); mc.pack(fill="x", padx=20, pady=(0, 4))
            self.tc_mode_var = tk.StringVar(value="both")
            tk.Radiobutton(mc,
                           text=("  Single sheet (Long Format)  —  saare 8 sections aur saare "
                                 "states ek hi sheet mein stack. Easy filter & pivot."),
                           variable=self.tc_mode_var, value="single",
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9), wraplength=700, justify="left"
                           ).pack(anchor="w", pady=(2, 0))
            tk.Radiobutton(mc,
                           text=("  Multi-sheet (8 sheets)  —  Original GSTN format jaisa: "
                                 "har section ki alag sheet, saare states stacked within."),
                           variable=self.tc_mode_var, value="multi",
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9), wraplength=700, justify="left"
                           ).pack(anchor="w", pady=(4, 0))
            tk.Radiobutton(mc,
                           text=("  Both (Recommended)  —  Multi-sheet + Single sheet + "
                                 "State-Summary pivot in ek hi file."),
                           variable=self.tc_mode_var, value="both",
                           bg=CARD, activebackground=CARD, fg=TEXT,
                           selectcolor=ENTRY_BG, cursor="hand2",
                           font=("Segoe UI", 9), wraplength=700, justify="left"
                           ).pack(anchor="w", pady=(4, 0))

        def _browse_tc_src(self):
            d = filedialog.askdirectory()
            if d: self.tc_src_var.set(d)

        def _browse_tc_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="GST_Tax_Comparison_Consolidated.xlsx")
            if p: self.tc_out_var.set(p)

        def _run_consolidate(self):
            src = self.tc_src_var.get().strip()
            if not src or not os.path.isdir(src):
                messagebox.showerror("Error", "Select a valid input folder.")
                return
            out_path = self.tc_out_var.get().strip()
            if not out_path:
                out_path = os.path.join(os.path.dirname(src) or src,
                                        "GST_Tax_Comparison_Consolidated.xlsx")
            try:
                if not out_path.lower().endswith(".xlsx"):
                    out_path += ".xlsx"
                os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
                # Check writable
                if os.path.exists(out_path):
                    with open(out_path, "ab"): pass
            except PermissionError:
                messagebox.showerror("Output locked",
                                     f"The output file is open in Excel:\n{out_path}\n\n"
                                     "Close it and re-run.")
                return
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot prepare output path:\n{e}")
                return
            mode = self.tc_mode_var.get()
            self.run_btn3.configure(state="disabled", text="⏳  CONSOLIDATING…")
            self._clear(self.log3)
            threading.Thread(target=self._consolidate_worker,
                             args=(src, out_path, mode),
                             daemon=True).start()

        def _consolidate_worker(self, in_folder, out_path, mode):
            try:
                # lazy import — only when this feature is used
                import openpyxl
                from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side
                from openpyxl.utils import get_column_letter

                files = [os.path.join(in_folder, f) for f in os.listdir(in_folder)
                         if f.lower().endswith(".xlsx") and "comparison" in f.lower()]
                if not files:
                    self._log(self.log3,
                              "✗  No 'Tax_liability_and_ITC_comparison' Excel files found "
                              f"in:\n   {in_folder}", "err")
                    return

                mode_label = {"single": "Single sheet (Long Format)",
                              "multi":  "Multi-sheet (8 sheets)",
                              "both":   "Both (Multi-sheet + Long + Pivot)"}[mode]
                self._log(self.log3,
                          f"Mode: {mode_label}\nFound {len(files)} comparison file(s)…",
                          "hdr")

                # Parse all files
                all_data = []  # list of {meta, sheets: {sheet_name: rows}}
                for fp in sorted(files):
                    fn = os.path.basename(fp)
                    try:
                        data = parse_tax_comparison_file(fp)
                        all_data.append(data)
                        self._log(self.log3,
                                  f"   ✓  {data['meta']['state_name']:<22} "
                                  f"({data['meta']['gstin']})  {fn}",
                                  "ok")
                    except Exception as e:
                        self._log(self.log3,
                                  f"   ✗  {fn}  →  {type(e).__name__}: {e}", "err")

                if not all_data:
                    self._log(self.log3, "\n✗  No files could be parsed.", "err")
                    return

                # Write output
                self._log(self.log3, f"\nWriting output: {out_path}", "hdr")
                write_consolidated_comparison(all_data, out_path, mode)

                self._log(self.log3,
                          f"\n✅  Done!  {len(all_data)} state(s)/GSTIN(s) consolidated.\n"
                          f"   →  {out_path}", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log3,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn3.configure(state="normal",
                                        text="📊   CONSOLIDATE COMPARISON FILES")

        # ════════════════════════════════════════════════════
        #  TABS 4 & 5 — GSTR-1 / GSTR-3B BULK EXTRACTORS
        # ════════════════════════════════════════════════════
        def _build_extractor_tab(self, tab, kind):
            """kind = 'gstr1' or 'gstr3b'"""
            assert kind in ("gstr1", "gstr3b")
            is_g1 = (kind == "gstr1")
            title_emoji = "📋" if is_g1 else "🧾"
            title_short = "GSTR-1" if is_g1 else "GSTR-3B"
            button_text = f"▶   EXTRACT {title_short}  →  EXCEL"
            default_out = f"{title_short.replace('-','')}_Consolidated.xlsx"
            desc = ("Reads every GSTR-1 PDF from a folder and produces ONE consolidated "
                    "Excel with every table head (4A, 4B, 5, 6A, 6B, 6C, 7, 8, 9A/9B/9C, "
                    "10, 11A/11B, 12, 13, 14, 14A, 15, 15A) per return."
                    if is_g1 else
                    "Reads every GSTR-3B PDF from a folder and produces ONE consolidated "
                    "Excel with every section — 3.1, 3.1.1, 3.2, 4 (ITC), 5, 5.1, "
                    "6.1 (Tax Payment), and the period-wise breakup.")

            # Pin log + run button to bottom
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            log_box = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                              font=("Consolas", 9),
                              relief="flat", wrap="word", height=4,
                              insertbackground=TEXT,
                              highlightthickness=1, highlightbackground=BORDER,
                              state="disabled")
            log_box.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(log_frame, command=log_box.yview)
            sb.pack(side="right", fill="y")
            log_box.configure(yscrollcommand=sb.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                log_box.tag_config(tag, foreground=color,
                                   font=("Consolas", 9, weight))
            # progress bar above run button
            prog_row = tk.Frame(tab, bg=CARD)
            prog_row.pack(side="bottom", fill="x", padx=20, pady=(0, 4))
            prog_label = tk.Label(prog_row, text="",
                                  font=("Segoe UI", 9), bg=CARD, fg=SUBTEXT)
            prog_label.pack(side="right")
            progress = ttk.Progressbar(prog_row, mode="determinate")
            progress.pack(side="left", fill="x", expand=True, padx=(0, 8))

            run_btn = tk.Button(tab,
                                text=button_text,
                                font=("Segoe UI", 13, "bold"),
                                bg=SUCCESS, fg="white",
                                activebackground="#16a34a",
                                activeforeground="white",
                                relief="flat", cursor="hand2",
                                pady=14)
            run_btn.pack(side="bottom", fill="x", padx=20, pady=(4, 6))

            # Top: form
            if HAS_PDFPLUMBER:
                txt = (f"  ✅  Ready — embedded {title_short} extractor engine. "
                       f"No external binaries needed.")
                bgc = "#0f2a1a"; fgc = SUCCESS
            else:
                txt = ("  ⚠  pdfplumber not installed.  In CMD run:\n"
                       "       pip install pdfplumber openpyxl\n"
                       "  Then restart this app.")
                bgc = "#2d1f0e"; fgc = WARNING
            info = tk.Frame(tab, bg=bgc)
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info, text=txt, font=("Segoe UI", 9),
                     bg=bgc, fg=fgc, justify="left", wraplength=700
                     ).pack(padx=10, pady=6, anchor="w")

            # Short description as a wrapped label (NOT a section header to avoid overflow)
            desc_frame = tk.Frame(tab, bg=CARD)
            desc_frame.pack(side="top", fill="x", padx=20, pady=(8, 0))
            tk.Label(desc_frame,
                     text=f"{title_emoji}  {title_short}  Bulk Extractor",
                     font=("Segoe UI", 11, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w")
            tk.Label(desc_frame, text=desc,
                     font=("Segoe UI", 9), bg=CARD, fg=SUBTEXT,
                     wraplength=700, justify="left").pack(anchor="w", pady=(2, 0))

            self._sec(tab, "📁  Input folder  (contains all PDFs across states/months)")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            src_var = tk.StringVar()
            self._entry(sf, src_var).pack(side="left", fill="x",
                                          expand=True, ipady=6, padx=(0, 8))
            def _browse_src():
                d = filedialog.askdirectory()
                if d: src_var.set(d)
            self._btn(sf, "Browse folder", _browse_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file")
            of = tk.Frame(tab, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            out_var = tk.StringVar()
            self._entry(of, out_var).pack(side="left", fill="x",
                                          expand=True, ipady=6, padx=(0, 8))
            def _browse_out():
                p = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx")],
                    initialfile=default_out)
                if p: out_var.set(p)
            self._btn(of, "Save as", _browse_out).pack(side="left")

            # Store widgets on self with kind suffix
            if is_g1:
                self.g1_src_var = src_var
                self.g1_out_var = out_var
                self.g1_log = log_box
                self.g1_progress = progress
                self.g1_progress_label = prog_label
                self.run_btn_g1 = run_btn
                run_btn.configure(command=lambda: self._run_extractor("gstr1"))
            else:
                self.g3_src_var = src_var
                self.g3_out_var = out_var
                self.g3_log = log_box
                self.g3_progress = progress
                self.g3_progress_label = prog_label
                self.run_btn_g3 = run_btn
                run_btn.configure(command=lambda: self._run_extractor("gstr3b"))

        def _run_extractor(self, kind):
            is_g1 = (kind == "gstr1")
            src_var  = self.g1_src_var  if is_g1 else self.g3_src_var
            out_var  = self.g1_out_var  if is_g1 else self.g3_out_var
            log_box  = self.g1_log      if is_g1 else self.g3_log
            run_btn  = self.run_btn_g1  if is_g1 else self.run_btn_g3
            title_short = "GSTR-1" if is_g1 else "GSTR-3B"

            in_text  = src_var.get().strip()
            out_text = out_var.get().strip()
            if not in_text:
                messagebox.showerror("Input missing",
                                     f"Please select the folder containing {title_short} PDFs.")
                return
            if not os.path.isdir(in_text):
                messagebox.showerror("Folder not found",
                                     f"This folder does not exist:\n{in_text}")
                return
            if not out_text:
                out_text = os.path.join(os.path.dirname(in_text) or in_text,
                                        f"{title_short.replace('-','')}_Consolidated.xlsx")
                out_var.set(out_text)

            # Lazy-load the embedded engine
            try:
                _load_extractor_engines()
            except Exception as e:
                messagebox.showerror("Engine load failed",
                                     f"Could not load the {title_short} extractor engine:\n\n{e}")
                return
            ns = _GSTR1_NS if is_g1 else _GSTR3B_NS

            try:
                resolve_output_path = ns["resolve_output_path"]
                out_path = resolve_output_path(out_text)
            except PermissionError as e:
                messagebox.showerror("Output file is locked", str(e))
                return
            except Exception as e:
                messagebox.showerror("Output path error",
                                     f"Could not prepare output path:\n{e}")
                return
            out_var.set(str(out_path))

            self._clear(log_box)
            run_btn.configure(state="disabled", text=f"⏳  EXTRACTING…")
            threading.Thread(target=self._extractor_worker,
                             args=(kind, in_text, out_path),
                             daemon=True).start()

        def _extractor_worker(self, kind, in_folder, out_path):
            is_g1 = (kind == "gstr1")
            log_box   = self.g1_log if is_g1 else self.g3_log
            run_btn   = self.run_btn_g1 if is_g1 else self.run_btn_g3
            progress  = self.g1_progress if is_g1 else self.g3_progress
            prog_lbl  = self.g1_progress_label if is_g1 else self.g3_progress_label
            ns        = _GSTR1_NS if is_g1 else _GSTR3B_NS
            title_short = "GSTR-1" if is_g1 else "GSTR-3B"
            button_text = f"▶   EXTRACT {title_short}  →  EXCEL"
            from pathlib import Path as _Path

            def _log(msg, tag="info"):
                tag_map = {"ok": "ok", "fail": "err", "info": "hdr"}
                t = tag_map.get(tag, "hdr")
                log_box.configure(state="normal")
                log_box.insert("end", msg + "\n", t)
                log_box.see("end")
                log_box.configure(state="disabled")

            def _progress(i, total):
                progress["maximum"] = total
                progress["value"] = i
                pct = (i / total * 100) if total else 0
                prog_lbl.configure(text=f"{i} / {total}  ({pct:0.1f}%)")

            try:
                process_pdfs = ns["process_pdfs"]
                _log(f"Input folder:   {in_folder}", "info")
                _log(f"Output file:    {out_path}", "info")
                _log("─" * 60, "info")

                while True:
                    try:
                        ok, fail, total, out_path = process_pdfs(
                            _Path(in_folder), out_path,
                            on_progress=_progress, on_log=_log)
                        break
                    except PermissionError as pe:
                        _log(f"\n⚠  Cannot write: {pe}", "fail")
                        ans = messagebox.askretrycancel(
                            "Output locked",
                            f"Output file is locked (probably open in Excel):\n"
                            f"{out_path}\n\nClose it and click Retry.")
                        if not ans:
                            _log("Aborted by user.", "fail")
                            return

                _log("─" * 60, "info")
                _log(f"\n✅  Done!  {ok} OK, {fail} failed, total {total}.", "ok")
                _log(f"     →  {out_path}", "info")
                if fail > 0:
                    messagebox.showwarning(
                        f"{title_short} Extractor — Completed with failures",
                        f"{ok} succeeded, {fail} failed.\n\n"
                        "See the 'Processing Log' sheet for details.\n\n"
                        f"Saved to:\n{out_path}")
                else:
                    messagebox.showinfo(
                        f"{title_short} Extractor — Done",
                        f"All {total} PDFs processed successfully.\n\nSaved to:\n{out_path}")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except FileNotFoundError as e:
                _log(f"\n✗  {e}", "fail")
                messagebox.showerror(f"{title_short} Extractor", str(e))
            except Exception as e:
                _log(f"\n✗  {type(e).__name__}: {e}\n{traceback.format_exc()}", "fail")
                messagebox.showerror(f"{title_short} Extractor — Error",
                                     f"{type(e).__name__}: {e}")
            finally:
                run_btn.configure(state="normal", text=button_text)

        # ════════════════════════════════════════════════════
        #  TAB 6 — GSTR-2B CONSOLIDATOR
        # ════════════════════════════════════════════════════
        def _build_gstr2b_tab(self, tab):
            # Pin log + run button to bottom
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log6 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log6.pack(side="left", fill="both", expand=True)
            sb6 = ttk.Scrollbar(log_frame, command=self.log6.yview)
            sb6.pack(side="right", fill="y")
            self.log6.configure(yscrollcommand=sb6.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log6.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            self.run_btn6 = tk.Button(tab,
                                      text="📑   CONSOLIDATE GSTR-2B FILES",
                                      command=self._run_gstr2b,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn6.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            # Top: info banner
            info = tk.Frame(tab, bg="#0f2a1a")
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info,
                     text=("  ℹ  Consolidates GSTR-2B Excel files from multiple "
                           "states/months into ONE workbook.\n"
                           "      • Summary cards (ITC Available / Not Available / "
                           "Reversal / Rejected) — stacked across states\n"
                           "      • Console sheet — every invoice/note/document in unified "
                           "invoice-level columns (B2B/CDNR/ECO/ISD/IMPG)\n"
                           "      • Per-sheet outputs — all 24 GSTN sheets preserved in "
                           "original format, all states stacked\n"
                           "      • Credit Note rows auto-flipped to NEGATIVE "
                           "(highlighted in orange) for correct netting"),
                     font=("Segoe UI", 9), bg="#0f2a1a", fg=SUCCESS,
                     justify="left", wraplength=720
                     ).pack(padx=10, pady=8, anchor="w")

            self._sec(tab, "📁  Folder containing GSTR-2B Excel files")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self.g2b_src_var = tk.StringVar()
            self._entry(sf, self.g2b_src_var).pack(side="left", fill="x",
                                                   expand=True, ipady=6, padx=(0, 8))
            self._btn(sf, "Browse", self._browse_g2b_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file  (blank = auto-named next to source)")
            of = tk.Frame(tab, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self.g2b_out_var = tk.StringVar()
            self._entry(of, self.g2b_out_var).pack(side="left", fill="x",
                                                   expand=True, ipady=6, padx=(0, 8))
            self._btn(of, "Save as", self._browse_g2b_out).pack(side="left")

        def _browse_g2b_src(self):
            d = filedialog.askdirectory()
            if d: self.g2b_src_var.set(d)

        def _browse_g2b_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="GSTR2B_Consolidated.xlsx")
            if p: self.g2b_out_var.set(p)

        def _run_gstr2b(self):
            src = self.g2b_src_var.get().strip()
            if not src or not os.path.isdir(src):
                messagebox.showerror("Error", "Select a valid input folder.")
                return
            out_path = self.g2b_out_var.get().strip()
            if not out_path:
                out_path = os.path.join(os.path.dirname(src) or src,
                                        "GSTR2B_Consolidated.xlsx")
            try:
                if not out_path.lower().endswith(".xlsx"):
                    out_path += ".xlsx"
                os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
                if os.path.exists(out_path):
                    with open(out_path, "ab"): pass
            except PermissionError:
                messagebox.showerror("Output locked",
                                     f"Output file is open in Excel:\n{out_path}\n\n"
                                     "Close it and re-run.")
                return
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot prepare output path:\n{e}")
                return
            self.run_btn6.configure(state="disabled", text="⏳  CONSOLIDATING…")
            self._clear(self.log6)
            threading.Thread(target=self._gstr2b_worker,
                             args=(src, out_path),
                             daemon=True).start()

        def _gstr2b_worker(self, in_folder, out_path):
            try:
                # Find all candidate GSTR-2B files
                files = [os.path.join(in_folder, f) for f in os.listdir(in_folder)
                         if f.lower().endswith(".xlsx") and "gstr2b" in f.lower()]
                if not files:
                    # Fallback: any xlsx in folder
                    files = [os.path.join(in_folder, f) for f in os.listdir(in_folder)
                             if f.lower().endswith(".xlsx")]
                if not files:
                    self._log(self.log6,
                              "✗  No Excel files found in this folder.", "err")
                    return

                self._log(self.log6,
                          f"Found {len(files)} Excel file(s). Parsing…", "hdr")

                all_data = []
                missing_meta_files = []  # Files where meta couldn't be extracted
                for fp in sorted(files):
                    fn = os.path.basename(fp)
                    try:
                        data = parse_gstr2b_file(fp)
                        m = data["meta"]
                        # Detect missing meta — would cause blank Month/State rows
                        if not m.get("gstin") or not m.get("month_abbr") \
                                or not m.get("state_code"):
                            missing_meta_files.append(
                                (fn, m.get("gstin", "?"),
                                 m.get("month_abbr", "?"),
                                 m.get("state_code", "?")))
                            self._log(self.log6,
                                      f"   ⚠  {fn}  →  Meta incomplete: "
                                      f"GSTIN={m.get('gstin') or 'MISSING'}, "
                                      f"Month={m.get('month_abbr') or 'MISSING'}, "
                                      f"State={m.get('state_code') or 'MISSING'} "
                                      "(check Read Me sheet + filename pattern)",
                                      "ren")
                        all_data.append(data)
                        n_txn = sum(len(t["rows"]) for t in data["transactions"].values())
                        n_summ = sum(len(s) for s in data["summary"].values())
                        # Only log success line if meta is complete (avoid double-noise)
                        if m.get("gstin") and m.get("month_abbr"):
                            self._log(self.log6,
                                      f"   ✓  {m['state_name']:<20} "
                                      f"{m['month_abbr']:<8}  "
                                      f"({m['gstin']})  →  "
                                      f"{len(data['transactions'])} txn sheets, "
                                      f"{n_txn} txn rows, {n_summ} summary rows",
                                      "ok")
                    except Exception as e:
                        self._log(self.log6,
                                  f"   ✗  {fn}  →  {type(e).__name__}: {e}", "err")

                if not all_data:
                    self._log(self.log6, "\n✗  No files could be parsed.", "err")
                    return

                # Surface missing-meta summary BEFORE writing
                if missing_meta_files:
                    self._log(self.log6,
                              f"\n⚠  {len(missing_meta_files)} file(s) had "
                              "incomplete meta — those rows will show blank "
                              "Month/State/GSTIN in the output.", "ren")
                    self._log(self.log6,
                              "   Common cause: Read Me sheet missing OR "
                              "filename doesn't match GSTN export pattern. "
                              "Re-download from portal to fix.", "ren")

                # ─── Merge split files (big-state 2-part downloads) ──
                files_before = len(all_data)
                grp = _g2b_group_split_files(all_data)
                all_data = grp["merged_list"]
                split_groups = grp["split_groups"]
                redownloads = grp["redownloads"]
                no_meta_files = grp["no_meta_files"]

                if split_groups:
                    self._log(self.log6,
                              f"\n📎  Multi-part download(s) detected — "
                              f"{len(split_groups)} group(s) merged:",
                              "hdr")
                    for g in split_groups:
                        self._log(self.log6,
                                  f"   • {g['state']:<18} {g['period']:<10} "
                                  f"({g['gstin']})  →  {g['parts']} parts "
                                  f"merged,  {g['total_rows']} total txn rows",
                                  "ok")
                        # Also list the actual file names for audit
                        for sf in g["sources"]:
                            self._log(self.log6, f"          {sf}", "dim")

                if redownloads:
                    self._log(self.log6,
                              f"\n⚠  Re-download detected — same GSTIN+Period "
                              "had multiple generation dates. Kept LATEST only:",
                              "ren")
                    for rd in redownloads:
                        self._log(self.log6,
                                  f"   • {rd['state']:<18} {rd['period']:<10} "
                                  f"({rd['gstin']})  →  kept gen-date "
                                  f"{rd['kept_gen_date']}, discarded "
                                  f"{len(rd['discarded_files'])} older file(s)",
                                  "ren")
                        for sf in rd["discarded_files"]:
                            self._log(self.log6, f"          (skipped) {sf}",
                                      "dim")

                if no_meta_files:
                    self._log(self.log6,
                              f"\n⚠  {len(no_meta_files)} file(s) had no "
                              "parseable (GSTIN, Period) — kept as standalone "
                              "entries (NOT merged with anything):", "ren")
                    for sf in no_meta_files:
                        self._log(self.log6, f"   • {sf}", "ren")

                # Reassuring accounting summary
                merge_absorbed = sum(g["parts"] - 1 for g in split_groups)
                rd_discarded = sum(len(rd["discarded_files"])
                                   for rd in redownloads)
                if split_groups or redownloads:
                    self._log(self.log6,
                              f"\n📊  Accounting:  {files_before} input file(s) → "
                              f"{len(all_data)} unique entries\n"
                              f"     ({merge_absorbed} absorbed via multi-part "
                              f"merge, {rd_discarded} discarded as older "
                              f"re-downloads, {len(no_meta_files)} kept "
                              "as-is despite missing meta)", "ok")

                # ─── Period coverage summary ────────────────────
                def _month_sort_key(mo):
                    if not mo or "-" not in mo: return "9999-99"
                    abbr, yy = mo.split("-", 1)
                    mn = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04",
                          "May":"05","Jun":"06","Jul":"07","Aug":"08",
                          "Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
                    full_year = ("20" + yy.strip()) if len(yy.strip()) == 2 else yy.strip()
                    return full_year + mn.get(abbr, "00")

                month_count = {}
                gstin_set = set()
                for d in all_data:
                    mo = d["meta"].get("month_abbr", "")
                    g = d["meta"].get("gstin", "")
                    if mo and g:
                        month_count[mo] = month_count.get(mo, 0) + 1
                        gstin_set.add(g)
                if month_count:
                    avg_states_per_month = sum(month_count.values()) / len(month_count)
                    self._log(self.log6,
                              f"\n📅  Period coverage  ({len(gstin_set)} unique "
                              f"GSTINs across {len(month_count)} months):",
                              "hdr")
                    for mo in sorted(month_count.keys(), key=_month_sort_key):
                        flag = ""
                        if month_count[mo] < avg_states_per_month * 0.7:
                            flag = "  ⚠  noticeably fewer than other months"
                        self._log(self.log6,
                                  f"     {mo:<10}  {month_count[mo]:>3} GSTIN(s)"
                                  + flag,
                                  "ren" if flag else "ok")
                    self._log(self.log6,
                              f"\n   Check the 'Coverage' sheet in the output "
                              "Excel for full State × Month grid.", "dim")

                self._log(self.log6,
                          f"\nWriting consolidated output: {out_path}", "hdr")
                flip_count = write_consolidated_gstr2b(all_data, out_path)

                done_msg = (f"\n✅  Done!  {len(all_data)} state-month entries "
                            "consolidated")
                if split_groups:
                    done_msg += (f"  (from {files_before} input files, "
                                 f"{len(split_groups)} multi-part merged"
                                 + (f", {rd_discarded} re-download(s) discarded"
                                    if rd_discarded else "")
                                 + ")")
                done_msg += (f".\n   →  {flip_count} Credit Note rows "
                             f"flipped to negative.\n   →  {out_path}")
                self._log(self.log6, done_msg, "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log6,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn6.configure(state="normal",
                                        text="📑   CONSOLIDATE GSTR-2B FILES")

        # ════════════════════════════════════════════════════
        #  TAB 8 — GSTR-9 / 9C CONSOLIDATOR (PDF-based)
        # ════════════════════════════════════════════════════
        def _build_gstr9_9c_tab(self, tab):
            # Log + run button pinned to bottom
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log8 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log8.pack(side="left", fill="both", expand=True)
            sb8 = ttk.Scrollbar(log_frame, command=self.log8.yview)
            sb8.pack(side="right", fill="y")
            self.log8.configure(yscrollcommand=sb8.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log8.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            self.run_btn8 = tk.Button(tab,
                                      text="📔   CONSOLIDATE GSTR-9 / 9C PDFs",
                                      command=self._run_gstr9_9c,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn8.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            # Info banner
            info = tk.Frame(tab, bg="#0f2a1a")
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info,
                     text=("  ℹ  Extracts data from GSTR-9 (Annual Return) and "
                           "GSTR-9C (Reconciliation) PDFs.\n"
                           "      • Auto-detects form type from PDF content\n"
                           "      • Picks up GSTIN / FY / Legal Name / ARN / "
                           "Filing Date from header\n"
                           "      • Output: ONE Excel with separate GSTR-9 + "
                           "GSTR-9C Console sheets (long-format, all rows)\n"
                           "      • Filterable by State / Pt / Table / Sr.No — "
                           "ideal for multi-state / multi-FY comparison"),
                     font=("Segoe UI", 9), bg="#0f2a1a", fg=SUCCESS,
                     justify="left", wraplength=720
                     ).pack(padx=10, pady=8, anchor="w")

            self._sec(tab, "📁  Folder containing GSTR-9 and/or GSTR-9C PDF files")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self.g9_src_var = tk.StringVar()
            self._entry(sf, self.g9_src_var).pack(side="left", fill="x",
                                                   expand=True, ipady=6, padx=(0, 8))
            self._btn(sf, "Browse", self._browse_g9_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file  (blank = auto-named)")
            of = tk.Frame(tab, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self.g9_out_var = tk.StringVar()
            self._entry(of, self.g9_out_var).pack(side="left", fill="x",
                                                   expand=True, ipady=6, padx=(0, 8))
            self._btn(of, "Save as", self._browse_g9_out).pack(side="left")

        def _browse_g9_src(self):
            d = filedialog.askdirectory()
            if d: self.g9_src_var.set(d)

        def _browse_g9_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="GSTR9_9C_Consolidated.xlsx")
            if p: self.g9_out_var.set(p)

        def _run_gstr9_9c(self):
            src = self.g9_src_var.get().strip()
            if not src or not os.path.isdir(src):
                messagebox.showerror("Error", "Select a valid input folder.")
                return
            out_path = self.g9_out_var.get().strip()
            if not out_path:
                out_path = os.path.join(os.path.dirname(src) or src,
                                        "GSTR9_9C_Consolidated.xlsx")
            try:
                if not out_path.lower().endswith(".xlsx"):
                    out_path += ".xlsx"
                os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
                if os.path.exists(out_path):
                    with open(out_path, "ab"): pass
            except PermissionError:
                messagebox.showerror("Output locked",
                                     f"Output file is open in Excel:\n{out_path}\n\n"
                                     "Close it and re-run.")
                return
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot prepare output path:\n{e}")
                return
            self.run_btn8.configure(state="disabled", text="⏳  EXTRACTING…")
            self._clear(self.log8)
            threading.Thread(target=self._gstr9_worker,
                             args=(src, out_path),
                             daemon=True).start()

        def _gstr9_worker(self, in_folder, out_path):
            try:
                # Find PDF files
                files = [os.path.join(in_folder, f) for f in os.listdir(in_folder)
                         if f.lower().endswith(".pdf")]
                if not files:
                    self._log(self.log8,
                              "✗  No PDF files found in this folder.", "err")
                    return

                self._log(self.log8,
                          f"Found {len(files)} PDF file(s). Parsing…", "hdr")

                all_data = []
                for fp in sorted(files):
                    fn = os.path.basename(fp)
                    try:
                        data = parse_gstr9_or_9c_pdf(fp)
                        m = data["meta"]
                        all_data.append(data)
                        self._log(self.log8,
                                  f"   ✓  {m['form_type']:<8}  {m['state_name']:<20} "
                                  f"FY {m['fy']:<8}  ({m['gstin']})  →  "
                                  f"{len(data['rows'])} rows extracted",
                                  "ok")
                    except Exception as e:
                        self._log(self.log8,
                                  f"   ✗  {fn}  →  {type(e).__name__}: {e}",
                                  "err")

                if not all_data:
                    self._log(self.log8,
                              "\n✗  No PDFs could be parsed.", "err")
                    return

                self._log(self.log8,
                          f"\nWriting consolidated output: {out_path}", "hdr")
                n_9, n_9c = write_consolidated_gstr9_9c(all_data, out_path)

                self._log(self.log8,
                          f"\n✅  Done!  {n_9} GSTR-9 + {n_9c} GSTR-9C consolidated.\n"
                          f"   →  {out_path}", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log8,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn8.configure(state="normal",
                                        text="📔   CONSOLIDATE GSTR-9 / 9C PDFs")

        # ════════════════════════════════════════════════════
        #  TAB 9 — ECRRS (Reclaim Statement) CONSOLIDATOR
        # ════════════════════════════════════════════════════
        def _build_ecrrs_tab(self, tab):
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log9 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log9.pack(side="left", fill="both", expand=True)
            sb9 = ttk.Scrollbar(log_frame, command=self.log9.yview)
            sb9.pack(side="right", fill="y")
            self.log9.configure(yscrollcommand=sb9.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log9.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            self.run_btn9 = tk.Button(tab,
                                      text="🔄   CONSOLIDATE RECLAIM STATEMENTS",
                                      command=self._run_ecrrs,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn9.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            info = tk.Frame(tab, bg="#0f2a1a")
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info,
                     text=("  ℹ  Consolidates GSTN portal CSV exports of "
                           "'Electronic Credit Reversal & Re-claimed Statement'.\n"
                           "      • 🔍 AUTO-DEDUP: Same GSTIN's overlapping-period "
                           "downloads merged automatically (duplicate rows by "
                           "Reference No removed)\n"
                           "      • Output: Cover + Console (long format) + "
                           "Summary (per-GSTIN totals with health flags) + "
                           "Per-GSTIN sheets\n"
                           "      • Health: 🟢 Mostly reclaimed · 🟡 Partial reclaim "
                           "· 🟠 Heavy parked balance · 🔴 Negative balance\n"
                           "      • Opening/Closing rows highlighted, negatives in "
                           "red — multi-state comparison friendly"),
                     font=("Segoe UI", 9), bg="#0f2a1a", fg=SUCCESS,
                     justify="left", wraplength=720
                     ).pack(padx=10, pady=8, anchor="w")

            self._sec(tab, "📁  Folder containing ECRRS CSV files "
                           "(from GSTN portal)")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self.ecr_src_var = tk.StringVar()
            self._entry(sf, self.ecr_src_var).pack(side="left", fill="x",
                                                    expand=True, ipady=6,
                                                    padx=(0, 8))
            self._btn(sf, "Browse", self._browse_ecr_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file  (blank = auto-named)")
            of = tk.Frame(tab, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self.ecr_out_var = tk.StringVar()
            self._entry(of, self.ecr_out_var).pack(side="left", fill="x",
                                                    expand=True, ipady=6,
                                                    padx=(0, 8))
            self._btn(of, "Save as", self._browse_ecr_out).pack(side="left")

        def _browse_ecr_src(self):
            d = filedialog.askdirectory()
            if d: self.ecr_src_var.set(d)

        def _browse_ecr_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="ECRRS_Reclaim_Consolidated.xlsx")
            if p: self.ecr_out_var.set(p)

        def _run_ecrrs(self):
            src = self.ecr_src_var.get().strip()
            if not src or not os.path.isdir(src):
                messagebox.showerror("Error", "Select a valid input folder.")
                return
            out_path = self.ecr_out_var.get().strip()
            if not out_path:
                out_path = os.path.join(os.path.dirname(src) or src,
                                        "ECRRS_Reclaim_Consolidated.xlsx")
            try:
                if not out_path.lower().endswith(".xlsx"):
                    out_path += ".xlsx"
                os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
                if os.path.exists(out_path):
                    with open(out_path, "ab"): pass
            except PermissionError:
                messagebox.showerror("Output locked",
                                     f"Output file is open in Excel:\n{out_path}\n\n"
                                     "Close it and re-run.")
                return
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot prepare output path:\n{e}")
                return
            self.run_btn9.configure(state="disabled", text="⏳  CONSOLIDATING…")
            self._clear(self.log9)
            threading.Thread(target=self._ecrrs_worker,
                             args=(src, out_path),
                             daemon=True).start()

        def _ecrrs_worker(self, in_folder, out_path):
            try:
                files = [os.path.join(in_folder, f) for f in os.listdir(in_folder)
                         if f.lower().endswith(".csv")]
                if not files:
                    self._log(self.log9,
                              "✗  No CSV files found in this folder.", "err")
                    return

                self._log(self.log9,
                          f"Found {len(files)} CSV file(s). Parsing…", "hdr")

                all_data = []
                for fp in sorted(files):
                    fn = os.path.basename(fp)
                    try:
                        data = parse_ecrrs_csv(fp)
                        m = data["meta"]
                        if not m["gstin"]:
                            self._log(self.log9,
                                      f"   ✗  {fn} — not a valid ECRRS CSV "
                                      "(no GSTIN found, skipping)", "err")
                            continue
                        all_data.append(data)
                        monthly = sum(1 for r in data["rows"]
                                      if r["row_type"] == "Monthly")
                        self._log(self.log9,
                                  f"   ✓  {m['state_name']:<20} "
                                  f"({m['gstin']})  {m['from_date']} → "
                                  f"{m['to_date']}  →  {monthly} monthly rows",
                                  "ok")
                    except Exception as e:
                        self._log(self.log9,
                                  f"   ✗  {fn}  →  {type(e).__name__}: {e}",
                                  "err")

                if not all_data:
                    self._log(self.log9,
                              "\n✗  No valid CSVs parsed.", "err")
                    return

                self._log(self.log9,
                          f"\nWriting consolidated output: {out_path}", "hdr")
                n_gstins, dedup_stats = write_consolidated_ecrrs(
                    all_data, out_path)

                # Log dedup info if duplicates were removed or files merged
                if dedup_stats:
                    self._log(self.log9,
                              "\n🔍  De-duplication summary:", "hdr")
                    total_dups = sum(s["duplicates_skipped"]
                                     for s in dedup_stats)
                    for s in dedup_stats:
                        if s["files_merged"] > 1:
                            self._log(self.log9,
                                      f"   • {s['gstin']}: merged "
                                      f"{s['files_merged']} overlapping files "
                                      f"→ {s['unique_monthly']} unique monthly rows",
                                      "ok")
                        if s["duplicates_skipped"] > 0:
                            ex = (", ".join(s["duplicate_examples"])
                                  if s["duplicate_examples"] else "")
                            self._log(self.log9,
                                      f"     ↳ removed {s['duplicates_skipped']} "
                                      f"duplicate row(s) by Reference No"
                                      + (f" — e.g. {ex}" if ex else ""),
                                      "ren")
                    if total_dups > 0:
                        self._log(self.log9,
                                  f"\n   Total duplicate rows removed: "
                                  f"{total_dups}", "hdr")

                self._log(self.log9,
                          f"\n✅  Done!  {n_gstins} unique GSTIN(s) consolidated "
                          f"from {len(all_data)} CSV file(s).\n"
                          f"   →  Sheets: Cover + Console + Summary + "
                          f"{n_gstins} detail sheets\n"
                          f"   →  {out_path}", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log9,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn9.configure(state="normal",
                                        text="🔄   CONSOLIDATE RECLAIM STATEMENTS")

        # ════════════════════════════════════════════════════
        #  TAB 10 — Electronic Credit Ledger consolidator
        # ════════════════════════════════════════════════════
        def _build_ecl_tab(self, tab):
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log10 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                  font=("Consolas", 9),
                                  relief="flat", wrap="word", height=4,
                                  insertbackground=TEXT,
                                  highlightthickness=1, highlightbackground=BORDER,
                                  state="disabled")
            self.log10.pack(side="left", fill="both", expand=True)
            sb10 = ttk.Scrollbar(log_frame, command=self.log10.yview)
            sb10.pack(side="right", fill="y")
            self.log10.configure(yscrollcommand=sb10.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log10.tag_config(tag, foreground=color,
                                       font=("Consolas", 9, weight))

            self.run_btn10 = tk.Button(tab,
                                        text="💳   CONSOLIDATE CREDIT LEDGERS",
                                        command=self._run_ecl,
                                        font=("Segoe UI", 13, "bold"),
                                        bg=SUCCESS, fg="white",
                                        activebackground="#16a34a",
                                        activeforeground="white",
                                        relief="flat", cursor="hand2",
                                        pady=14)
            self.run_btn10.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            info = tk.Frame(tab, bg="#0f2a1a")
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info,
                     text=("  ℹ  Consolidates GSTN portal CSV exports of "
                           "'Electronic Credit Ledger' (ITC ledger).\n"
                           "      • Multi-GSTIN / multi-state support — drop "
                           "all CSVs in one folder\n"
                           "      • Output: Cover + Console (long format) + "
                           "Summary (per-GSTIN totals) + per-GSTIN detail sheets\n"
                           "      • Amount + Balance split per row: "
                           "IGST / CGST / SGST / CESS / Total\n"
                           "      • Opening / Closing rows highlighted, "
                           "negative values in red"),
                     font=("Segoe UI", 9), bg="#0f2a1a", fg=SUCCESS,
                     justify="left", wraplength=720
                     ).pack(padx=10, pady=8, anchor="w")

            self._sec(tab, "📁  Folder containing Electronic Credit Ledger "
                           "CSV files (from GSTN portal)")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self.ecl_src_var = tk.StringVar()
            self._entry(sf, self.ecl_src_var).pack(side="left", fill="x",
                                                    expand=True, ipady=6,
                                                    padx=(0, 8))
            self._btn(sf, "Browse", self._browse_ecl_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file  (blank = auto-named)")
            of = tk.Frame(tab, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self.ecl_out_var = tk.StringVar()
            self._entry(of, self.ecl_out_var).pack(side="left", fill="x",
                                                    expand=True, ipady=6,
                                                    padx=(0, 8))
            self._btn(of, "Save as", self._browse_ecl_out).pack(side="left")

        def _browse_ecl_src(self):
            d = filedialog.askdirectory()
            if d: self.ecl_src_var.set(d)

        def _browse_ecl_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Credit_Ledger_Consolidated.xlsx")
            if p: self.ecl_out_var.set(p)

        def _run_ecl(self):
            src = self.ecl_src_var.get().strip()
            if not src or not os.path.isdir(src):
                messagebox.showerror("Error", "Select a valid input folder.")
                return
            out_path = self.ecl_out_var.get().strip()
            if not out_path:
                out_path = os.path.join(os.path.dirname(src) or src,
                                        "Credit_Ledger_Consolidated.xlsx")
            try:
                if not out_path.lower().endswith(".xlsx"):
                    out_path += ".xlsx"
                os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
                if os.path.exists(out_path):
                    with open(out_path, "ab"): pass
            except PermissionError:
                messagebox.showerror("Output locked",
                                     f"Output file is open in Excel:\n{out_path}\n\n"
                                     "Close it and re-run.")
                return
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot prepare output path:\n{e}")
                return
            self.run_btn10.configure(state="disabled",
                                      text="⏳  CONSOLIDATING…")
            self._clear(self.log10)
            threading.Thread(target=self._ecl_worker,
                             args=(src, out_path), daemon=True).start()

        def _ecl_worker(self, in_folder, out_path):
            try:
                files = [os.path.join(in_folder, f)
                         for f in os.listdir(in_folder)
                         if f.lower().endswith(".csv")]
                if not files:
                    self._log(self.log10,
                              "✗  No CSV files found in this folder.", "err")
                    return
                self._log(self.log10,
                          f"Found {len(files)} CSV file(s). Parsing…", "hdr")

                all_data = []
                for fp in sorted(files):
                    fn = os.path.basename(fp)
                    try:
                        data = parse_ecl_csv(fp)
                        m = data["meta"]
                        if not m["gstin"]:
                            self._log(self.log10,
                                      f"   ✗  {fn} — no GSTIN found, "
                                      "skipping (not an ECL CSV?)", "err")
                            continue
                        all_data.append(data)
                        credit_n = sum(1 for r in data["rows"]
                                       if r["row_type"] == "Credit")
                        debit_n = sum(1 for r in data["rows"]
                                      if r["row_type"] == "Debit")
                        self._log(self.log10,
                                  f"   ✓  {m['state_name']:<20} "
                                  f"({m['gstin']})  {m['from_date']} → "
                                  f"{m['to_date']}  →  {credit_n} Cr, "
                                  f"{debit_n} Dr rows",
                                  "ok")
                    except Exception as e:
                        self._log(self.log10,
                                  f"   ✗  {fn}  →  {type(e).__name__}: {e}",
                                  "err")

                if not all_data:
                    self._log(self.log10,
                              "\n✗  No valid CSVs parsed.", "err")
                    return

                # Dedup overlapping period downloads (same GSTIN, different
                # date ranges) — e.g. Apr-Feb + Feb-Mar would double-count Feb
                files_before = len(all_data)
                all_data, dedup_stats = dedup_ledger_files(all_data)
                merge_groups = [s for s in dedup_stats if s['files_merged'] > 1]
                if merge_groups:
                    self._log(self.log10,
                              f"\n🔍  Detected {len(merge_groups)} GSTIN(s) "
                              "with multiple overlapping CSV files — merged "
                              "into single entries:", "hdr")
                    total_dups = 0
                    for s in merge_groups:
                        self._log(self.log10,
                                  f"   • {s['state_name']:<18} "
                                  f"({s['gstin']})  →  {s['files_merged']} "
                                  f"files merged, {s['duplicates_skipped']} "
                                  f"duplicate txn(s) skipped, "
                                  f"{s['unique_txns']} unique txn(s)", "ok")
                        total_dups += s['duplicates_skipped']
                    self._log(self.log10,
                              f"\n   {files_before} input file(s)  →  "
                              f"{len(all_data)} unique GSTIN(s)   "
                              f"({total_dups} duplicate txn(s) skipped)",
                              "dim")

                self._log(self.log10,
                          f"\nWriting consolidated output: {out_path}", "hdr")
                n = write_consolidated_ecl(all_data, out_path)
                self._log(self.log10,
                          f"\n✅  Done!  {n} GSTIN(s) consolidated.\n"
                          f"   →  {out_path}", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log10,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn10.configure(state="normal",
                                          text="💳   CONSOLIDATE CREDIT LEDGERS")

        # ════════════════════════════════════════════════════
        #  TAB 11 — Electronic Cash Ledger consolidator
        # ════════════════════════════════════════════════════
        def _build_ecashl_tab(self, tab):
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 12))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log11 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                  font=("Consolas", 9),
                                  relief="flat", wrap="word", height=4,
                                  insertbackground=TEXT,
                                  highlightthickness=1, highlightbackground=BORDER,
                                  state="disabled")
            self.log11.pack(side="left", fill="both", expand=True)
            sb11 = ttk.Scrollbar(log_frame, command=self.log11.yview)
            sb11.pack(side="right", fill="y")
            self.log11.configure(yscrollcommand=sb11.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log11.tag_config(tag, foreground=color,
                                       font=("Consolas", 9, weight))

            self.run_btn11 = tk.Button(tab,
                                        text="💵   CONSOLIDATE CASH LEDGERS",
                                        command=self._run_ecashl,
                                        font=("Segoe UI", 13, "bold"),
                                        bg=SUCCESS, fg="white",
                                        activebackground="#16a34a",
                                        activeforeground="white",
                                        relief="flat", cursor="hand2",
                                        pady=14)
            self.run_btn11.pack(side="bottom", fill="x", padx=20, pady=(4, 8))

            info = tk.Frame(tab, bg="#0f2a1a")
            info.pack(side="top", fill="x", padx=20, pady=(10, 0))
            tk.Label(info,
                     text=("  ℹ  Consolidates GSTN portal CSV exports of "
                           "'Electronic Cash Ledger'.\n"
                           "      • Multi-GSTIN / multi-state support — drop "
                           "all CSVs in one folder\n"
                           "      • Cash Ledger has FINER granularity: each "
                           "tax type has Tax / Interest / Penalty / Fee / "
                           "Others / Total\n"
                           "      • Console shows Totals (compact); per-GSTIN "
                           "sheet has full 48-column breakdown\n"
                           "      • Output: Cover + Console + Summary + "
                           "per-GSTIN sheets · TDS Credit / RCM / Voluntary "
                           "payments classified"),
                     font=("Segoe UI", 9), bg="#0f2a1a", fg=SUCCESS,
                     justify="left", wraplength=720
                     ).pack(padx=10, pady=8, anchor="w")

            self._sec(tab, "📁  Folder containing Electronic Cash Ledger "
                           "CSV files (from GSTN portal)")
            sf = tk.Frame(tab, bg=CARD); sf.pack(fill="x", padx=20, pady=(0, 8))
            self.ecash_src_var = tk.StringVar()
            self._entry(sf, self.ecash_src_var).pack(side="left", fill="x",
                                                      expand=True, ipady=6,
                                                      padx=(0, 8))
            self._btn(sf, "Browse", self._browse_ecash_src).pack(side="left")

            self._sec(tab, "💾  Output Excel file  (blank = auto-named)")
            of = tk.Frame(tab, bg=CARD); of.pack(fill="x", padx=20, pady=(0, 4))
            self.ecash_out_var = tk.StringVar()
            self._entry(of, self.ecash_out_var).pack(side="left", fill="x",
                                                      expand=True, ipady=6,
                                                      padx=(0, 8))
            self._btn(of, "Save as", self._browse_ecash_out).pack(side="left")

        def _browse_ecash_src(self):
            d = filedialog.askdirectory()
            if d: self.ecash_src_var.set(d)

        def _browse_ecash_out(self):
            p = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Cash_Ledger_Consolidated.xlsx")
            if p: self.ecash_out_var.set(p)

        def _run_ecashl(self):
            src = self.ecash_src_var.get().strip()
            if not src or not os.path.isdir(src):
                messagebox.showerror("Error", "Select a valid input folder.")
                return
            out_path = self.ecash_out_var.get().strip()
            if not out_path:
                out_path = os.path.join(os.path.dirname(src) or src,
                                        "Cash_Ledger_Consolidated.xlsx")
            try:
                if not out_path.lower().endswith(".xlsx"):
                    out_path += ".xlsx"
                os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
                if os.path.exists(out_path):
                    with open(out_path, "ab"): pass
            except PermissionError:
                messagebox.showerror("Output locked",
                                     f"Output file is open in Excel:\n{out_path}\n\n"
                                     "Close it and re-run.")
                return
            except Exception as e:
                messagebox.showerror("Error",
                                     f"Cannot prepare output path:\n{e}")
                return
            self.run_btn11.configure(state="disabled",
                                      text="⏳  CONSOLIDATING…")
            self._clear(self.log11)
            threading.Thread(target=self._ecashl_worker,
                             args=(src, out_path), daemon=True).start()

        def _ecashl_worker(self, in_folder, out_path):
            try:
                files = [os.path.join(in_folder, f)
                         for f in os.listdir(in_folder)
                         if f.lower().endswith(".csv")]
                if not files:
                    self._log(self.log11,
                              "✗  No CSV files found in this folder.", "err")
                    return
                self._log(self.log11,
                          f"Found {len(files)} CSV file(s). Parsing…", "hdr")

                all_data = []
                for fp in sorted(files):
                    fn = os.path.basename(fp)
                    try:
                        data = parse_ecashl_csv(fp)
                        m = data["meta"]
                        if not m["gstin"]:
                            self._log(self.log11,
                                      f"   ✗  {fn} — no GSTIN found, "
                                      "skipping (not a Cash Ledger CSV?)",
                                      "err")
                            continue
                        all_data.append(data)
                        credit_n = sum(1 for r in data["rows"]
                                       if r["row_type"] == "Credit")
                        debit_n = sum(1 for r in data["rows"]
                                      if r["row_type"] == "Debit")
                        self._log(self.log11,
                                  f"   ✓  {m['state_name']:<20} "
                                  f"({m['gstin']})  {m['from_date']} → "
                                  f"{m['to_date']}  →  {credit_n} Cr, "
                                  f"{debit_n} Dr rows",
                                  "ok")
                    except Exception as e:
                        self._log(self.log11,
                                  f"   ✗  {fn}  →  {type(e).__name__}: {e}",
                                  "err")

                if not all_data:
                    self._log(self.log11,
                              "\n✗  No valid CSVs parsed.", "err")
                    return

                # Dedup overlapping period downloads
                files_before = len(all_data)
                all_data, dedup_stats = dedup_ledger_files(all_data)
                merge_groups = [s for s in dedup_stats if s['files_merged'] > 1]
                if merge_groups:
                    self._log(self.log11,
                              f"\n🔍  Detected {len(merge_groups)} GSTIN(s) "
                              "with multiple overlapping CSV files — merged "
                              "into single entries:", "hdr")
                    total_dups = 0
                    for s in merge_groups:
                        self._log(self.log11,
                                  f"   • {s['state_name']:<18} "
                                  f"({s['gstin']})  →  {s['files_merged']} "
                                  f"files merged, {s['duplicates_skipped']} "
                                  f"duplicate txn(s) skipped, "
                                  f"{s['unique_txns']} unique txn(s)", "ok")
                        total_dups += s['duplicates_skipped']
                    self._log(self.log11,
                              f"\n   {files_before} input file(s)  →  "
                              f"{len(all_data)} unique GSTIN(s)   "
                              f"({total_dups} duplicate txn(s) skipped)",
                              "dim")

                self._log(self.log11,
                          f"\nWriting consolidated output: {out_path}", "hdr")
                n = write_consolidated_ecashl(all_data, out_path)
                self._log(self.log11,
                          f"\n✅  Done!  {n} GSTIN(s) consolidated.\n"
                          f"   →  {out_path}", "ok")
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(os.path.dirname(out_path))
                except Exception: pass
            except Exception as e:
                self._log(self.log11,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn11.configure(state="normal",
                                          text="💵   CONSOLIDATE CASH LEDGERS")

        # ════════════════════════════════════════════════════
        #  TAB 7 — FILE MANAGEMENT (11 sub-tools)
        # ════════════════════════════════════════════════════
        # Catalogue: each tool has key, label, description, expected inputs
        FM_TOOLS = [
            ("search_copy",    "🔍  Search & Copy Files (by name + extension)",
             "Search files containing each name from a list, with chosen extension, "
             "and copy to destination. Optional: also create sub-folder per name."),
            ("list_files",     "📋  List All File Names & Extensions",
             "List every file in a folder with its base name and extension; save to Excel."),
            ("rename_files",   "✏  Bulk Rename Files (from Excel mapping)",
             "Read Old→New name list from Excel, find matching files in source, copy "
             "to destination with new names. Tries multiple formats."),
            ("find_paths",     "🧭  Find Full Path of Each File Name",
             "For each name in your list, find its full file path in a folder; save with Found/Not Found."),
            ("collect_paths",  "🔗  Collect All Files in a Folder (with hyperlinks)",
             "List every file in a folder with its full path + clickable hyperlink in Excel."),
            ("list_subfolders","📂  List Subfolder Names",
             "List the names of all subfolders inside a parent folder."),
            ("rename_folders", "✒  Bulk Rename Folders",
             "Read Old→New folder name list from Excel, copy each old folder to "
             "destination with the new name (preserves contents)."),
            ("jpg_to_pdf",     "🖼  Convert JPG/JPEG/PNG → PDF",
             "Convert every image in a folder to its own PDF (no Adobe required)."),
            ("collect_multi",  "📦  Collect Files from Multiple Subfolders",
             "Copy every file from every subfolder of source into one destination folder. Logs name+extension."),
            ("print_pdfs",     "🖨  Print All PDFs in a Folder (Windows)",
             "Send every PDF in a folder to the default printer using Windows shell. "
             "Requires the system default PDF handler (Adobe / Chrome / Edge etc.)."),
            ("group_by_id",    "👥  Group Files into Per-ID Subfolders",
             "Search files matching each ID/keyword (across pdf/jpg/xlsx/docx), copy "
             "to destination AND auto-create one subfolder per ID. (Like the "
             "'Dynamic Employee Folder' macro.)"),
        ]

        # Per-tool template specs (used by Download Template + Names From Folder)
        _FM_TEMPLATES = {
            "search_copy": {
                "headers":   ["S.No", "Name to Search"],
                "examples":  [[1, "EMP001"], [2, "INV-2024-001"], [3, "CompanyName"]],
                "filename":  "Search_Names_Template.xlsx",
                "title":     "Search & Copy — Name List Template",
                "instructions": (
                    "Fill column B with the names/keywords to search for in file names. "
                    "Each row = one keyword. File extension is NOT included here — "
                    "set that in the tool's 'File extension to match' field. "
                    "Delete the example rows before saving."
                ),
                "pair_mode":   False,
                "from_folder": "files",
            },
            "rename_files": {
                "headers":   ["S.No", "Old File Name", "New File Name"],
                "examples":  [[1, "old_name_1", "new_name_1"],
                              [2, "old_name_2", "new_name_2"]],
                "filename":  "Rename_Files_Template.xlsx",
                "title":     "Bulk Rename Files — Template",
                "instructions": (
                    "Column B: current file name WITHOUT extension. "
                    "Column C: new file name WITHOUT extension. "
                    "Tool will try common extensions (pdf/jpg/docx/xlsx etc.) automatically. "
                    "Delete example rows before saving."
                ),
                "pair_mode":   True,
                "from_folder": "files",
            },
            "find_paths": {
                "headers":   ["S.No", "Name to Find"],
                "examples":  [[1, "EMP001"], [2, "INVOICE-001"]],
                "filename":  "Find_Paths_Template.xlsx",
                "title":     "Find Full Path — Name List Template",
                "instructions": (
                    "Fill column B with the names/keywords to find paths for. "
                    "Tool will scan the folder and return full path for each name."
                ),
                "pair_mode":   False,
                "from_folder": "files",
            },
            "rename_folders": {
                "headers":   ["S.No", "Old Folder Name", "New Folder Name"],
                "examples":  [[1, "OldFolder1", "NewFolder1"],
                              [2, "OldFolder2", "NewFolder2"]],
                "filename":  "Rename_Folders_Template.xlsx",
                "title":     "Bulk Rename Folders — Template",
                "instructions": (
                    "Column B: current folder name (exact match). "
                    "Column C: new folder name. "
                    "Tool will copy each folder to destination with the new name."
                ),
                "pair_mode":   True,
                "from_folder": "folders",
            },
            "group_by_id": {
                "headers":   ["S.No", "ID / Keyword"],
                "examples":  [[1, "EMP001"], [2, "EMP002"], [3, "EMP003"]],
                "filename":  "Group_By_ID_Template.xlsx",
                "title":     "Group Files by ID — Template",
                "instructions": (
                    "Fill column B with the IDs/keywords. Tool will create one "
                    "subfolder per ID and copy matching files into it."
                ),
                "pair_mode":   False,
                "from_folder": "files",
            },
        }

        def _build_file_mgmt_tab(self, tab):
            # Persistent state
            self.fm_tool_var = tk.StringVar(value=self.FM_TOOLS[0][0])
            self.fm_src_var = tk.StringVar()
            self.fm_dst_var = tk.StringVar()
            self.fm_list_path_var = tk.StringVar()
            self.fm_extension_var = tk.StringVar(value="pdf")
            self.fm_out_excel_var = tk.StringVar()
            self.fm_recursive_var = tk.BooleanVar(value=True)

            # Pin log + run button to bottom
            log_section = tk.Frame(tab, bg=CARD)
            log_section.pack(side="bottom", fill="both", expand=True,
                             padx=20, pady=(0, 10))
            tk.Label(log_section, text="📋  Log",
                     font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT2).pack(anchor="w", pady=(4, 4))
            log_frame = tk.Frame(log_section, bg=CARD)
            log_frame.pack(fill="both", expand=True)
            self.log7 = tk.Text(log_frame, bg=ENTRY_BG, fg=TEXT,
                                font=("Consolas", 9),
                                relief="flat", wrap="word", height=4,
                                insertbackground=TEXT,
                                highlightthickness=1, highlightbackground=BORDER,
                                state="disabled")
            self.log7.pack(side="left", fill="both", expand=True)
            sb7 = ttk.Scrollbar(log_frame, command=self.log7.yview)
            sb7.pack(side="right", fill="y")
            self.log7.configure(yscrollcommand=sb7.set)
            for tag, color, weight in [("ok", SUCCESS, "normal"),
                                       ("err", ERROR, "normal"),
                                       ("ren", WARNING, "normal"),
                                       ("hdr", ACCENT2, "bold"),
                                       ("dim", SUBTEXT, "normal")]:
                self.log7.tag_config(tag, foreground=color,
                                     font=("Consolas", 9, weight))

            self.run_btn7 = tk.Button(tab,
                                      text="▶   RUN SELECTED TOOL",
                                      command=self._run_file_mgmt,
                                      font=("Segoe UI", 13, "bold"),
                                      bg=SUCCESS, fg="white",
                                      activebackground="#16a34a",
                                      activeforeground="white",
                                      relief="flat", cursor="hand2",
                                      pady=14)
            self.run_btn7.pack(side="bottom", fill="x", padx=20, pady=(4, 6))

            # ─── Top: Tool selector ─────────────────────────
            self._sec(tab, "🛠  Select Tool")
            sel_frame = tk.Frame(tab, bg=CARD)
            sel_frame.pack(fill="x", padx=20, pady=(0, 8))
            tool_labels = [t[1] for t in self.FM_TOOLS]
            self.fm_combo = ttk.Combobox(sel_frame,
                                         values=tool_labels,
                                         state="readonly",
                                         font=("Segoe UI", 10))
            self.fm_combo.current(0)
            self.fm_combo.pack(fill="x", ipady=4)
            self.fm_combo.bind("<<ComboboxSelected>>", self._fm_on_tool_change)

            # Description label
            self.fm_desc_label = tk.Label(tab,
                                          text=self.FM_TOOLS[0][2],
                                          font=("Segoe UI", 9, "italic"),
                                          bg=CARD, fg=SUBTEXT,
                                          wraplength=720, justify="left")
            self.fm_desc_label.pack(anchor="w", padx=20, pady=(0, 4))

            # ─── Dynamic input container ───────────────────
            self.fm_inputs_container = tk.Frame(tab, bg=CARD)
            self.fm_inputs_container.pack(fill="x", padx=0, pady=(0, 4))

            # Initial render
            self._fm_render_inputs(self.FM_TOOLS[0][0])

        def _fm_on_tool_change(self, event=None):
            idx = self.fm_combo.current()
            key = self.FM_TOOLS[idx][0]
            self.fm_tool_var.set(key)
            self.fm_desc_label.configure(text=self.FM_TOOLS[idx][2])
            self._fm_render_inputs(key)

        def _fm_render_inputs(self, tool_key):
            """Render input fields based on which tool is selected."""
            for child in self.fm_inputs_container.winfo_children():
                child.destroy()
            parent = self.fm_inputs_container

            def add_folder_row(label, var, browse_text="Browse"):
                self._sec(parent, label)
                rf = tk.Frame(parent, bg=CARD)
                rf.pack(fill="x", padx=20, pady=(0, 4))
                self._entry(rf, var).pack(side="left", fill="x",
                                          expand=True, ipady=5, padx=(0, 8))
                def _br():
                    d = filedialog.askdirectory()
                    if d: var.set(d)
                self._btn(rf, browse_text, _br).pack(side="left")

            def add_file_row(label, var, save_as=False, filetypes=None, initialfile=""):
                self._sec(parent, label)
                rf = tk.Frame(parent, bg=CARD)
                rf.pack(fill="x", padx=20, pady=(0, 4))
                self._entry(rf, var).pack(side="left", fill="x",
                                          expand=True, ipady=5, padx=(0, 8))
                ft = filetypes or [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
                def _br():
                    if save_as:
                        p = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=ft,
                                                         initialfile=initialfile)
                    else:
                        p = filedialog.askopenfilename(filetypes=ft)
                    if p: var.set(p)
                self._btn(rf, "Browse" if not save_as else "Save as", _br
                          ).pack(side="left")

            def add_text_row(label, var, hint=""):
                self._sec(parent, label)
                rf = tk.Frame(parent, bg=CARD)
                rf.pack(fill="x", padx=20, pady=(0, 4))
                self._entry(rf, var).pack(fill="x", ipady=5)
                if hint:
                    tk.Label(rf, text=hint, font=("Segoe UI", 8, "italic"),
                             bg=CARD, fg=SUBTEXT).pack(anchor="w", pady=(2, 0))

            def add_check(label, var):
                cf = tk.Frame(parent, bg=CARD)
                cf.pack(fill="x", padx=20, pady=(4, 0))
                tk.Checkbutton(cf, text=f"  {label}", variable=var,
                               bg=CARD, activebackground=CARD, fg=TEXT,
                               font=("Segoe UI", 9), selectcolor=ENTRY_BG,
                               cursor="hand2").pack(anchor="w")

            def add_list_helpers(tk_key):
                """Helper row under list-file input — Download Template + Names From Folder."""
                h = tk.Frame(parent, bg=CARD)
                h.pack(fill="x", padx=20, pady=(0, 4))
                tk.Label(h, text="💡  No list ready?  →",
                         font=("Segoe UI", 9, "italic"),
                         bg=CARD, fg=SUBTEXT).pack(side="left", padx=(0, 8))
                self._btn(h, "📥  Download Template",
                          lambda k=tk_key: self._fm_download_template(k),
                          font=("Segoe UI", 9, "bold"),
                          pady=4, padx=10).pack(side="left", padx=(0, 6))
                self._btn(h, "📁  Names From Folder",
                          lambda k=tk_key: self._fm_names_from_folder(
                              self.fm_list_path_var, k),
                          font=("Segoe UI", 9, "bold"),
                          pady=4, padx=10).pack(side="left")

            # ─── Render fields per tool ──────────────────────
            if tool_key == "search_copy":
                add_folder_row("📁  Source Folder (where to search)", self.fm_src_var)
                add_folder_row("📂  Destination Folder", self.fm_dst_var)
                add_file_row("📝  Excel/CSV with names to search "
                             "(names in column A or B)",
                             self.fm_list_path_var,
                             filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
                add_list_helpers(tool_key)
                add_text_row("📑  File extension to match  (e.g. pdf, jpg, docx, xlsx)",
                             self.fm_extension_var,
                             "Without dot. Use comma to try multiple: pdf,jpg,docx")
                add_check("Search inside subfolders too (recursive)",
                          self.fm_recursive_var)

            elif tool_key == "list_files":
                add_folder_row("📁  Folder", self.fm_src_var)
                add_file_row("💾  Output Excel file (blank = auto-named)",
                             self.fm_out_excel_var, save_as=True,
                             initialfile="File_Names_List.xlsx")
                add_check("Include files from subfolders (recursive)",
                          self.fm_recursive_var)

            elif tool_key == "rename_files":
                add_folder_row("📁  Source Folder", self.fm_src_var)
                add_folder_row("📂  Destination Folder", self.fm_dst_var)
                add_file_row("📝  Excel with Old→New names "
                             "(col A=Old, col B=New, with header row)",
                             self.fm_list_path_var,
                             filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
                add_list_helpers(tool_key)

            elif tool_key == "find_paths":
                add_folder_row("📁  Folder to search", self.fm_src_var)
                add_file_row("📝  Excel/CSV with names to find "
                             "(names in column A or B)",
                             self.fm_list_path_var,
                             filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
                add_list_helpers(tool_key)
                add_file_row("💾  Output Excel file",
                             self.fm_out_excel_var, save_as=True,
                             initialfile="File_Paths.xlsx")
                add_check("Search inside subfolders too (recursive)",
                          self.fm_recursive_var)

            elif tool_key == "collect_paths":
                add_folder_row("📁  Folder", self.fm_src_var)
                add_file_row("💾  Output Excel file (with file paths & hyperlinks)",
                             self.fm_out_excel_var, save_as=True,
                             initialfile="File_Paths_Hyperlinks.xlsx")
                add_check("Include files from subfolders (recursive)",
                          self.fm_recursive_var)

            elif tool_key == "list_subfolders":
                add_folder_row("📁  Parent Folder", self.fm_src_var)
                add_file_row("💾  Output Excel file",
                             self.fm_out_excel_var, save_as=True,
                             initialfile="Subfolder_Names.xlsx")

            elif tool_key == "rename_folders":
                add_folder_row("📁  Source Folder (containing folders to copy)",
                               self.fm_src_var)
                add_folder_row("📂  Destination Folder (where renamed folders go)",
                               self.fm_dst_var)
                add_file_row("📝  Excel with Old→New folder names "
                             "(col A=Old, col B=New, with header row)",
                             self.fm_list_path_var,
                             filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
                add_list_helpers(tool_key)

            elif tool_key == "jpg_to_pdf":
                add_folder_row("📁  Source Folder (with .jpg/.jpeg/.png files)",
                               self.fm_src_var)
                add_folder_row("📂  Destination Folder (PDFs go here)",
                               self.fm_dst_var)

            elif tool_key == "collect_multi":
                add_folder_row("📁  Source Folder (with multiple subfolders)",
                               self.fm_src_var)
                add_folder_row("📂  Destination Folder (everything dumped here)",
                               self.fm_dst_var)
                add_file_row("💾  Output Excel log (optional, blank = no log)",
                             self.fm_out_excel_var, save_as=True,
                             initialfile="Collected_Files_Log.xlsx")

            elif tool_key == "print_pdfs":
                add_folder_row("📁  Folder containing PDF files", self.fm_src_var)
                # Add a warning
                warn = tk.Frame(parent, bg="#2d1f0e")
                warn.pack(fill="x", padx=20, pady=(6, 4))
                tk.Label(warn,
                         text=("  ⚠  Will send every .pdf in this folder to the "
                               "default printer. Make sure printer is ready.\n"
                               "      Works only on Windows."),
                         font=("Segoe UI", 9), bg="#2d1f0e", fg=WARNING,
                         justify="left", wraplength=700
                         ).pack(padx=8, pady=6, anchor="w")

            elif tool_key == "group_by_id":
                add_folder_row("📁  Source Folder (where to search files)",
                               self.fm_src_var)
                add_folder_row("📂  Destination Folder (per-ID subfolders go here)",
                               self.fm_dst_var)
                add_file_row("📝  Excel/CSV with ID/Name list "
                             "(names in column A or B)",
                             self.fm_list_path_var,
                             filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
                add_list_helpers(tool_key)
                add_check("Search inside subfolders too (recursive)",
                          self.fm_recursive_var)

        def _run_file_mgmt(self):
            tool_key = self.fm_tool_var.get()
            self._clear(self.log7)
            self.run_btn7.configure(state="disabled", text="⏳  RUNNING…")
            threading.Thread(target=self._fm_worker,
                             args=(tool_key,), daemon=True).start()

        def _fm_worker(self, tool_key):
            try:
                # Find the tool label
                label = next(t[1] for t in self.FM_TOOLS if t[0] == tool_key)
                self._log(self.log7, f"\n=== {label} ===\n", "hdr")

                # Dispatch
                handler = getattr(self, f"_fm_run_{tool_key}", None)
                if handler is None:
                    self._log(self.log7,
                              f"✗  Tool '{tool_key}' not implemented.", "err")
                    return
                handler()
            except Exception as e:
                self._log(self.log7,
                          f"\n✗  Error: {e}\n{traceback.format_exc()}", "err")
            finally:
                self.run_btn7.configure(state="normal",
                                        text="▶   RUN SELECTED TOOL")

        # ─── Helper: read names list from Excel/CSV ──────────
        def _fm_read_name_list(self, path, want_pairs=False):
            """Read names from .xlsx/.xls/.csv. By default returns a flat list of
            names from col A or B. If want_pairs=True, returns list of (old,new) tuples
            from col A and B."""
            names = []
            pairs = []
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xlsx", ".xlsm", ".xls"):
                import openpyxl
                wb = openpyxl.load_workbook(path, data_only=True)
                ws = wb.active
                first_row_skipped = False
                for r in range(1, ws.max_row + 1):
                    a = ws.cell(r, 1).value
                    b = ws.cell(r, 2).value
                    a_s = str(a).strip() if a is not None else ""
                    b_s = str(b).strip() if b is not None else ""
                    # Skip header row (first row with strings only)
                    if not first_row_skipped:
                        first_row_skipped = True
                        # If looks like header (no number, contains 'name' etc), skip
                        ah = a_s.lower()
                        bh = b_s.lower()
                        if any(w in ah or w in bh for w in
                               ("name", "old", "new", "sr", "s.no", "no.")):
                            continue
                    if want_pairs:
                        if a_s and b_s:
                            pairs.append((a_s, b_s))
                    else:
                        # Pick from column B if it has content (matches VBA),
                        # otherwise A
                        v = b_s if b_s else a_s
                        if v: names.append(v)
                wb.close()
            elif ext == ".csv":
                import csv as _csv
                with open(path, encoding="utf-8", errors="replace") as f:
                    reader = _csv.reader(f)
                    first = True
                    for row in reader:
                        if not row: continue
                        a = row[0].strip() if len(row) > 0 else ""
                        b = row[1].strip() if len(row) > 1 else ""
                        if first:
                            first = False
                            if any(w in (a + b).lower()
                                   for w in ("name", "old", "new", "sr", "s.no", "no.")):
                                continue
                        if want_pairs:
                            if a and b: pairs.append((a, b))
                        else:
                            v = b if b else a
                            if v: names.append(v)
            else:
                raise ValueError(f"Unsupported list file type: {ext}")
            return pairs if want_pairs else names

        # ─── Helper: walk folder & yield file paths ──────────
        def _fm_walk(self, folder, recursive=True):
            if recursive:
                for root, dirs, files in os.walk(folder):
                    for f in files:
                        yield os.path.join(root, f)
            else:
                for f in os.listdir(folder):
                    fp = os.path.join(folder, f)
                    if os.path.isfile(fp):
                        yield fp

        # ─── Helper: write Excel with optional hyperlinks ─────
        def _fm_write_excel(self, headers, rows, out_path,
                            title="", hyperlink_col=None):
            import openpyxl
            from openpyxl.styles import Font as XF, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Data"

            thin = Side(border_style="thin", color="B4B4B4")
            BORDER_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)

            ncols = max(len(headers), 1)
            if title:
                ws.merge_cells(start_row=1, start_column=1,
                               end_row=1, end_column=ncols)
                ws["A1"] = title
                ws["A1"].font = XF(name="Calibri", bold=True, size=14, color="FFFFFF")
                ws["A1"].fill = PatternFill("solid", start_color="1F4E78")
                ws["A1"].alignment = Alignment(horizontal="center",
                                               vertical="center")
                ws.row_dimensions[1].height = 26
                start_row = 3
            else:
                start_row = 1

            for c, h in enumerate(headers, 1):
                cell = ws.cell(row=start_row, column=c, value=h)
                cell.font = XF(name="Calibri", bold=True,
                               color="FFFFFF", size=10)
                cell.fill = PatternFill("solid", start_color="2E75B6")
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center", wrap_text=True)
                cell.border = BORDER_ALL
            ws.row_dimensions[start_row].height = 22

            for ri, row in enumerate(rows, start_row + 1):
                for ci, v in enumerate(row, 1):
                    cell = ws.cell(row=ri, column=ci, value=v)
                    cell.font = XF(name="Calibri", size=10)
                    cell.border = BORDER_ALL
                    if ci == 1 and isinstance(v, int):
                        cell.alignment = Alignment(horizontal="center")
                    else:
                        cell.alignment = Alignment(horizontal="left",
                                                   wrap_text=True)
                    if hyperlink_col and ci == hyperlink_col and v:
                        try:
                            cell.hyperlink = v
                            cell.font = XF(name="Calibri", size=10,
                                           color="0563C1", underline="single")
                        except Exception:
                            pass

            # Auto-width
            for c in range(1, ncols + 1):
                col = get_column_letter(c)
                max_len = max((len(str(ws.cell(r, c).value or ""))
                               for r in range(1, ws.max_row + 1)), default=10)
                ws.column_dimensions[col].width = min(max(max_len + 2, 8), 60)
            ws.freeze_panes = ws.cell(row=start_row + 1, column=1).coordinate
            if rows:
                ws.auto_filter.ref = (
                    f"A{start_row}:{get_column_letter(ncols)}"
                    f"{start_row + len(rows)}")
            wb.save(out_path)

        # ─── Helper: Download Excel template for a tool ──────
        def _fm_download_template(self, tool_key):
            """Create and save an Excel template for the given tool."""
            tmpl = self._FM_TEMPLATES.get(tool_key)
            if not tmpl:
                messagebox.showerror("No template",
                                     f"No template defined for tool '{tool_key}'.")
                return
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=tmpl["filename"],
                title=f"Save '{tmpl['title']}' as...")
            if not save_path:
                return

            try:
                import openpyxl
                from openpyxl.styles import (Font as XF, PatternFill, Alignment,
                                              Border, Side)
                from openpyxl.utils import get_column_letter

                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Template"

                headers = tmpl["headers"]
                examples = tmpl["examples"]
                title_text = tmpl["title"]
                instructions = tmpl["instructions"]
                ncols = len(headers)

                thin = Side(border_style="thin", color="B4B4B4")
                BORDER_ALL = Border(left=thin, right=thin,
                                    top=thin, bottom=thin)

                # Title bar
                ws.merge_cells(start_row=1, start_column=1,
                               end_row=1, end_column=ncols)
                ws["A1"] = title_text
                ws["A1"].font = XF(name="Calibri", bold=True, size=14,
                                   color="FFFFFF")
                ws["A1"].fill = PatternFill("solid", start_color="1F4E78")
                ws["A1"].alignment = Alignment(horizontal="center",
                                               vertical="center")
                ws.row_dimensions[1].height = 28

                # Instructions banner
                ws.merge_cells(start_row=2, start_column=1,
                               end_row=2, end_column=ncols)
                ws["A2"] = "INSTRUCTIONS:  " + instructions
                ws["A2"].font = XF(name="Calibri", italic=True,
                                   size=10, color="7F6000")
                ws["A2"].fill = PatternFill("solid", start_color="FFF2CC")
                ws["A2"].alignment = Alignment(horizontal="left",
                                               vertical="center", wrap_text=True)
                ws.row_dimensions[2].height = 70

                # Header row (row 4)
                for c, h in enumerate(headers, 1):
                    cell = ws.cell(row=4, column=c, value=h)
                    cell.font = XF(name="Calibri", bold=True,
                                   color="FFFFFF", size=11)
                    cell.fill = PatternFill("solid", start_color="2E75B6")
                    cell.alignment = Alignment(horizontal="center",
                                               vertical="center")
                    cell.border = BORDER_ALL
                ws.row_dimensions[4].height = 24

                # Example rows (italic grey — visually obvious they're examples)
                for ri, row in enumerate(examples, 5):
                    for ci, v in enumerate(row, 1):
                        cell = ws.cell(row=ri, column=ci, value=v)
                        cell.font = XF(name="Calibri", italic=True,
                                       size=10, color="808080")
                        cell.alignment = Alignment(
                            horizontal="center" if ci == 1 else "left")
                        cell.border = BORDER_ALL

                # Add a note row below examples
                example_end = 4 + len(examples)
                note_row = example_end + 1
                ws.merge_cells(start_row=note_row, start_column=1,
                               end_row=note_row, end_column=ncols)
                ws.cell(row=note_row, column=1,
                        value="↑  Delete these grey example rows, then fill in your own data below  ↓")
                ws.cell(row=note_row, column=1).font = XF(
                    name="Calibri", italic=True, size=9, color="C00000")
                ws.cell(row=note_row, column=1).alignment = Alignment(
                    horizontal="center")

                # Column widths
                ws.column_dimensions["A"].width = 8
                for c in range(2, ncols + 1):
                    ws.column_dimensions[get_column_letter(c)].width = 32
                ws.freeze_panes = "A5"

                wb.save(save_path)
                self._log(self.log7,
                          f"✅  Template saved → {save_path}", "ok")
                ans = messagebox.askyesno(
                    "Template Created",
                    f"Excel template saved to:\n{save_path}\n\n"
                    "Open it now? You can:\n"
                    "  • Delete grey example rows\n"
                    "  • Fill in your data\n"
                    "  • Save\n"
                    "  • Use 'Browse' above to upload it back")
                if ans and sys.platform.startswith("win"):
                    try: os.startfile(save_path)
                    except Exception: pass
            except Exception as e:
                self._log(self.log7,
                          f"✗  Could not create template: {e}", "err")
                messagebox.showerror("Template error", str(e))

        # ─── Helper: Generate name list FROM a folder ────────
        def _fm_names_from_folder(self, target_var, tool_key):
            """Pick a folder, list its file (or sub-folder) names, save to Excel,
            and populate the list-file field. Supports pair-mode tools where the
            'New' column is pre-filled with the same name for the user to edit."""
            tmpl = self._FM_TEMPLATES.get(tool_key, {})
            from_kind = tmpl.get("from_folder", "files")  # 'files' or 'folders'
            pair_mode = tmpl.get("pair_mode", False)

            title_prompt = ("Pick parent folder (its subfolders will be listed)"
                            if from_kind == "folders" else
                            "Pick folder containing files to list")
            folder = filedialog.askdirectory(title=title_prompt)
            if not folder: return
            if not os.path.isdir(folder):
                messagebox.showerror("Not a folder", folder); return

            # Collect names
            names = []
            try:
                if from_kind == "folders":
                    names = sorted([e for e in os.listdir(folder)
                                    if os.path.isdir(os.path.join(folder, e))])
                else:
                    seen = set()
                    for f in sorted(os.listdir(folder)):
                        fp = os.path.join(folder, f)
                        if os.path.isfile(fp):
                            base = os.path.splitext(f)[0]
                            if base and base not in seen:
                                seen.add(base)
                                names.append(base)
            except Exception as e:
                messagebox.showerror("Could not read folder", str(e))
                return

            if not names:
                kind_label = "subfolders" if from_kind == "folders" else "files"
                messagebox.showinfo("Nothing found",
                                    f"No {kind_label} found in:\n{folder}")
                return

            # Save dialog
            default_name = ("Rename_From_Folder.xlsx" if pair_mode
                            else "Names_From_Folder.xlsx")
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_name,
                title="Save the generated list as...")
            if not save_path: return

            # Build rows from template headers
            headers = tmpl.get("headers",
                               ["S.No", "Old Name", "New Name"]
                               if pair_mode else ["S.No", "Name"])
            rows = []
            for i, name in enumerate(names, 1):
                if pair_mode:
                    rows.append([i, name, name])  # pre-fill new = old
                else:
                    rows.append([i, name])

            try:
                # Use the existing writer (already has a polished output)
                kind_label = ("subfolders" if from_kind == "folders" else "files")
                self._fm_write_excel(
                    headers, rows, save_path,
                    title=f"Names from {os.path.basename(folder)}  "
                          f"({len(names)} {kind_label})")
            except Exception as e:
                self._log(self.log7,
                          f"✗  Failed to write list: {e}", "err")
                return

            # Populate the field
            target_var.set(save_path)
            self._log(self.log7,
                      f"✅  Generated list with {len(names)} {kind_label}  →  "
                      f"{save_path}", "ok")
            target_var_path = save_path
            if pair_mode:
                ans = messagebox.askyesno(
                    "List Created",
                    f"List created with {len(names)} entries:\n{save_path}\n\n"
                    "The 'Old' and 'New' columns are pre-filled with the SAME names.\n\n"
                    "Open it in Excel now to edit the NEW names?")
                if ans and sys.platform.startswith("win"):
                    try: os.startfile(target_var_path)
                    except Exception: pass
            else:
                messagebox.showinfo(
                    "List Created",
                    f"List created with {len(names)} entries:\n{save_path}\n\n"
                    "The Excel field is already filled with this path. "
                    "Click RUN to proceed.")

        # ─── Tool 1: Search & Copy ───────────────────────────
        def _fm_run_search_copy(self):
            src = self.fm_src_var.get().strip()
            dst = self.fm_dst_var.get().strip()
            lst = self.fm_list_path_var.get().strip()
            exts_raw = self.fm_extension_var.get().strip()
            recursive = self.fm_recursive_var.get()

            if not (src and dst and lst and exts_raw):
                self._log(self.log7,
                          "✗  Please fill Source / Destination / List file / Extension.",
                          "err"); return
            if not os.path.isdir(src):
                self._log(self.log7, f"✗  Source folder not found: {src}", "err"); return
            os.makedirs(dst, exist_ok=True)
            exts = [e.strip().lower().lstrip(".") for e in exts_raw.split(",") if e.strip()]
            names = self._fm_read_name_list(lst)
            self._log(self.log7, f"Names to search: {len(names)}", "hdr")
            self._log(self.log7, f"Extensions: {exts}", "dim")
            self._log(self.log7, f"Recursive: {recursive}", "dim")

            done = 0; pending = 0
            for name in names:
                found_any = False
                for fp in self._fm_walk(src, recursive=recursive):
                    fname = os.path.basename(fp)
                    base, fext = os.path.splitext(fname)
                    fext = fext.lower().lstrip(".")
                    if fext not in exts: continue
                    if name.lower() in fname.lower():
                        try:
                            shutil.copy2(fp, os.path.join(dst, fname))
                            found_any = True
                            self._log(self.log7,
                                      f"   ✓  {name}  →  {fname}", "ok")
                        except Exception as e:
                            self._log(self.log7,
                                      f"   ✗  copy failed: {fname}: {e}", "err")
                if found_any:
                    done += 1
                else:
                    pending += 1
                    self._log(self.log7, f"   ·  {name}  →  pending", "ren")
            self._log(self.log7,
                      f"\n✅  Done: {done}, Pending: {pending} of {len(names)}",
                      "ok")
            try:
                if sys.platform.startswith("win"): os.startfile(dst)
            except Exception: pass

        # ─── Tool 2: List Files ──────────────────────────────
        def _fm_run_list_files(self):
            src = self.fm_src_var.get().strip()
            out = self.fm_out_excel_var.get().strip()
            recursive = self.fm_recursive_var.get()
            if not src or not os.path.isdir(src):
                self._log(self.log7, "✗  Pick a valid source folder.", "err"); return
            if not out:
                out = os.path.join(os.path.dirname(src) or src,
                                   "File_Names_List.xlsx")
            rows = []
            for fp in self._fm_walk(src, recursive=recursive):
                rel = os.path.relpath(fp, src)
                fname = os.path.basename(fp)
                base, ext = os.path.splitext(fname)
                rows.append([len(rows) + 1, base, ext.lstrip("."),
                             os.path.dirname(rel) or "."])
            headers = ["S.No.", "File Name (no ext)", "Extension", "Sub-folder"]
            self._fm_write_excel(headers, rows, out,
                                 title=f"File Names — {os.path.basename(src)}")
            self._log(self.log7,
                      f"✅  Listed {len(rows)} file(s)  →  {out}", "ok")
            try:
                if sys.platform.startswith("win"):
                    os.startfile(os.path.dirname(out))
            except Exception: pass

        # ─── Tool 3: Bulk Rename Files ───────────────────────
        def _fm_run_rename_files(self):
            src = self.fm_src_var.get().strip()
            dst = self.fm_dst_var.get().strip()
            lst = self.fm_list_path_var.get().strip()
            if not (src and dst and lst):
                self._log(self.log7,
                          "✗  Need Source, Destination and Excel mapping.", "err"); return
            if not os.path.isdir(src):
                self._log(self.log7, f"✗  Source folder not found: {src}", "err"); return
            os.makedirs(dst, exist_ok=True)
            pairs = self._fm_read_name_list(lst, want_pairs=True)
            self._log(self.log7, f"Rename pairs: {len(pairs)}", "hdr")

            EXTS = ("pdf", "jpg", "jpeg", "png", "docx", "xlsx", "xls", "xlsm", "xlsb")
            done = pending = 0
            for old, new in pairs:
                found = False
                for ext in EXTS:
                    cand = os.path.join(src, f"{old}.{ext}")
                    if os.path.isfile(cand):
                        try:
                            shutil.copy2(cand, os.path.join(dst, f"{new}.{ext}"))
                            self._log(self.log7,
                                      f"   ✓  {old}.{ext}  →  {new}.{ext}", "ok")
                            found = True
                            break
                        except Exception as e:
                            self._log(self.log7, f"   ✗  copy failed: {e}", "err")
                if found:
                    done += 1
                else:
                    pending += 1
                    self._log(self.log7, f"   ·  {old}  →  pending (file not found)",
                              "ren")
            self._log(self.log7,
                      f"\n✅  Renamed: {done}, Pending: {pending}", "ok")
            try:
                if sys.platform.startswith("win"): os.startfile(dst)
            except Exception: pass

        # ─── Tool 4: Find Full Path of Each File Name ────────
        def _fm_run_find_paths(self):
            src = self.fm_src_var.get().strip()
            lst = self.fm_list_path_var.get().strip()
            out = self.fm_out_excel_var.get().strip()
            recursive = self.fm_recursive_var.get()
            if not (src and lst):
                self._log(self.log7, "✗  Need folder and list file.", "err"); return
            if not out:
                out = os.path.join(os.path.dirname(lst) or ".",
                                   "File_Paths.xlsx")
            names = self._fm_read_name_list(lst)
            self._log(self.log7, f"Searching {len(names)} name(s)…", "hdr")

            # Build cache of all files
            cache = list(self._fm_walk(src, recursive=recursive))
            self._log(self.log7, f"   Scanned {len(cache)} file(s) in folder.",
                      "dim")

            rows = []
            found = nf = 0
            for i, name in enumerate(names, 1):
                match = None
                for fp in cache:
                    if name.lower() in os.path.basename(fp).lower():
                        match = fp
                        break
                if match:
                    rows.append([i, name, match, "Found"])
                    found += 1
                else:
                    rows.append([i, name, "", "Not Found"])
                    nf += 1
            headers = ["S.No.", "File Name", "Full Path", "Status"]
            self._fm_write_excel(headers, rows, out,
                                 title=f"File Paths — {os.path.basename(src)}",
                                 hyperlink_col=3)
            self._log(self.log7,
                      f"\n✅  Found: {found}, Not Found: {nf}  →  {out}", "ok")
            try:
                if sys.platform.startswith("win"):
                    os.startfile(os.path.dirname(out))
            except Exception: pass

        # ─── Tool 5: Collect All Files w/ hyperlinks ─────────
        def _fm_run_collect_paths(self):
            src = self.fm_src_var.get().strip()
            out = self.fm_out_excel_var.get().strip()
            recursive = self.fm_recursive_var.get()
            if not src or not os.path.isdir(src):
                self._log(self.log7, "✗  Pick a valid source folder.", "err"); return
            if not out:
                out = os.path.join(os.path.dirname(src) or src,
                                   "File_Paths_Hyperlinks.xlsx")
            rows = []
            for fp in self._fm_walk(src, recursive=recursive):
                rows.append([len(rows) + 1, os.path.basename(fp), fp, fp])
            headers = ["S.No.", "File Name", "Full Path", "Open File"]
            self._fm_write_excel(headers, rows, out,
                                 title=f"All Files in {os.path.basename(src)}",
                                 hyperlink_col=4)
            self._log(self.log7,
                      f"✅  Listed {len(rows)} file(s)  →  {out}", "ok")
            try:
                if sys.platform.startswith("win"):
                    os.startfile(os.path.dirname(out))
            except Exception: pass

        # ─── Tool 6: List Subfolders ─────────────────────────
        def _fm_run_list_subfolders(self):
            src = self.fm_src_var.get().strip()
            out = self.fm_out_excel_var.get().strip()
            if not src or not os.path.isdir(src):
                self._log(self.log7, "✗  Pick a valid parent folder.", "err"); return
            if not out:
                out = os.path.join(os.path.dirname(src) or src,
                                   "Subfolder_Names.xlsx")
            rows = []
            for entry in sorted(os.listdir(src)):
                fp = os.path.join(src, entry)
                if os.path.isdir(fp):
                    n_files = sum(1 for _ in self._fm_walk(fp, recursive=True))
                    rows.append([len(rows) + 1, entry, fp, n_files])
            headers = ["S.No.", "Folder Name", "Full Path", "File Count (recursive)"]
            self._fm_write_excel(headers, rows, out,
                                 title=f"Subfolders of {os.path.basename(src)}")
            self._log(self.log7,
                      f"✅  Listed {len(rows)} subfolder(s)  →  {out}", "ok")
            try:
                if sys.platform.startswith("win"):
                    os.startfile(os.path.dirname(out))
            except Exception: pass

        # ─── Tool 7: Bulk Rename Folders ─────────────────────
        def _fm_run_rename_folders(self):
            src = self.fm_src_var.get().strip()
            dst = self.fm_dst_var.get().strip()
            lst = self.fm_list_path_var.get().strip()
            if not (src and dst and lst):
                self._log(self.log7,
                          "✗  Need Source, Destination and Excel mapping.", "err"); return
            os.makedirs(dst, exist_ok=True)
            pairs = self._fm_read_name_list(lst, want_pairs=True)
            self._log(self.log7, f"Folder rename pairs: {len(pairs)}", "hdr")
            done = pending = 0
            for old, new in pairs:
                src_fp = os.path.join(src, old)
                dst_fp = os.path.join(dst, new)
                if os.path.isdir(src_fp):
                    try:
                        shutil.copytree(src_fp, dst_fp, dirs_exist_ok=True)
                        self._log(self.log7,
                                  f"   ✓  {old}  →  {new}", "ok")
                        done += 1
                    except Exception as e:
                        pending += 1
                        self._log(self.log7,
                                  f"   ✗  {old}  failed: {e}", "err")
                else:
                    pending += 1
                    self._log(self.log7,
                              f"   ·  {old}  →  folder not found", "ren")
            self._log(self.log7,
                      f"\n✅  Renamed: {done}, Pending: {pending}", "ok")
            try:
                if sys.platform.startswith("win"): os.startfile(dst)
            except Exception: pass

        # ─── Tool 8: JPG → PDF ───────────────────────────────
        def _fm_run_jpg_to_pdf(self):
            src = self.fm_src_var.get().strip()
            dst = self.fm_dst_var.get().strip()
            if not (src and dst):
                self._log(self.log7, "✗  Need Source and Destination.", "err"); return
            if not HAS_PILLOW:
                self._log(self.log7,
                          "✗  Pillow not installed. Try:  pip install pillow",
                          "err"); return
            os.makedirs(dst, exist_ok=True)
            from PIL import Image
            done = fail = 0
            for f in sorted(os.listdir(src)):
                fp = os.path.join(src, f)
                if not os.path.isfile(fp): continue
                ext = os.path.splitext(f)[1].lower()
                if ext not in (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"):
                    continue
                try:
                    img = Image.open(fp)
                    if img.mode in ("RGBA", "P"):
                        img = img.convert("RGB")
                    out_name = os.path.splitext(f)[0] + ".pdf"
                    img.save(os.path.join(dst, out_name), "PDF",
                             resolution=150.0)
                    self._log(self.log7, f"   ✓  {f}  →  {out_name}", "ok")
                    done += 1
                except Exception as e:
                    self._log(self.log7, f"   ✗  {f}: {e}", "err")
                    fail += 1
            self._log(self.log7,
                      f"\n✅  Converted: {done}, Failed: {fail}", "ok")
            try:
                if sys.platform.startswith("win"): os.startfile(dst)
            except Exception: pass

        # ─── Tool 9: Collect From Multiple Subfolders ────────
        def _fm_run_collect_multi(self):
            src = self.fm_src_var.get().strip()
            dst = self.fm_dst_var.get().strip()
            out = self.fm_out_excel_var.get().strip()
            if not (src and dst):
                self._log(self.log7, "✗  Need Source and Destination.", "err"); return
            os.makedirs(dst, exist_ok=True)
            rows = []
            done = fail = 0
            for sub in sorted(os.listdir(src)):
                sub_fp = os.path.join(src, sub)
                if not os.path.isdir(sub_fp): continue
                for f in os.listdir(sub_fp):
                    fp = os.path.join(sub_fp, f)
                    if not os.path.isfile(fp): continue
                    try:
                        out_path = os.path.join(dst, f)
                        # If duplicate, add counter
                        if os.path.exists(out_path):
                            base, ext = os.path.splitext(f)
                            i = 1
                            while os.path.exists(os.path.join(dst,
                                                              f"{base}_{i}{ext}")):
                                i += 1
                            out_path = os.path.join(dst, f"{base}_{i}{ext}")
                        shutil.copy2(fp, out_path)
                        done += 1
                        base, ext = os.path.splitext(f)
                        rows.append([len(rows) + 1, base, ext.lstrip("."),
                                     sub, os.path.basename(out_path)])
                    except Exception as e:
                        fail += 1
                        self._log(self.log7,
                                  f"   ✗  {f} from {sub}: {e}", "err")
            self._log(self.log7,
                      f"\n✅  Copied: {done}, Failed: {fail}", "ok")
            if out and rows:
                headers = ["S.No.", "File Name (no ext)", "Extension",
                           "Source Subfolder", "Saved As"]
                self._fm_write_excel(headers, rows, out,
                                     title="Files Collected from Subfolders")
                self._log(self.log7, f"   Log saved: {out}", "ok")
            try:
                if sys.platform.startswith("win"): os.startfile(dst)
            except Exception: pass

        # ─── Tool 10: Print PDFs (Windows shell) ─────────────
        def _fm_run_print_pdfs(self):
            src = self.fm_src_var.get().strip()
            if not src or not os.path.isdir(src):
                self._log(self.log7, "✗  Pick a valid folder.", "err"); return
            if not sys.platform.startswith("win"):
                self._log(self.log7,
                          "✗  This feature works only on Windows.", "err"); return
            pdfs = [os.path.join(src, f) for f in sorted(os.listdir(src))
                    if f.lower().endswith(".pdf")
                    and os.path.isfile(os.path.join(src, f))]
            if not pdfs:
                self._log(self.log7, "✗  No PDFs found in folder.", "err"); return
            self._log(self.log7,
                      f"Sending {len(pdfs)} PDF(s) to default printer…", "hdr")
            done = fail = 0
            for fp in pdfs:
                try:
                    os.startfile(fp, "print")
                    self._log(self.log7,
                              f"   ✓  {os.path.basename(fp)}", "ok")
                    done += 1
                except Exception as e:
                    self._log(self.log7,
                              f"   ✗  {os.path.basename(fp)}: {e}", "err")
                    fail += 1
            self._log(self.log7,
                      f"\n✅  Sent: {done}, Failed: {fail}", "ok")

        # ─── Tool 11: Group Files into Per-ID Subfolders ─────
        def _fm_run_group_by_id(self):
            src = self.fm_src_var.get().strip()
            dst = self.fm_dst_var.get().strip()
            lst = self.fm_list_path_var.get().strip()
            recursive = self.fm_recursive_var.get()
            if not (src and dst and lst):
                self._log(self.log7,
                          "✗  Need Source, Destination, and ID list.", "err"); return
            os.makedirs(dst, exist_ok=True)
            ids = self._fm_read_name_list(lst)
            self._log(self.log7, f"IDs to group: {len(ids)}", "hdr")
            EXTS = ("pdf", "jpg", "jpeg", "png", "docx", "xlsx",
                    "xls", "xlsm", "xlsb")

            # Cache all files once
            cache = list(self._fm_walk(src, recursive=recursive))

            done = pending = 0
            for emp_id in ids:
                emp_folder = os.path.join(dst, emp_id)
                found_any = False
                matches = []
                for fp in cache:
                    fname = os.path.basename(fp)
                    base, ext = os.path.splitext(fname)
                    ext_clean = ext.lower().lstrip(".")
                    if ext_clean not in EXTS: continue
                    if emp_id.lower() in fname.lower():
                        matches.append(fp)
                if matches:
                    os.makedirs(emp_folder, exist_ok=True)
                    for src_fp in matches:
                        try:
                            shutil.copy2(src_fp,
                                         os.path.join(emp_folder,
                                                      os.path.basename(src_fp)))
                            found_any = True
                        except Exception as e:
                            self._log(self.log7,
                                      f"   ✗  {os.path.basename(src_fp)}: {e}",
                                      "err")
                if found_any:
                    done += 1
                    self._log(self.log7,
                              f"   ✓  {emp_id}  →  {len(matches)} file(s)", "ok")
                else:
                    pending += 1
                    self._log(self.log7,
                              f"   ·  {emp_id}  →  no files found", "ren")
            self._log(self.log7,
                      f"\n✅  Done: {done}, Pending: {pending} of {len(ids)}",
                      "ok")
            try:
                if sys.platform.startswith("win"): os.startfile(dst)
            except Exception: pass

    # ── launch ────────────────────────────────────────────────
    App().mainloop()

except Exception:
    err = traceback.format_exc()
    try:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("GST Tools Suite — Startup Error",
                             f"App failed to start:\n\n{err}")
    except Exception:
        pass
    try:
        with open(os.path.join(os.path.expanduser("~"),
                               "gst_tools_error.txt"), "w") as f:
            f.write(err)
    except Exception:
        pass
