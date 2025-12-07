from PIL import Image, ImageOps, ImageChops
import os
from io import BytesIO
import tempfile
import subprocess
import shutil

# ---------------- Plugin registration (safe) ----------------
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

try:
    import pillow_avif  # registers AVIF support
except Exception:
    pass


# ---------------- Size dictionaries in mm ----------------
A_SIZES_MM = {
    "A1": (594, 841),
    "A2": (420, 594),
    "A3": (297, 420),
    "A4": (210, 297),
}

TWO_PER_THREE_SIZES_MM = {
    "24X36": (610, 910),
    "20X30": (510, 760),
    "16X24": (410, 610),
    "12X18": (300, 460),
    "8X12": (200, 300),
}

FOUR_PER_FIVE_SIZES_MM = {
    "24X30": (600, 760),
    "20X25": (510, 630),
    "16X20": (400, 500),
    "12X15": (300, 380),
    "8X10": (200, 250),
}

THREE_PER_FOUR_SIZES_MM = {
    "18X24": (460, 610),
    "15X20": (380, 510),
    "12X16": (300, 400),
    "9X12": (230, 300),
}

ELEVEN_X_FOURTEEN_SIZES_MM = {
    "11X14": (280, 360),
}


# ---------------- mm -> px helpers ----------------
def mm_to_px(mm, dpi):
    inches = mm / 25.4
    return int(round(inches * dpi))

def paper_size_px(size_key="A4", dpi=300, landscape="vertical", sizes_mm=A_SIZES_MM):

    """
    orientation:
      - "vertical"  -> use (w, h) as defined
      - "horizontal" -> swap to (h, w)
    """

    key = size_key.upper()
    if key not in sizes_mm:
        raise ValueError(f"Unknown size_key '{size_key}'. Available: {list(sizes_mm.keys())}")

    w_mm, h_mm = sizes_mm[key]
    if landscape.lower() in ("horizontal", "landscape", "h"):
        w_mm, h_mm = h_mm, w_mm
    return mm_to_px(w_mm, dpi), mm_to_px(h_mm, dpi)


# ---------------- Extensions ----------------
RAW_EXTS = {
    ".cr2", ".cr3", ".nef", ".arw", ".dng", ".rw2", ".orf", ".raf", ".pef"
}
VECTOR_EXTS = {".svg"}
PS_PDF_EXTS = {".eps", ".ps", ".ai", ".pdf"}  # require Ghostscript


# ---------------- Path resolver ----------------
def resolve_path(path: str) -> str:
    if os.path.isabs(path) and os.path.exists(path):
        return path

    cwd_candidate = os.path.abspath(path)
    if os.path.exists(cwd_candidate):
        return cwd_candidate

    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_candidate = os.path.join(script_dir, path)
        if os.path.exists(script_candidate):
            return script_candidate
    except Exception:
        pass

    return cwd_candidate


# ---------------- Auto-crop helpers ----------------
def autocrop_white(im: Image.Image, bg=(255, 255, 255)):
    if im.mode != "RGB":
        im = im.convert("RGB")
    bg_img = Image.new("RGB", im.size, bg)
    diff = ImageChops.difference(im, bg_img)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)
    return im

def autocrop_auto(im: Image.Image):
    try:
        rgba = im.convert("RGBA")
        alpha = rgba.split()[-1]
        bbox = alpha.getbbox()
        if bbox:
            return rgba.crop(bbox)
    except Exception:
        pass

    try:
        return autocrop_white(im)
    except Exception:
        return im

def flatten_alpha_to_bg(im: Image.Image, bg_color=(255, 255, 255)):
    if im.mode != "RGBA":
        return im
    bg = Image.new("RGB", im.size, bg_color)
    bg.paste(im, mask=im.split()[-1])
    return bg


# ---------------- Ghostscript helpers ----------------
def _find_gs_exe():
    for name in ("gs", "gswin64c", "gswin32c"):
        if shutil.which(name):
            return name
    return None

def _rasterize_with_ghostscript(path, dpi=300):
    gs = _find_gs_exe()
    if not gs:
        raise RuntimeError("Ghostscript not found in PATH. Install Ghostscript to open EPS/PS/AI/PDF.")

    path_abs = resolve_path(path)
    if not os.path.exists(path_abs):
        raise FileNotFoundError(f"File not found: {path_abs}")

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp_path = tmp.name
    tmp.close()

    ext = os.path.splitext(path_abs)[1].lower()

    cmd = [gs, "-dSAFER", "-dBATCH", "-dNOPAUSE"]

    if ext in {".eps", ".ps", ".ai"}:
        cmd += ["-dEPSCrop"]
    elif ext == ".pdf":
        cmd += ["-dUseCropBox"]

    cmd += [
        "-sDEVICE=pngalpha",
        f"-r{dpi}",
        f"-sOutputFile={tmp_path}",
        path_abs
    ]

    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        img = Image.open(tmp_path)
        img._temp_raster_path = tmp_path
        return img
    except subprocess.CalledProcessError as e:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        raise RuntimeError(
            "Ghostscript failed to render this file.\n"
            f"File: {path_abs}\n"
            f"Command: {' '.join(cmd)}\n\n"
            f"--- Ghostscript stderr ---\n{(e.stderr or '')[:2000]}"
        )


# ---------------- Inkscape SVG fallback ----------------
def _find_inkscape_exe():
    # Allow optional override by env var
    env = os.environ.get("INKSCAPE_EXE")
    if env and os.path.exists(env):
        return env

    p = shutil.which("inkscape")
    if p:
        return p

    # Common Windows locations
    candidates = [
        r"C:\Program Files\Inkscape\bin\inkscape.exe",
        r"C:\Program Files\Inkscape\inkscape.exe",
        r"C:\Program Files (x86)\Inkscape\bin\inkscape.exe",
        r"C:\Program Files (x86)\Inkscape\inkscape.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c

    return None

def _rasterize_svg_with_inkscape(path, dpi=300):
    inkscape = _find_inkscape_exe()
    if not inkscape:
        raise RuntimeError(
            "Inkscape not found in PATH. Install Inkscape and add it to PATH "
            "to enable SVG -> PNG conversion on Windows."
        )

    path_abs = resolve_path(path)
    if not os.path.exists(path_abs):
        raise FileNotFoundError(f"File not found: {path_abs}")

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp_path = tmp.name
    tmp.close()

    cmd = [
        inkscape,
        path_abs,
        "--export-type=png",
        f"--export-dpi={dpi}",
        f"--export-filename={tmp_path}",
    ]

    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        img = Image.open(tmp_path)
        img._temp_raster_path = tmp_path
        return img
    except subprocess.CalledProcessError as e:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
        raise RuntimeError(
            "Inkscape failed to render this SVG.\n"
            f"File: {path_abs}\n"
            f"Command: {' '.join(cmd)}\n\n"
            f"--- Inkscape stderr ---\n{(e.stderr or '')[:2000]}"
        )


# ---------------- "Open almost any image" ----------------
def open_image_any(path, raster_dpi=300, skip_gs_files=False, skip_svg_files=True):
    path_abs = resolve_path(path)
    if not os.path.exists(path_abs):
        raise FileNotFoundError(f"File not found: {path_abs}")

    ext = os.path.splitext(path_abs)[1].lower()

    # 1) GS formats FIRST
    if ext in PS_PDF_EXTS:
        try:
            return _rasterize_with_ghostscript(path_abs, dpi=raster_dpi)
        except Exception:
            if skip_gs_files:
                return None
            raise

    # 2) Pillow raster
    try:
        return Image.open(path_abs)
    except Exception:
        pass

    # 3) RAW fallback
    if ext in RAW_EXTS:
        try:
            import rawpy
            with rawpy.imread(path_abs) as raw:
                rgb = raw.postprocess()
            return Image.fromarray(rgb)
        except Exception as e:
            raise ValueError(
                f"RAW file detected but could not decode: {path_abs}\n"
                f"Reason: {e}\n"
                "Install: pip install rawpy numpy"
            )

    # 4) SVG fallback
    if ext in VECTOR_EXTS:
        # Try CairoSVG first
        try:
            import cairosvg
            try:
                png_bytes = cairosvg.svg2png(url=path_abs, dpi=raster_dpi)
                return Image.open(BytesIO(png_bytes))
            except OSError:
                pass
            except Exception:
                pass
        except Exception:
            pass

        # Try Inkscape fallback
        try:
            return _rasterize_svg_with_inkscape(path_abs, dpi=raster_dpi)
        except Exception:
            if skip_svg_files:
                return None
            raise ValueError(
                "SVG file detected but cannot be rasterized on this system.\n"
                "Fix options:\n"
                "1) Install Inkscape and add to PATH (recommended)\n"
                "2) Or convert SVG to PNG first\n"
                f"File: {path_abs}"
            )

    # ✅ IMPORTANT: final fallback error
    raise ValueError(f"Unsupported or unreadable image: {path_abs}")


# ---------------- Resize to target paper size ----------------
def resize_to_a_size(
    src_path,
    out_path,
    sizes_mm=A_SIZES_MM,   # dict of sizes
    size_key="A4",         # key inside dict
    dpi=300,
    mode="fit",            # "fit", "fill", "stretch"
    landscape="vertical",
    bg_color=(255, 255, 255),
    skip_gs_files=True,
    skip_svg_files=True,
    autocrop=True,
    flatten_alpha=True
):
    target_w, target_h = paper_size_px(
        size_key=size_key,
        dpi=dpi,
        landscape=landscape,
        sizes_mm=sizes_mm
    )

    temp_raster = None
    im = open_image_any(
        src_path,
        raster_dpi=dpi,
        skip_gs_files=skip_gs_files,
        skip_svg_files=skip_svg_files
    )

    if im is None:
        ext = os.path.splitext(resolve_path(src_path))[1].lower()
        if ext in PS_PDF_EXTS:
            return {
                "error": "GS_REQUIRED",
                "message": "EPS/AI/PDF needs Ghostscript. Install Ghostscript or convert to PNG/JPG first.",
                "src_path": resolve_path(src_path)
            }
        if ext in VECTOR_EXTS:
            return {
                "error": "SVG_TOOL_REQUIRED",
                "message": "SVG needs Inkscape on Windows (recommended) or a working Cairo runtime.",
                "src_path": resolve_path(src_path)
            }
        return {
            "error": "UNREADABLE",
            "message": "File could not be opened.",
            "src_path": resolve_path(src_path)
        }

    try:
        temp_raster = getattr(im, "_temp_raster_path", None)

        # EXIF orientation
        try:
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass

        # Crop artboard whitespace (useful for EPS/PDF/SVG)
        if autocrop:
            im = autocrop_auto(im)

        # Fix black/transparent artifacts
        if flatten_alpha:
            im = flatten_alpha_to_bg(im.convert("RGBA"), bg_color)

        has_alpha = im.mode in ("RGBA", "LA") or ("transparency" in im.info)

        if mode == "stretch":
            out = im.resize((target_w, target_h), Image.LANCZOS)

        elif mode == "fill":
            out = ImageOps.fit(im, (target_w, target_h), method=Image.LANCZOS, centering=(0.5, 0.5))

        else:  # "fit" with upscale allowed
            ratio = min(target_w / im.width, target_h / im.height)
            new_w = max(1, int(round(im.width * ratio)))
            new_h = max(1, int(round(im.height * ratio)))
            im_copy = im.resize((new_w, new_h), Image.LANCZOS)

            if has_alpha and out_path.lower().endswith(".png"):
                out = Image.new("RGBA", (target_w, target_h), (255, 255, 255, 0))
            else:
                out = Image.new("RGB", (target_w, target_h), bg_color)

            x = (target_w - im_copy.width) // 2
            y = (target_h - im_copy.height) // 2
            out.paste(im_copy, (x, y))

        # JPG doesn't support alpha
        if not out_path.lower().endswith(".png") and out.mode == "RGBA":
            out = out.convert("RGB")

        out.save(out_path, dpi=(dpi, dpi))

        return {
            "size_dict": (
                "A_SIZES_MM" if sizes_mm is A_SIZES_MM else
                "TWO_PER_THREE_SIZES_MM" if sizes_mm is TWO_PER_THREE_SIZES_MM else
                "FOUR_PER_FIVE_SIZES_MM" if sizes_mm is FOUR_PER_FIVE_SIZES_MM else
                "THREE_PER_FOUR_SIZES_MM" if sizes_mm is THREE_PER_FOUR_SIZES_MM else
                "ELEVEN_X_FOURTEEN_SIZES_MM" if sizes_mm is ELEVEN_X_FOURTEEN_SIZES_MM else
                "CUSTOM"
            ),
            "size_key": size_key.upper(),
            "dpi": dpi,
            "mode": mode,
            "landscape": landscape,
            "target_px": (target_w, target_h),
            "output": out_path,
            "src_path": resolve_path(src_path)
        }

    finally:
        try:
            im.close()
        except Exception:
            pass

        if temp_raster:
            try:
                os.remove(temp_raster)
            except Exception:
                pass


# ---------------- Example usage ----------------
if __name__ == "__main__":
    # Example 1) A-series
    info = resize_to_a_size(
        src_path="Image/16.ai",
        out_path="Image/test.png",
        sizes_mm=A_SIZES_MM,
        size_key="A4",
        dpi=300,
        mode="fit",
        landscape="vertical",
        skip_gs_files=False,
        skip_svg_files=False
    )
    print(info)
