FROM python:3.11-slim

# ---- System deps for imaging + vector ----
RUN apt-get update && apt-get install -y --no-install-recommends \
    ghostscript \
    inkscape \
    libcairo2 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf2.0-0 \
    libjpeg62-turbo \
    libpng16-16 \
    libtiff6 \
    libwebp7 \
    libopenjp2-7 \
    libheif1 \
    libde265-0 \
    libraw20 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# ---- Python deps ----
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ---- Your script ----
COPY get_dpi_pic.py .

# Default command (you can override)
CMD ["python", "get_dpi_pic.py"]
