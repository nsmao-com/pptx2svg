FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    WORK_ROOT=/tmp/ppt-to-svg \
    APP_PORT=8321 \
    DOWNLOAD_TIMEOUT_SECONDS=120 \
    COMMAND_TIMEOUT_SECONDS=240 \
    MAX_DOWNLOAD_MB=100 \
    LIBREOFFICE_COMMAND=soffice \
    MUPDF_COMMAND=mutool

RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    fontconfig \
    libfontconfig1 \
    libfreetype6 \
    libharfbuzz0b \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    fonts-arphic-ukai \
    fonts-arphic-uming \
    fonts-liberation2 \
    libxinerama1 \
    libgl1 \
    libdbus-1-3 \
    libcups2 \
    libxrender1 \
    libxext6 \
    libsm6 \
    libice6 \
    libglib2.0-0 \
    libreoffice-impress \
    mupdf-tools \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY docker/fontconfig/99-pptx2svg-fonts.conf /etc/fonts/conf.d/99-pptx2svg-fonts.conf
COPY docker/entrypoint.sh /usr/local/bin/pptx2svg-entrypoint.sh
COPY fonts /usr/local/share/fonts/custom
COPY app ./app

RUN chmod +x /usr/local/bin/pptx2svg-entrypoint.sh && fc-cache -fv

EXPOSE 8321

ENTRYPOINT ["/usr/local/bin/pptx2svg-entrypoint.sh"]
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8321"]
