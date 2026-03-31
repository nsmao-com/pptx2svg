FROM python:3.12-slim

ARG LIBREOFFICE_VERSION=26.2.1
ARG LIBREOFFICE_SERIES=26.2
ARG LIBREOFFICE_TARBALL=LibreOffice_${LIBREOFFICE_VERSION}_Linux_x86-64_deb.tar.gz
ARG LIBREOFFICE_DOWNLOAD_URL=https://download.documentfoundation.org/libreoffice/stable/${LIBREOFFICE_VERSION}/deb/x86_64/${LIBREOFFICE_TARBALL}

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    WORK_ROOT=/tmp/ppt-to-svg \
    APP_PORT=8321 \
    DOWNLOAD_TIMEOUT_SECONDS=120 \
    COMMAND_TIMEOUT_SECONDS=240 \
    MAX_DOWNLOAD_MB=100 \
    LIBREOFFICE_START_TIMEOUT_SECONDS=45 \
    LIBREOFFICE_PROGRAM_DIR=/opt/libreoffice${LIBREOFFICE_SERIES}/program

RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    curl \
    tar \
    xz-utils \
    fontconfig \
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
    default-jre-headless \
    && rm -rf /var/lib/apt/lists/*

RUN curl -fsSL "$LIBREOFFICE_DOWNLOAD_URL" -o /tmp/libreoffice.tar.gz \
    && mkdir -p /tmp/libreoffice \
    && tar -xzf /tmp/libreoffice.tar.gz -C /tmp/libreoffice --strip-components=1 \
    && apt-get update \
    && apt-get install -y --no-install-recommends /tmp/libreoffice/DEBS/*.deb \
    && rm -rf /var/lib/apt/lists/* /tmp/libreoffice /tmp/libreoffice.tar.gz

WORKDIR /app

COPY python_api/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY docker/fontconfig/99-pptx2svg-fonts.conf /etc/fonts/conf.d/99-pptx2svg-fonts.conf
COPY docker/entrypoint.sh /usr/local/bin/pptx2svg-entrypoint.sh
COPY fonts /usr/local/share/fonts/custom
COPY python_api/app ./app

RUN chmod +x /usr/local/bin/pptx2svg-entrypoint.sh && fc-cache -fv

EXPOSE 8321

ENTRYPOINT ["/usr/local/bin/pptx2svg-entrypoint.sh"]
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8321"]
