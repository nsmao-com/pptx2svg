# input: 当前目录源码、requirements.txt、字体文件与fontconfig规则
# output: 可运行2pptxsvg统一API的Docker镜像
# pos: 2pptxsvg生产部署容器构建文件
# 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    WORK_ROOT=/tmp/ppt-to-svg \
    APP_PORT=2222 \
    DOWNLOAD_TIMEOUT_SECONDS=120 \
    COMMAND_TIMEOUT_SECONDS=240 \
    PPTX_PNG_RENDER_DPI=144 \
    MAX_DOWNLOAD_MB=100 \
    SVG2PPTX_CONVERT_TIMEOUT_SECONDS=900 \
    LIBREOFFICE_COMMAND=soffice \
    MUPDF_COMMAND=mutool \
    SVG2PPTX_CUSTOM_FONT_DIR=/app/fonts

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates \
    fontconfig \
    libfontconfig1 \
    libfreetype6 \
    libharfbuzz0b \
    libcairo2 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf-2.0-0 \
    shared-mime-info \
    fonts-dejavu-core \
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

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY docker/fontconfig/99-pptx2svg-fonts.conf /etc/fonts/conf.d/99-pptx2svg-fonts.conf
COPY . ./

RUN mkdir -p /data/jobs /usr/local/share/fonts/custom /tmp/ppt-to-svg/downloads \
    && sed -i 's/\r$//' /app/docker-entrypoint.sh \
    && chmod +x /app/docker-entrypoint.sh \
    && fc-cache -f

ENV JOB_WORKDIR=/data/jobs \
    MAX_CONCURRENT_JOBS=4 \
    PPTX_MAX_CONCURRENT_JOBS=2 \
    PPTX_JOB_TTL_SECONDS=3600 \
    JOB_TTL_SECONDS=3600

EXPOSE 2222

ENTRYPOINT ["/app/docker-entrypoint.sh"]
CMD ["uvicorn", "api_server:app", "--host", "0.0.0.0", "--port", "2222", "--workers", "1"]
