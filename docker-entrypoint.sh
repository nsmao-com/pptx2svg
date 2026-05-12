#!/bin/sh
# input: 容器环境变量与可选的挂载字体目录
# output: 刷新字体缓存后启动 2pptxsvg API 进程
# pos: 2pptxsvg 容器运行时字体初始化入口
# 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。

set -eu

CUSTOM_FONT_DIR="${SVG2PPTX_CUSTOM_FONT_DIR:-/app/fonts}"
TARGET_FONT_DIR="/usr/local/share/fonts/custom"

if [ -d "${CUSTOM_FONT_DIR}" ]; then
  mkdir -p "${TARGET_FONT_DIR}"
  FONT_COUNT="$(find "${CUSTOM_FONT_DIR}" -maxdepth 6 -type f \( -iname '*.ttf' -o -iname '*.otf' -o -iname '*.ttc' \) | wc -l | tr -d ' ')"
  if [ "${FONT_COUNT}" -gt 0 ]; then
    find "${CUSTOM_FONT_DIR}" -maxdepth 6 -type f \( -iname '*.ttf' -o -iname '*.otf' -o -iname '*.ttc' \) -exec cp -f {} "${TARGET_FONT_DIR}/" \;
    if command -v fc-cache >/dev/null 2>&1; then
      fc-cache -f >/dev/null 2>&1 || true
    fi
    echo "[2pptxsvg] loaded custom fonts: ${FONT_COUNT}"
  fi
fi

exec "$@"
