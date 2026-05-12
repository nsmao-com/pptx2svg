<!-- input: 维护者在本目录内新增/修改文件 -->
<!-- output: 2pptxsvg 统一服务的极简架构与文件职责说明 -->
<!-- pos: 2pptxsvg 子系统总览文档 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# 2pptxsvg/ - PPT/SVG 双向转换统一服务

> 一旦我所属的文件夹有所变化，请更新我

一个目录只启动一个服务，端口固定 `2222`。  
合并原 `svg2ppt` 与 `pptx2svg` 能力：SVG 转可编辑 PPTX/PNG，PPT/PPTX 转逐页 SVG/文本，PPT 美化解析转逐页 PNG 截图和原稿图片 metadata，并支持逐页 ready 让上游边转图边识图。  
根入口是 `api_server.py`，PPTX 解析核心在 `pptx2svg_api/app/`，SVG 转 PPTX 核心在 `svg_to_editable_pptx.py`。

## 文件结构

| 名称 | 地位 | 功能 |
|------|------|------|
| `api_server.py` | 统一入口 | 暴露 `2222` 上的所有 HTTP 接口：SVG->PPTX、SVG->PNG、PPT/PPTX->SVG、beautify PNG 解析、逐页 ready 美化任务、文本提取与下载 |
| `svg_to_editable_pptx.py` | SVG转PPT核心 | 把 SVG 元素重建为可编辑 PPTX 形状，并提供 SVG 栅格化能力 |
| `pptx2svg_api/` | PPT解析子系统 | 保留 PPT/PPTX 下载、LibreOffice 导出、拆页、metadata 与文本提取逻辑 |
| `fonts/` | 字体资产 | 本地 HarmonyOS Sans SC 分片字体，供栅格化和 LibreOffice 字体替代使用 |
| `docker/` | 容器配置 | fontconfig 字体替代规则 |
| `Dockerfile` | 镜像构建 | 构建含 LibreOffice/MuPDF/cairo/pango/skia 的统一镜像 |
| `docker-compose.yml` | 编排入口 | 一键启动 `2pptxsvg-api`，只映射 `2222:2222` |
| `docker-entrypoint.sh` | 启动脚本 | 启动前复制字体并刷新 fontconfig 缓存 |
| `requirements.txt` | Python依赖 | FastAPI、PPTX、SVG解析、栅格化与 HTTP 下载依赖 |
| `SVG_TO_EDITABLE_PPTX_AI_DOC.md` | 设计文档 | SVG 转可编辑 PPTX 的能力边界说明 |
| `.dockerignore` | 构建过滤 | 排除运行产物、虚拟环境和日志 |
| `memory.md` | 协作记忆 | 记录合并后的关键决策和运维注意事项 |

## API

### 通用

- `GET /healthz`：统一健康检查，同时校验 PPTX 解析依赖。

### SVG -> PPTX / PNG

- `POST /v1/jobs/convert`：异步转换，`svgs[]` -> `.pptx`。
- `GET /v1/jobs/{job_id}`：查询异步任务。
- `GET /v1/jobs/{job_id}/download`：下载异步产物。
- `DELETE /v1/jobs/{job_id}`：删除已完成/失败任务。
- `POST /v1/convert`：同步转换，直接返回 `.pptx`。
- `POST /v1/rasterize`：同步栅格化单个 SVG，返回 `image/png`。

### PPT/PPTX -> SVG / PNG metadata / 文本

- `POST /api/v1/convert/ppt-to-svg`：远程 PPT/PPTX 链接 -> 逐页 SVG ZIP。
- `POST /api/v1/analyze/pptx-beautify`：远程 PPT/PPTX 链接 -> 逐页 PNG ZIP + `metadata.json`，metadata 保留整页截图、遮挡内嵌截图区域的 vision PNG、抽取文字、从 PPT/PPTX 文件解析出的原稿图片资产（含背景图/形状图片填充）和图表数据，供美化识图使用。
- `POST /api/v1/extract/ppt-text`：远程 PPT/PPTX 链接 -> 按页 TXT 文本。
- `POST /api/v1/pptx/jobs/convert`、`/beautify`、`/beautify-pages`、`/extract`：PPT/PPTX 异步任务入口，和同步接口共用同一排队并发；`beautify-pages` 会在每页 PNG 渲染完成后立即把该页加入 `pages[]`。
- `GET /api/v1/pptx/jobs/{job_id}`、`/download`、`/metadata`、`DELETE /api/v1/pptx/jobs/{job_id}`：异步任务状态、下载、元数据与清理。
- `GET /api/v1/pptx/jobs/{job_id}/pages/{page_index}/image`：下载 `beautify-pages` 已 ready 的原始单页 PNG。
- `GET /api/v1/pptx/jobs/{job_id}/pages/{page_index}/vision-image`：下载 `beautify-pages` 已 ready 的遮挡识图 PNG，用于避免模型读取内嵌截图内容。
- `GET /api/v1/pptx/jobs/{job_id}/assets/{asset_path}`：下载 `beautify-pages` 已解析出的原稿图片资产。
- `GET /downloads/{filename}`：下载 `url=true` 模式保存的 ZIP/TXT/metadata。

## 本地/服务器启动

```bash
docker compose up -d --build
```

服务地址：

```text
http://127.0.0.1:2222
```

当前本机开发态容器挂载源码但没有启用 Python 自动重载，修改 `converter.py`、`api_server.py` 等文件后需要重启 `2pptxsvg-api` 才会生效。

Go 后端 PPTX 美化解析应调用：

```text
http://127.0.0.1:2222/api/v1/analyze/pptx-beautify
```

## 环境变量

- `JOB_WORKDIR`：SVG->PPTX 任务目录，默认 `/data/jobs`
- `PPTX_JOB_WORKDIR`：PPT/PPTX 异步任务产物目录，默认跟随 `JOB_WORKDIR`
- `MAX_CONCURRENT_JOBS`：SVG->PPTX 并发数
- `PPTX_MAX_CONCURRENT_JOBS`：PPT/PPTX->SVG 并发数，默认跟随 `MAX_CONCURRENT_JOBS`
- `JOB_TTL_SECONDS`：任务保留秒数
- `PPTX_JOB_TTL_SECONDS`：PPT/PPTX 任务保留秒数，默认跟随 `JOB_TTL_SECONDS`
- `WORK_ROOT`：PPT/PPTX->SVG 临时目录，默认 `/tmp/ppt-to-svg`
- `DOWNLOAD_TIMEOUT_SECONDS`：下载超时秒数
- `COMMAND_TIMEOUT_SECONDS`：LibreOffice/MuPDF 命令超时秒数
- `MAX_DOWNLOAD_MB`：PPT/PPTX 最大下载体积
- `PPTX_PNG_RENDER_DPI`：PPT 美化整页 PNG 渲染 DPI，默认 `144`
- `SVG2PPTX_CONVERT_TIMEOUT_SECONDS`：SVG->PPTX 单次转换总时长上限，默认 `900`
- `UPLOAD_SVG_IMAGES_TO_OSS` 与 `ALIYUN_OSS_*`：可选 OSS 外链化配置
