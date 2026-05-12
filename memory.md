<!-- input: 2pptxsvg 目录中的关键改动结论、线上问题复盘与协作偏好 -->
<!-- output: 供后续会话快速继承的项目记忆清单 -->
<!-- pos: 2pptxsvg 项目长期记忆（跨会话） -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# memory

## 2026-05-11

- 2026-05-12：PPT 美化新增遮挡识图图：`converter.py` 会把大块内嵌图片/文件名像截图的图片标记为 `role=screenshot`、`ocr_ignore=true`，并生成 `slide-xxx-vision.png` 灰块遮挡图；`beautify-pages` 新增 `/pages/{page}/vision-image` 下载口，原始 PNG 仍用于预览/Nano，vision PNG 只给 Go 识图链路使用。验证：`python -m py_compile api_server.py pptx2svg_api\app\converter.py` 通过。
- 2026-05-12：PPT 美化新增逐页 ready 解析任务 `/api/v1/pptx/jobs/beautify-pages`：LibreOffice 先导出 PDF，MuPDF 按页渲染 PNG，每页 ready 后立即写入 job `pages[]` 并暴露单页图片/资产下载接口；job 状态会带真实总页数，供 Go 端边转图边识图。验证：`python -m py_compile api_server.py pptx2svg_api\app\converter.py` 通过；重启容器后用 `37666ezachd8b.pptx` 冒烟返回 18 页、12 个图片资产、第一页 PNG 下载 200。
- `analyze_ppt_url_for_beautify` 已从 SVG 美化解析改为 PNG 识图解析：复用 LibreOffice 导出 PDF，再用 MuPDF `mutool draw -r 144 -F png` 生成 `slide-001.png` 等整页截图；`metadata.json` 新增 `render_mode=png_vision`、`image_filename/image_width/image_height`。
- PPT 美化仍保留逐页图片资产：`converter.py` 从 PPTX OOXML 的 `p:pic/a:blip` 关系直接提取内嵌图片，写入 ZIP 的 `assets/slide-xxx/image-xxx.*`，metadata.images 保留坐标、尺寸、role、filename，供 Go 上传和前端原稿资产开关使用。
- 2026-05-12 修正：PPT 美化原稿资产继续从 PPT/PPTX 文件本身解析，不依赖 SVG；`converter.py` 除普通 `p:pic` 外补充识别 slide 背景图和形状图片填充里的 `a:blip`，减少 PNG 识图链路下“原稿资产识别不出”的情况。
- 2026-05-12 用 `37666ezachd8b.pptx` 验证运行态：重启 `2pptxsvg-api` 后 `/api/v1/analyze/pptx-beautify` 返回 18 页、12 个位图资产、0 个 parse error；如果只看到第 11 页 4 张图，说明容器还在跑旧 Python 进程，需要重启服务。
- 按 sandun 要求新建 `2pptxsvg/`，保留原 `svg2ppt/` 与 `pptx2svg/` 目录不动，把两边能力合并到一个 FastAPI 服务。
- 新服务只保留 `2222` 端口；根入口是 `api_server.py`，同时暴露原 `svg2ppt` 的 `/v1/*` 接口和原 `pptx2svg` 的 `/api/v1/*` 接口。
- `pptx2svg_api/app/` 保存 PPT/PPTX 下载、LibreOffice 导出、拆页、beautify metadata 和文本提取逻辑；`svg_to_editable_pptx.py` 保存 SVG 转可编辑 PPTX 与栅格化逻辑。
- Docker 镜像合并安装 LibreOffice/MuPDF 与 cairo/pango/skia 相关依赖，`docker-compose.yml` 映射 `2222:2222`，服务名/容器名为 `2pptxsvg-api`。
- Go 后端 `PPTXService` 需要从旧 `http://127.0.0.1:8321/api/v1/analyze/pptx-beautify` 改为 `http://127.0.0.1:2222/api/v1/analyze/pptx-beautify`。
- 2026-05-11 补充验证：本机临时服务已跑通 `/v1/rasterize`、`/v1/convert`、`/v1/jobs/convert`、任务查询与下载；`/healthz` 在宿主机上报 `soffice` 缺失，属于本机环境限制。Docker 直接构建仍受 `python:3.12-slim` 拉取失败影响。
- 2026-05-11 继续优化：PPT/PPTX 三个接口改成共用队列，新增 `/api/v1/pptx/jobs/convert|beautify|extract` 异步任务入口；SVG->PPTX 同步 `convert` 也改为队列执行，并加 `SVG2PPTX_CONVERT_TIMEOUT_SECONDS` 总时长保护。
