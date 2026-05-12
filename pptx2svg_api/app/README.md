<!-- input: 维护者在本目录内新增/修改文件 -->
<!-- output: PPT/PPTX->SVG/PNG app子包的极简架构与文件职责说明 -->
<!-- pos: pptx2svg_api/app 子包总览文档 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

一旦我所属的文件夹有所变化，请更新我。

本目录是 PPT/PPTX 转 SVG/PNG 的核心子包。  
`config.py` 负责运行参数，`converter.py` 负责下载、导出、拆页、PNG 美化 metadata、遮挡截图区域的 vision PNG、逐页 ready PNG 回调、从 PPT/PPTX OOXML 解析原稿图片资产和文本提取。  
`main.py` 保留原始 FastAPI 路由，便于单独复用和对照。

2026-05-12：`converter.py` 对大块内嵌图片/文件名像截图的图片标记 `role=screenshot` 与 `ocr_ignore=true`，额外生成 `slide-xxx-vision.png` 灰块遮挡图；原始 PNG 仍保留给预览和 Nano 参考，vision PNG 只给美化识图使用。

2026-05-12：`converter.py` 新增 `analyze_ppt_url_for_beautify_progressive`，PDF 总页数确定后逐页渲染 PNG，并在每页 ready 时回调页面 metadata，供 Go 后端边转图边启动识图。

2026-05-12：`converter.py` 的 beautify 图片资产解析继续从 PPT/PPTX 文件本身读取，除 `p:pic` 外补充识别 slide 背景图和形状图片填充，不依赖 SVG。

2026-05-11：`config.py`、`converter.py`、`main.py` 的文件说明已统一放在文件开头，业务逻辑不变。

- `__init__.py`：包标记。
- `config.py`：运行参数配置，读取工作目录、命令路径、固定字体、PNG 渲染 DPI 和限制。
- `converter.py`：下载 PPT/PPTX、统一字体、导出 SVG、拆页、PPT 美化 PNG 截图、截图区域遮挡 vision PNG、逐页 ready PNG 回调、从 PPT/PPTX OOXML 提取原稿图片资产（含背景图/形状图片填充）、文本提取与 beautify metadata 生成。
- `main.py`：原始 FastAPI 路由，提供健康检查、转换、分析、文本提取和下载接口。
