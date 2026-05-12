<!-- input: 维护者在本目录内新增/修改文件 -->
<!-- output: pptx2svg 子系统的极简架构与文件职责说明 -->
<!-- pos: 2pptxsvg 中的 PPT/PPTX->SVG 子系统总览文档 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# pptx2svg_api/

> 一旦我所属的文件夹有所变化，请更新我

本目录保留原 `pptx2svg` 的下载、解析、导出、文本提取逻辑。  
统一服务由根目录 `api_server.py` 调用这里的 converter/config。  
如果单独跑这套逻辑，仍可直接复用 `app/main.py`。

## 文件结构

| 名称 | 地位 | 功能 |
|------|------|------|
| `__init__.py` | 包标记 | 让 `pptx2svg_api` 可被 Python import |
| `app/` | 子包 | 保存 `config.py`、`converter.py`、`main.py` 与 `README.md`；`converter.py` 负责逐页 PNG ready 回调和原稿资产解析 |
