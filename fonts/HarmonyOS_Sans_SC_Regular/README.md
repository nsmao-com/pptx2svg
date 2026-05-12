<!-- input: HarmonyOS Sans SC Regular result.css 与 TTF分片 -->
<!-- output: 容器可加载的 Regular 中文字体资产说明 -->
<!-- pos: 2pptxsvg 字体资产子目录说明 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# HarmonyOS_Sans_SC_Regular/

> 一旦我所属的文件夹有所变化，请更新我

本目录保存 `HarmonyOS Sans SC` Regular 字重字体分片。  
`result.css` 记录字体分片映射。  
`*.ttf` 由容器启动脚本复制到系统字体目录。

## 文件结构

| 名称 | 地位 | 功能 |
|------|------|------|
| `result.css` | 字体CSS | 记录 Regular 字重分片映射 |
| `*.ttf` | 字体分片 | 供 fontconfig 与渲染后端加载 |
