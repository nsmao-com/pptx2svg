<!-- input: 容器入口脚本与 fontconfig 规则 -->
<!-- output: 2pptxsvg 镜像启动时的字体初始化与替代配置 -->
<!-- pos: 2pptxsvg 容器配置目录 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# docker/

> 一旦我所属的文件夹有所变化，请更新我

本目录只放容器运行时相关配置。  
根目录 `docker-entrypoint.sh` 负责复制字体并刷新缓存。  
`fontconfig/` 负责字体替代规则，不放业务代码。

## 文件结构

| 名称 | 地位 | 功能 |
|------|------|------|
| `README.md` | 目录说明 | 说明容器配置目录的边界 |
| `fontconfig/99-pptx2svg-fonts.conf` | 字体规则 | 为 LibreOffice / SVG 渲染提供字体替代与映射 |
