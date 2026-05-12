<!-- input: fontconfig XML规则文件 -->
<!-- output: 容器内字体替代和字体族映射说明 -->
<!-- pos: 2pptxsvg 字体规则子目录说明 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# fontconfig/

> 一旦我所属的文件夹有所变化，请更新我

本目录只放 fontconfig XML 规则。  
规则用于让 LibreOffice 与 SVG 渲染命中中文字体替代。  
业务代码不放在这里。

## 文件结构

| 名称 | 地位 | 功能 |
|------|------|------|
| `99-pptx2svg-fonts.conf` | 字体映射规则 | 为常见中英文字体名配置替代字体 |
