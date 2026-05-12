<!-- input: HarmonyOS 字体 CSS 地址与下载/转换脚本 -->
<!-- output: 容器可加载的本地 TTF 字体分片清单 -->
<!-- pos: 2pptxsvg 字体资产目录说明 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# fonts/

> 一旦我所属的文件夹有所变化，请更新我

本目录只放容器渲染需要的本地字体文件（TTF）。  
当前来源是 `HarmonyOS Sans SC` 的 Bold/Regular 两套 `result.css`。  
目标是让 `/v1/rasterize` 与 PPT/PPTX 导出在无外网时仍能命中中文字体。

## 当前字体概览

- Family: `HarmonyOS Sans SC`
- Subfamily `Bold` (`weight=700`): `165` 个 TTF 分片
- Subfamily `Regular` (`weight=400`): `164` 个 TTF 分片

## 来源地址

- [HarmonyOS_Sans_SC_Bold/result.css](https://content-sandunppt.oss-cn-guangzhou.aliyuncs.com/fonts/HarmonyOS_Sans_SC_Bold/result.css)
- [HarmonyOS_Sans_SC_Regular/result.css](https://content-sandunppt.oss-cn-guangzhou.aliyuncs.com/fonts/HarmonyOS_Sans_SC_Regular/result.css)

## 目录结构

- `HarmonyOS_Sans_SC_Bold/`: `result.css` + Bold 分片 `.ttf`
- `HarmonyOS_Sans_SC_Regular/`: `result.css` + Regular 分片 `.ttf`
