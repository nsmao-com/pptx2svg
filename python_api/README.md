# PPT To SVG API

一个可直接放进 Docker 的 Python API，输入 PPT/PPTX 链接，输出按页切分后的 SVG 压缩包。

## 技术方案

转换链路：

1. 下载远程 `.ppt` / `.pptx`
2. 使用 `LibreOffice` 无头模式转成 PDF
3. 使用 `pdftocairo` 将 PDF 每一页转成一个 SVG
4. 打包成 ZIP 返回

## 接口

### 健康检查

`GET /healthz`

### 转换接口

`POST /api/v1/convert/ppt-to-svg`

请求体：

```json
{
  "ppt_url": "https://example.com/demo.pptx"
}
```

返回：

- `200 OK`
- `Content-Type: application/zip`
- 压缩包内文件名格式：`slide-001.svg`、`slide-002.svg`

## 本地运行

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8321
```

## Docker

构建镜像：

```bash
docker build -t ppt-to-svg-api .
```

运行容器：

```bash
docker run --rm -p 8321:8321 ppt-to-svg-api
```

## 调用示例

```bash
curl -X POST "http://127.0.0.1:8321/api/v1/convert/ppt-to-svg" \
  -H "Content-Type: application/json" \
  -d "{\"ppt_url\":\"https://example.com/demo.pptx\"}" \
  --output slides.zip
```

## 环境变量

- `WORK_ROOT`: 临时文件目录，默认 `/tmp/ppt-to-svg`
- `APP_PORT`: 服务端口，默认 `8321`
- `DOWNLOAD_TIMEOUT_SECONDS`: 下载超时，默认 `120`
- `COMMAND_TIMEOUT_SECONDS`: 转换命令超时，默认 `240`
- `MAX_DOWNLOAD_MB`: 最大下载体积，默认 `100`
- `PAGE_CONVERT_WORKERS`: PDF 按页转 SVG 的并发数，默认 `4`

## 注意事项

- 容器镜像体积会比较大，主要来自 `LibreOffice`
- 某些复杂动画、特效、字体在 SVG 中可能会有样式损失
- 当前实现返回 ZIP，适合直接下载或让上游服务继续处理


## 字体优化

- 构建镜像前，可把常用 .ttf / .ttc 字体放到 [fonts](./fonts)
- 镜像内已增加常见中文字体包和 Windows 常见字体的替代映射
- 如果 PPT 使用了特殊商用字体，仍建议把原字体文件放进 [fonts](./fonts) 后重新构建


## 挂载本地字体

如果你有真实的微软雅黑等字体文件，不要提交到公开仓库，建议直接在服务器挂载字体目录：

```bash
docker run -d --name pptx2svg -p 8321:8321 -v /opt/pptx2svg/fonts:/usr/local/share/fonts/custom ppt-to-svg-api
```

容器启动时会自动刷新字体缓存。
