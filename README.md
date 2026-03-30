# pptx2svg

一个可直接部署到 Docker 的 PPT/PPTX 转 SVG API 服务。

## 功能

- 传入远程 PPT/PPTX 链接
- 服务端自动下载文件
- 使用 LibreOffice 无头转换为 PDF
- 使用 `pdftocairo` 将每页 PDF 转为 SVG
- 返回按页切分后的 SVG ZIP 压缩包

接口实现位于 [python_api](./python_api)。

## API

### 健康检查

`GET /healthz`

### PPT 转 SVG

`POST /api/v1/convert/ppt-to-svg`

请求体：

```json
{
  "ppt_url": "https://example.com/demo.pptx"
}
```

成功后返回 `application/zip`。

## 直接从 GitHub 构建 Docker

仓库推送到 GitHub 后，可以直接从仓库远程构建：

```bash
docker build -t pptx2svg-api https://github.com/nsmao-com/pptx2svg.git
```

运行：

```bash
docker run --rm -p 8000:8000 pptx2svg-api
```

## 本地构建 Docker

```bash
docker build -t pptx2svg-api .
docker run --rm -p 8000:8000 pptx2svg-api
```

## curl 调用示例

```bash
curl -X POST "http://127.0.0.1:8000/api/v1/convert/ppt-to-svg" \
  -H "Content-Type: application/json" \
  -d "{\"ppt_url\":\"https://example.com/demo.pptx\"}" \
  --output slides.zip
```

## 环境变量

- `WORK_ROOT`: 临时文件目录，默认 `/tmp/ppt-to-svg`
- `DOWNLOAD_TIMEOUT_SECONDS`: 下载超时秒数，默认 `120`
- `COMMAND_TIMEOUT_SECONDS`: 转换命令超时秒数，默认 `240`
- `MAX_DOWNLOAD_MB`: 最大下载体积，默认 `100`
