# pptx2svg

一个可直接部署到 Linux Docker 的 PPT / PPTX 转 SVG 服务。

官网：<https://www.nsmao.com>

## 当前能力

- 输入远程 `.ppt` / `.pptx` 文件链接
- 服务端自动下载源文件
- 按页导出为独立 SVG
- 返回 ZIP 压缩包，或返回服务端下载链接
- 兼容 Linux Docker 部署

## 当前导出策略

当前主链路是：

`PPT/PPTX -> Apache POI + Batik -> 每页原生 SVG -> ZIP`

说明：

- 普通文字默认保留为 SVG `text`
- 普通形状、线条、图片优先走原生 SVG 渲染
- 部分复杂图形对象会走 POI 可用的 fallback 渲染
- 输出结果仍然是每页一个 `.svg`

## API

### 健康检查

`GET /healthz`

### PPT 转 SVG

`POST /api/v1/convert/ppt-to-svg`

请求体：

```json
{
  "ppt_url": "https://example.com/demo.pptx",
  "url": false
}
```

返回规则：

- `url=false` 或不传：直接返回 `application/zip`
- `url=true`：保存 ZIP 到服务端本地目录，并返回 JSON 下载地址

返回示例：

```json
{
  "filename": "demo-a1b2c3d4e5f6.zip",
  "url": "/downloads/demo-a1b2c3d4e5f6.zip"
}
```

## Docker

### 本地构建

```bash
docker build -t pptx2svg-api .
docker run --rm -p 8321:8321 pptx2svg-api
```

### 直接从 GitHub 构建

```bash
docker build -t pptx2svg-api https://github.com/nsmao-com/pptx2svg.git
docker run --rm -p 8321:8321 pptx2svg-api
```

### GHCR 镜像

```bash
docker pull ghcr.io/nsmao-com/pptx2svg:latest
docker run --rm -p 8321:8321 ghcr.io/nsmao-com/pptx2svg:latest
```

## 调用示例

```bash
curl -X POST "http://127.0.0.1:8321/api/v1/convert/ppt-to-svg" \
  -H "Content-Type: application/json" \
  -d "{\"ppt_url\":\"https://example.com/demo.pptx\"}" \
  --output slides.zip
```

## 环境变量

- `WORK_ROOT`：临时工作目录，默认 `/tmp/ppt-to-svg`
- `DOWNLOADS_SUBDIR`：下载目录名，默认 `downloads`
- `DOWNLOAD_TIMEOUT_SECONDS`：下载超时秒数，默认 `120`
- `COMMAND_TIMEOUT_SECONDS`：转换命令超时秒数，默认 `240`
- `MAX_DOWNLOAD_MB`：最大下载体积，默认 `100`
- `JAVA_COMMAND`：Java 可执行文件，默认 `java`
- `JAVA_RENDERER_JAR`：Java 渲染器 jar 路径
- `SVG_TEXT_AS_SHAPES`：是否将文字转成形状，默认 `false`

## 字体

- 构建镜像前，可把常用 `.ttf` / `.ttc` 放到 [fonts](./fonts)
- 容器内已包含常见中文字体和常见替代字体
- 如 PPT 使用特殊商业字体，建议自行挂载真实字体目录

挂载示例：

```bash
docker run -d \
  --name pptx2svg \
  -p 8321:8321 \
  -v /opt/pptx2svg/fonts:/usr/local/share/fonts/custom \
  ghcr.io/nsmao-com/pptx2svg:latest
```

## 代码位置

- API 入口：[`python_api/app`](./python_api/app)
- Java SVG 渲染器：[`python_api/java_renderer`](./python_api/java_renderer)

