# PPT To SVG API

Python API，输入 PPT / PPTX 链接，输出按页拆分后的 SVG ZIP。

官网：<https://www.nsmao.com>

## 转换链路

当前实现：

`PPT/PPTX -> Apache POI + Batik -> SVG -> ZIP`

说明：

- 默认保留文字节点，不强制转 path
- 普通图片仍按图片方式输出
- 输出格式是每页一个 SVG，再统一打包为 ZIP

## 接口

### 健康检查

`GET /healthz`

### 转换接口

`POST /api/v1/convert/ppt-to-svg`

请求体：

```json
{
  "ppt_url": "https://example.com/demo.pptx",
  "url": false
}
```

返回：

- `url=false`：直接返回 ZIP
- `url=true`：返回下载链接 JSON

示例：

```json
{
  "filename": "demo-a1b2c3d4e5f6.zip",
  "url": "/downloads/demo-a1b2c3d4e5f6.zip"
}
```

## 本地运行

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8321
```

## Docker

构建：

```bash
docker build -t ppt-to-svg-api .
```

运行：

```bash
docker run --rm -p 8321:8321 ppt-to-svg-api
```

## curl 示例

```bash
curl -X POST "http://127.0.0.1:8321/api/v1/convert/ppt-to-svg" \
  -H "Content-Type: application/json" \
  -d "{\"ppt_url\":\"https://example.com/demo.pptx\"}" \
  --output slides.zip
```

## 环境变量

- `WORK_ROOT`：临时工作目录，默认 `/tmp/ppt-to-svg`
- `DOWNLOADS_SUBDIR`：下载目录名，默认 `downloads`
- `DOWNLOAD_TIMEOUT_SECONDS`：下载超时，默认 `120`
- `COMMAND_TIMEOUT_SECONDS`：命令超时，默认 `240`
- `MAX_DOWNLOAD_MB`：最大下载大小，默认 `100`
- `JAVA_COMMAND`：Java 命令，默认 `java`
- `JAVA_RENDERER_JAR`：渲染器 jar 路径
- `SVG_TEXT_AS_SHAPES`：是否把文字转成形状，默认 `false`

## 字体

- 可把自定义 `.ttf` / `.ttc` 放到 [fonts](./fonts)
- 镜像启动时会自动刷新字体缓存
- 特殊字体建议通过挂载方式提供，不建议直接提交到仓库
