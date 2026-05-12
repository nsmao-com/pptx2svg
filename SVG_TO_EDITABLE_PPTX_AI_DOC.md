<!-- input: SVG转PPTX核心代码与AI集成约定 -->
<!-- output: 可编辑PPTX转换能力边界和调用规范 -->
<!-- pos: 2pptxsvg 中 SVG->PPTX 的设计说明文档 -->
<!-- 一旦我被更新，务必更新我的开头注释，以及所属的文件夹README。 -->

# SVG to Editable PPTX - AI Integration Spec

## 1. Purpose

`svg_to_editable_pptx.py` converts SVG into editable PowerPoint shapes.

Primary goal:
- Keep front-layer elements editable in PPT (`p:sp`, `p:cxnSp`, text boxes, pictures).

Secondary goal:
- Optionally rasterize detected background group into one full-slide image for better visual fidelity.

---

## 2. Entry Point

Script:
- `D:\svg2pptx\svg_to_editable_pptx.py`

CLI:
```bash
python svg_to_editable_pptx.py <input> [-o <output.pptx>] [--no-bg-image]
```

Arguments:
- `input`: one `.svg` file or a directory containing `.svg` files.
- `-o, --output`: target `.pptx` file path.
- `--no-bg-image`: disable background rasterization and keep all elements editable vectors.

Exit behavior:
- Returns `0` when script completes (even if individual slide conversion logs `FAIL` lines).
- Returns non-zero only for hard input errors (missing path, no SVG files).

---

## 3. Dependencies

Required:
- `python-pptx`
- `svgpathtools`
- `numpy`
- `Pillow`

Optional raster backends for background image:
- Preferred: `cairosvg`
- Fallback: `skia-python`

Backend order:
1. CairoSVG
2. Skia
3. No background rasterization (warning printed)

---

## 4. High-Level Pipeline

Per SVG file:
1. Parse SVG root and viewBox.
2. Parse CSS `<style>` blocks (simple selectors only: tag, `.class`, `#id`).
3. Parse gradients (`linearGradient`, `radialGradient` approximated).
4. Create PPT slide with dimensions from first SVG.
5. If enabled, detect background group and rasterize it to one full-slide PNG.
6. Traverse SVG tree and rebuild elements as editable PPT objects.
7. Save PPTX.

---

## 5. Supported SVG Elements

Currently mapped:
- `rect`
- `circle`
- `ellipse`
- `line`
- `polyline`
- `polygon`
- `path`
- `text` / `tspan` (basic line handling)
- `image` (`http/https`, local file, `data:` URI)
- container: `svg`, `g`, `a`

Ignored/non-rendered containers:
- `defs`, `style`, `script`, `metadata`, `title`, `desc`, `clipPath`, `mask`, `filter`

---

## 6. Style and Color Rules

### 6.1 Style precedence
For each element:
1. Inherited parent style
2. CSS rule style (tag/class/id)
3. Inline `style` attribute
4. Direct SVG attributes (for known inherited keys)

### 6.2 Color parsing
Supported paint formats:
- Hex and named colors (via PIL)
- `rgb(...)`, `rgba(...)`
- `hsl(...)`, `hsla(...)`
- `transparent`
- `currentColor`

### 6.3 Opacity composition
Effective alpha is multiplicative:
- `parent_opacity * element_opacity * fill-opacity` for fill
- `parent_opacity * element_opacity * stroke-opacity` for stroke

### 6.4 Gradient handling
- `linearGradient`: mapped to PPT gradient with 2 effective stops (first + last).
- `radialGradient`: approximated as linear direction using radial center/radius projection.
- Multi-stop exact reproduction is not implemented.

---

## 7. Geometry and Transform

- Accumulates SVG transforms as matrices through the tree.
- Converts path curves/arcs to sampled polylines before building freeform shapes.
- Uses `build_freeform(...).convert_to_shape(...)` for editable geometry.
- `line` maps to PPT connector (`MSO_CONNECTOR.STRAIGHT`).

Notes:
- Complex path fidelity depends on sampling density.
- Shear transforms on text/images are not fully preserved as native transform; rotation is applied when safe.

---

## 8. Background Rasterization Contract

Background group detection:
- First pass exact IDs: `bg`, `background`, `bg_layer`, `background_layer`
- Second pass fuzzy: id contains `bg` or `background`

Rasterized output:
- One full-slide picture inserted at `(0,0)` with slide size.
- Detected background group is skipped from vector reconstruction.

Important implementation detail:
- Background sub-SVG is serialized with default SVG namespace (no `svg:` prefixed tags) to avoid transparent render issues in Skia.

---

## 9. Output Guarantees

When conversion succeeds:
- PPT slide size matches SVG viewBox of first input file.
- Non-background supported elements become editable PPT objects.
- Original SVG file is not embedded as an SVG media part in this converter.

Typical slide content mix:
- Background image (`p:pic`) if bg rasterization enabled and backend available.
- Editable vector/text/connectors (`p:sp`, `p:cxnSp`).
- Additional pictures for original SVG `<image>` elements.
- Native shapes explicitly clear PowerPoint theme default effects to avoid accidental shadow on clean cards.
- Text runs write the resolved `font-family` into `a:latin`, `a:ea`, and `a:cs` so PowerPoint can retain actual font names for Latin and East Asian text when the font is installed locally.

---

## 10. Known Limitations

- No full support for `clipPath`, `mask`, `filter`, blend modes, `textPath`.
- Path curves are approximated by polylines.
- Text layout is approximate (baseline, kerning, advanced typography differ from browser rendering).
- Web fonts referenced by SVG `@import` / `@font-face` are not embedded into `.pptx`; only the resolved font family name is written into the PPT text run.
- Gradient fidelity is simplified.
- Background detection is ID-based heuristic (not semantic scene analysis).

---

## 11. Recommended Usage Patterns

Single file:
```bash
python svg_to_editable_pptx.py 1.svg -o 1.editable.pptx
```

Directory batch:
```bash
python svg_to_editable_pptx.py . -o batch.editable.pptx
```

Disable background rasterization:
```bash
python svg_to_editable_pptx.py 1.svg -o 1.no_bg_image.pptx --no-bg-image
```

If output file is open in PowerPoint:
- close the file first or write to a new output filename.

---

## 12. Extension Targets (for future AI edits)

High-value next improvements:
1. Add explicit `--bg-id <id>` to avoid heuristic misses.
2. Add `clipPath` approximation for common rectangular clips.
3. Improve text layout with tspans/x/dy support and better bounding estimation.
4. Add optional higher path sampling quality flag.
5. Add per-slide conversion report (counts, skipped tags, unsupported style warnings).
