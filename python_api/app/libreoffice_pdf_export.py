from __future__ import annotations

import argparse
import time
from pathlib import Path

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException


def make_property(name: str, value: object) -> PropertyValue:
    property_value = PropertyValue()
    property_value.Name = name
    property_value.Value = value
    return property_value


def to_file_url(path: Path) -> str:
    return uno.systemPathToFileUrl(str(path.resolve()))


def connect(host: str, port: int, timeout_seconds: int):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local_context,
    )
    connection_string = (
        f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
    )
    deadline = time.time() + timeout_seconds
    last_error: Exception | None = None

    while time.time() < deadline:
        try:
            return resolver.resolve(connection_string)
        except NoConnectException as exc:
            last_error = exc
            time.sleep(0.5)

    raise RuntimeError(
        f"Timed out connecting to LibreOffice listener on {host}:{port}."
    ) from last_error


def iter_shapes(container):
    if not hasattr(container, "getCount"):
        return

    for index in range(container.getCount()):
        shape = container.getByIndex(index)
        yield shape
        if shape.supportsService("com.sun.star.drawing.GroupShape"):
            yield from iter_shapes(shape)


def iter_shape_texts(shape):
    if shape.supportsService("com.sun.star.drawing.TableShape"):
        table = shape.getTable()
        rows = table.Rows.getCount()
        columns = table.Columns.getCount()
        for row in range(rows):
            for column in range(columns):
                cell = table.getCellByPosition(column, row)
                if cell is not None:
                    yield cell.Text

    if hasattr(shape, "createEnumeration"):
        yield shape
        return

    if hasattr(shape, "getText"):
        text = shape.getText()
        if text is not None:
            yield text


def disable_asian_character_spacing(text) -> int:
    if not hasattr(text, "createEnumeration"):
        return 0

    changed = 0
    paragraphs = text.createEnumeration()
    while paragraphs.hasMoreElements():
        paragraph = paragraphs.nextElement()
        try:
            paragraph.ParaIsCharacterDistance = False
            changed += 1
        except Exception:
            continue
    return changed


def normalize_document(document) -> int:
    changed = 0
    draw_pages = document.getDrawPages()
    for page_index in range(draw_pages.getCount()):
        draw_page = draw_pages.getByIndex(page_index)
        for shape in iter_shapes(draw_page):
            for text in iter_shape_texts(shape):
                changed += disable_asian_character_spacing(text)
    return changed


def get_svg_filter_name(document) -> str:
    if document.supportsService("com.sun.star.presentation.PresentationDocument"):
        return "impress_svg_Export"
    if document.supportsService("com.sun.star.drawing.DrawingDocument"):
        return "draw_svg_Export"
    raise RuntimeError("LibreOffice loaded an unsupported document type for SVG export.")


def make_filter_data(page_number: int):
    filter_data = (make_property("PageNumber", page_number),)
    return uno.Any("[]com.sun.star.beans.PropertyValue", filter_data)


def export_svg_page(document, output_path: Path, filter_name: str, page_number: int) -> None:
    export_properties = (
        make_property("FilterName", filter_name),
        make_property("FilterData", make_filter_data(page_number)),
        make_property("Overwrite", True),
    )
    document.storeToURL(to_file_url(output_path), export_properties)


def export_svg_pages(
    host: str,
    port: int,
    input_path: Path,
    output_dir: Path,
    timeout_seconds: int,
) -> tuple[int, int]:
    context = connect(host, port, timeout_seconds)
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop",
        context,
    )

    load_properties = (
        make_property("Hidden", True),
        make_property("ReadOnly", False),
    )
    document = desktop.loadComponentFromURL(
        to_file_url(input_path),
        "_blank",
        0,
        load_properties,
    )
    if document is None:
        raise RuntimeError("LibreOffice failed to load the presentation.")

    try:
        changed = normalize_document(document)
        output_dir.mkdir(parents=True, exist_ok=True)
        filter_name = get_svg_filter_name(document)
        draw_pages = document.getDrawPages()
        exported = draw_pages.getCount()

        for page_number in range(1, exported + 1):
            export_svg_page(
                document,
                output_dir / f"slide-{page_number:03d}.svg",
                filter_name,
                page_number,
            )

        return changed, exported
    finally:
        document.close(True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, required=True)
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--timeout", type=int, default=45)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    changed, exported = export_svg_pages(
        args.host,
        args.port,
        args.input,
        args.output_dir,
        args.timeout,
    )
    print(f"normalized_paragraphs={changed}")
    print(f"exported_svgs={exported}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
