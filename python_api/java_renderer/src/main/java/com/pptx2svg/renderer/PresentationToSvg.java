package com.pptx2svg.renderer;

import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.lang.reflect.Constructor;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.batik.dom.GenericDOMImplementation;
import org.apache.batik.ext.awt.image.codec.png.PNGImageWriter;
import org.apache.batik.ext.awt.image.spi.ImageWriterRegistry;
import org.apache.poi.sl.draw.DrawFactory;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.xslf.draw.SVGPOIGraphics2D;
import org.apache.poi.xslf.usermodel.XSLFDiagram;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFShapeContainer;
import org.apache.poi.xslf.usermodel.XSLFSheet;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGroupShape;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;

public final class PresentationToSvg {
    private static final String SVG_NS = "http://www.w3.org/2000/svg";
    private static final String PRESENTATIONML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static final String DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private PresentationToSvg() {
    }

    public static void main(String[] args) throws Exception {
        Arguments parsed = Arguments.parse(args);
        Files.createDirectories(parsed.outputDir());
        ensureImageWriters();

        try (SlideShow<?, ?> slideShow = SlideShowFactory.create(parsed.input().toFile(), null, true)) {
            List<? extends Slide<?, ?>> slides = slideShow.getSlides();
            if (slides.isEmpty()) {
                throw new IllegalStateException("Presentation contains no slides.");
            }

            Dimension pageSize = slideShow.getPageSize();
            int index = 1;
            for (Slide<?, ?> slide : slides) {
                writeSlideSvg(slide, pageSize, parsed.outputDir().resolve(formatSlideName(index++)), parsed.textAsShapes());
            }
        }

        System.out.println("ok");
    }

    private static void ensureImageWriters() {
        ImageWriterRegistry registry = ImageWriterRegistry.getInstance();
        if (registry.getWriterFor("image/png") == null) {
            registry.register(new PNGImageWriter());
        }
    }

    private static void writeSlideSvg(
        Slide<?, ?> slide,
        Dimension pageSize,
        Path outputPath,
        boolean textAsShapes
    ) throws Exception {
        DOMImplementation domImplementation = GenericDOMImplementation.getDOMImplementation();
        Document document = domImplementation.createDocument(SVG_NS, "svg", null);
        SVGPOIGraphics2D graphics = new SVGPOIGraphics2D(document, textAsShapes);
        try {
            graphics.setSVGCanvasSize(pageSize);
            applyRenderingHints(graphics);
            slide.draw(graphics);
            if (slide instanceof XSLFSlide xslfSlide) {
                drawUnsupportedXslfShapes(xslfSlide, graphics);
            }
            try (Writer writer = new OutputStreamWriter(Files.newOutputStream(outputPath), StandardCharsets.UTF_8)) {
                graphics.stream(writer, true);
            }
        } finally {
            graphics.dispose();
        }
    }

    private static void applyRenderingHints(Graphics2D graphics) {
        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        graphics.setRenderingHint(RenderingHints.KEY_COLOR_RENDERING, RenderingHints.VALUE_COLOR_RENDER_QUALITY);
        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
        graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
        graphics.setRenderingHint(Drawable.CACHE_IMAGE_SOURCE, Boolean.TRUE);
    }

    private static void drawUnsupportedXslfShapes(XSLFShapeContainer container, Graphics2D graphics) {
        for (XSLFShape shape : container.getShapes()) {
            if (shape instanceof XSLFGraphicFrame frame && !frame.hasChart()) {
                XSLFGroupShape fallbackGroup = buildFallbackGroup(frame);
                if (fallbackGroup != null) {
                    drawShape(fallbackGroup, graphics);
                    continue;
                }

                XSLFPictureShape fallbackPicture = frame.getFallbackPicture();
                if (fallbackPicture != null) {
                    drawShape(fallbackPicture, graphics);
                    continue;
                }
            }

            if (shape instanceof XSLFDiagram diagram) {
                drawShape(diagram.getGroupShape(), graphics);
                continue;
            }

            if (shape instanceof XSLFGroupShape groupShape) {
                drawUnsupportedXslfShapes(groupShape, graphics);
            }
        }
    }

    private static XSLFGroupShape buildFallbackGroup(XSLFGraphicFrame frame) {
        try {
            XmlObject[] fallbackNodes = frame.getXmlObject().selectPath(
                "declare namespace p='" + PRESENTATIONML_NS + "'; " +
                "declare namespace mc='" + MC_NS + "' " +
                ".//mc:Fallback/*/*[self::p:sp or self::p:pic or self::p:grpSp or self::p:cxnSp]"
            );
            if (fallbackNodes.length == 0) {
                return null;
            }

            StringBuilder xml = new StringBuilder();
            xml.append("<p:grpSp xmlns:p=\"").append(PRESENTATIONML_NS).append("\" ")
                .append("xmlns:a=\"").append(DRAWINGML_NS).append("\">")
                .append("<p:nvGrpSpPr><p:cNvPr id=\"0\" name=\"fallback\"/>")
                .append("<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>")
                .append("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>")
                .append("<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");
            for (XmlObject fallbackNode : fallbackNodes) {
                xml.append(fallbackNode.xmlText());
            }
            xml.append("</p:grpSp>");

            CTGroupShape groupShapeXml = CTGroupShape.Factory.parse(xml.toString());
            XSLFGroupShape groupShape = instantiateGroupShape(groupShapeXml, frame.getSheet());
            Rectangle2D anchor = frame.getAnchor();
            groupShape.setAnchor(anchor);
            groupShape.setInteriorAnchor(new Rectangle2D.Double(0, 0, anchor.getWidth(), anchor.getHeight()));
            return groupShape;
        } catch (Exception ignored) {
            return null;
        }
    }

    private static XSLFGroupShape instantiateGroupShape(CTGroupShape groupShapeXml, XSLFSheet sheet) throws Exception {
        Constructor<XSLFGroupShape> constructor = XSLFGroupShape.class.getDeclaredConstructor(CTGroupShape.class, XSLFSheet.class);
        constructor.setAccessible(true);
        return constructor.newInstance(groupShapeXml, sheet);
    }

    private static void drawShape(XSLFShape shape, Graphics2D graphics) {
        if (shape == null) {
            return;
        }

        DrawFactory drawFactory = DrawFactory.getInstance(graphics);
        Drawable drawable = drawFactory.getDrawable(shape);
        if (drawable == null) {
            return;
        }

        AffineTransform originalTransform = graphics.getTransform();
        graphics.setRenderingHint(Drawable.GSAVE, true);
        try {
            drawable.applyTransform(graphics);
            drawable.draw(graphics);
        } finally {
            graphics.setTransform(originalTransform);
            graphics.setRenderingHint(Drawable.GRESTORE, true);
        }
    }

    private static String formatSlideName(int index) {
        return String.format(Locale.ROOT, "slide-%03d.svg", index);
    }

    private record Arguments(Path input, Path outputDir, boolean textAsShapes) {
        private static Arguments parse(String[] args) {
            Path input = null;
            Path outputDir = null;
            boolean textAsShapes = true;

            List<String> tokens = new ArrayList<>();
            for (String arg : args) {
                tokens.add(arg);
            }

            for (int index = 0; index < tokens.size(); index++) {
                String token = tokens.get(index);
                switch (token) {
                    case "--input" -> {
                        input = Path.of(requireValue(tokens, ++index, token));
                    }
                    case "--output-dir" -> {
                        outputDir = Path.of(requireValue(tokens, ++index, token));
                    }
                    case "--text-as-shapes" -> {
                        textAsShapes = Boolean.parseBoolean(requireValue(tokens, ++index, token));
                    }
                    default -> throw new IllegalArgumentException("Unknown argument: " + token);
                }
            }

            if (input == null) {
                throw new IllegalArgumentException("Missing required argument: --input");
            }
            if (outputDir == null) {
                throw new IllegalArgumentException("Missing required argument: --output-dir");
            }

            return new Arguments(input, outputDir, textAsShapes);
        }

        private static String requireValue(List<String> tokens, int index, String option) {
            if (index >= tokens.size()) {
                throw new IllegalArgumentException("Missing value for " + option);
            }
            return tokens.get(index);
        }
    }
}
