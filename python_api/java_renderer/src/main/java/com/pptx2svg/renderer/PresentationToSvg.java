package com.pptx2svg.renderer;

import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.batik.dom.GenericDOMImplementation;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.xslf.draw.SVGPOIGraphics2D;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;

public final class PresentationToSvg {
    private static final String SVG_NS = "http://www.w3.org/2000/svg";

    private PresentationToSvg() {
    }

    public static void main(String[] args) throws Exception {
        Arguments parsed = Arguments.parse(args);
        Files.createDirectories(parsed.outputDir());

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
