package com.pptx2svg.renderer;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import org.apache.batik.ext.awt.image.codec.png.PNGImageWriter;
import org.apache.batik.ext.awt.image.spi.ImageWriterRegistry;
import org.apache.poi.xslf.util.PPTX2PNG;

public final class PresentationToSvg {
    private PresentationToSvg() {
    }

    public static void main(String[] args) throws Exception {
        Arguments parsed = Arguments.parse(args);
        Files.createDirectories(parsed.outputDir());
        ensureImageWriters();
        PPTX2PNG.main(buildPoiArgs(parsed));
        System.out.println("ok");
    }

    private static void ensureImageWriters() {
        ImageWriterRegistry registry = ImageWriterRegistry.getInstance();
        if (registry.getWriterFor("image/png") == null) {
            registry.register(new PNGImageWriter());
        }
    }

    private static String[] buildPoiArgs(Arguments parsed) {
        List<String> poiArgs = new ArrayList<>();
        poiArgs.add("-format");
        poiArgs.add("svg");
        poiArgs.add("-outdir");
        poiArgs.add(parsed.outputDir().toString());
        poiArgs.add("-outpat");
        poiArgs.add("slide-${slideno}.${format}");
        if (parsed.textAsShapes()) {
            poiArgs.add("-textAsShapes");
        }
        poiArgs.add(parsed.input().toString());
        return poiArgs.toArray(String[]::new);
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
