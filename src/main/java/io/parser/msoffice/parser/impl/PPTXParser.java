package io.parser.msoffice.parser.impl;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.parser.msoffice.model.Page;
import io.parser.msoffice.model.ParsedDocument;
import io.parser.msoffice.model.TextBlockData;
import io.parser.msoffice.parser.Parser;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.boot.json.JacksonJsonParser;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class PPTXParser implements Parser {

    public static void main(String[] args) {
        Parser parser = new PPTXParser();
        parser.parse(null);
    }

    @Override
    public ParsedDocument parse(File documentToParsing) {
        String filePath = "c:\\Users\\User\\Downloads\\SBI AUTOMOTIVE OPPORTUNITIES FUND - LEAFLET (1).pptx";

        try (FileInputStream fis = new FileInputStream(filePath);
             XMLSlideShow ppt = new XMLSlideShow(fis)) {

            ppt.getSlides().forEach(slide -> {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textBlock = (XSLFTextShape) shape;
//                        for (XSLFTextParagraph paragraph : textBlock.getTextParagraphs()) {
//                            StringBuilder builder = new StringBuilder();
//                            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
//                                builder.append(textRun.getRawText());
//                            }
//                            System.out.println(builder);
//                        }
                        ObjectMapper mapper = new ObjectMapper();
                        String json = null;
                        try {
                            json = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(makeTextBlockData(textBlock));
                        } catch (JsonProcessingException e) {
                            throw new RuntimeException(e);
                        }
                        System.out.println(json);
                    }
                }
                System.out.println("-----------" + slide.getSlideNumber());
            });
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private Page makePage(XSLFSlide slide) {
        return null;
    }

    private TextBlockData makeTextBlockData(XSLFTextShape textShape) {
        StringBuilder builder = new StringBuilder();
        for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
            for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                builder.append(textRun.getRawText());
            }
        }
        return new TextBlockData(
                builder.toString(),
                null, //TODO: Make translation via AI
                textShape.getTextParagraphs().get(0).getBulletFontSize(),
                "",// toColor(textShape.getTextParagraphs().get(0).getBulletFontColor()).getRGB()
                getHex(textShape.getTextParagraphs().get(0).getTextRuns().get(0).getFontColor()),
                getFontStyle(textShape.getTextParagraphs().get(0).getTextRuns().get(0)),
                getX(textShape),
                getY(textShape),
                getTopX(textShape),
                getTopY(textShape)
        );
    }

    private Color toColor(PaintStyle paintStyle) {
        Color fontColor = null;
        PaintStyle.SolidPaint solidPaint = (PaintStyle.SolidPaint) paintStyle;
        fontColor = solidPaint.getSolidColor().getColor();
        return fontColor;
    }

    private String getHex(PaintStyle paintStyle) {
        Color color = toColor(paintStyle);
        return String.format(
                "#%02x%02x%02x",
                color.getRed(),
                color.getGreen(),
                color.getBlue()
        );
    }

    private String getFontStyle(XSLFTextRun textRun) {
        StringBuilder builder = new StringBuilder();
        if (textRun.isBold()) builder.append("bold");
        if (textRun.isItalic()) builder.append("italic");
        if (textRun.isUnderlined()) builder.append("underline");
        if (textRun.isStrikethrough()) builder.append("strike");
        if (textRun.isSubscript()) builder.append("subscript");
        if (textRun.isSuperscript()) builder.append("superscript");
        return builder.toString();
    }

    private Double getX(XSLFTextShape shape) {
        return getBounds(shape).getX();
    }

    private Double getY(XSLFTextShape shape) {
        return getBounds(shape).getY();
    }

    private Double getTopX(XSLFTextShape shape) {
        double height = getBounds(shape).getHeight();
        return getBounds(shape).getX() + height;
    }

    private Double getTopY(XSLFTextShape shape) {
        double width = getBounds(shape).getWidth();
        return getBounds(shape).getY() + width;
    }

    private Rectangle2D getBounds(XSLFTextShape shape) {
        PlaceableShape<?, ?> placeableShape = (PlaceableShape<?, ?>) shape;
        Rectangle2D bounds = placeableShape.getAnchor();
        return bounds;
    }
}
